"""Batch 159: the small authenticated backend service that replaces a
direct Turso/database connection from inside the app. See FEATURE_LOG.md's
Batch 157 entry for why this exists (three separate Turso failures, a
"laggy and slow" report, and -- the deciding factor -- Yitzi's concern
that the app's update-distribution GitHub repo is public, so a database
secret baked into the installer could eventually be pulled out of it by
anyone). This service is the fix: it is the only thing that holds real
Postgres credentials (as the DATABASE_URL environment variable on
whatever host runs this file -- never in git, never in the app), and the
desktop app never talks to Postgres directly again. All it ships with is
this service's own address, which is not a secret -- knowing it buys you
nothing without a valid YJ PO Generator username and password.

Authentication deliberately is NOT a new, separate secret for users to
learn. It reuses the exact username/password accounts po_core.py's users
table already has (the ones already checked for permissions throughout
the app) -- POST /login here is the same check LoginDialog does locally
against SQLite today, just against the shared Postgres users table
instead, over HTTPS. A successful login gets an opaque session token
(secrets.token_urlsafe -- unguessable, not a JWT, not derived from
anything -- there's nothing to decode); every other endpoint requires it.

Endpoints:
    GET  /health         -- liveness + a real DB round-trip; used both by
                             Render's own health check and by the app's
                             "can I reach the shared database at all"
                             probe before it ever tries to log in.
    POST /login           {username, password}
                          -> 200 {token, user}
                          -> 401 wrong username/password, or inactive account
                          -> 403 detail="must_set_password" (brand new
                             account -- see set_user_permissions()/
                             create_user()'s must_set_password flag in
                             po_core.py; the client should show its
                             "choose your password" flow and call
                             /set_password instead)
    POST /set_password    {username, password, new_password}
                          -> 200 {token, user}  (also logs the user in,
                             same as /login, so one round trip covers a
                             first-time password choice)
                          -> 401 wrong current password (skipped only when
                             the account is still must_set_password)
                          -> 400 blank new_password
    POST /execute          {token, sql, params}
                          -> 200 {rows, lastrowid, rowcount}
                          -> 400 the query itself failed (bad SQL, a
                             constraint violation, etc. -- reported back
                             so the client's existing try/except around
                             conn.execute(...) still works)
                          -> 401 token not recognised / session expired
    POST /commit            {token} -> {status: "ok"}
    POST /rollback          {token} -> {status: "ok"}
    POST /logout            {token} -> {status: "ok"}  (always succeeds,
                             even for an already-gone token, since the
                             client's own cleanup should never fail on
                             this)

Session model: a session is one held-open psycopg2 connection plus which
user it belongs to, keyed by the opaque token, kept in an in-memory dict.
It has to be a real held-open connection (not one-shot-per-request) so a
po_core.py function that runs several conn.execute() calls before a single
conn.commit() (e.g. save_po()) still behaves as one transaction over what
is otherwise a stateless HTTP API -- see the client-side _RemoteConnection
class (Batch 160+) that will make this invisible to po_core.py entirely.

Idle sessions (no request in SESSION_IDLE_TIMEOUT_SECONDS) are swept on
every /login and /set_password call so a Render free-tier instance -- or a
Neon free-tier connection cap -- never quietly fills up with connections
an app that crashed or lost network never got to /logout.
"""
import os
import secrets
import sys
import threading
import time
from pathlib import Path
from typing import Optional, Union

import psycopg2
import psycopg2.extras
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel

# hash_password()/verify_password() are pure hashlib/secrets functions
# with no database or UI dependency (see po_core.py's own module
# docstring), so importing them here does not pull in sqlite3, PySide6,
# or anything Windows-only.
#
# Batch 167: po_core.py can legitimately be found in two different places
# depending on how this file is actually running, and both need to keep
# working. In this repository (and in verify_batch160.py/verify_batch164.py,
# which run this exact file as a real subprocess), app.py lives in server/
# with po_core.py one directory up. But SETUP_GUIDE.md's Part B has Yitzi
# create a brand new, separate GitHub repo containing only the handful of
# server/ files -- flat, with no server/ subfolder at all -- which is also
# exactly what Render then deploys. In that real deployed layout, po_core.py
# (now a required file in that same upload, see SETUP_GUIDE.md) sits right
# alongside app.py, not one level up. Rather than pick one layout and break
# the other, both directories go on sys.path; whichever one actually has
# po_core.py in it is the one Python uses.
_APP_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(_APP_DIR.parent))
sys.path.insert(0, str(_APP_DIR))
from po_core import hash_password, verify_password  # noqa: E402

from sql_translate import translate_sql  # noqa: E402

DATABASE_URL = os.environ.get(
    "DATABASE_URL",
    "dbname=pogen_dev user=pogen_dev password=pogen_dev_pw host=localhost",
)
SESSION_IDLE_TIMEOUT_SECONDS = 30 * 60

app = FastAPI(title="YJ PO Generator shared-database backend")

_sessions = {}
_sessions_lock = threading.Lock()


class LoginRequest(BaseModel):
    username: str
    password: str


class SetPasswordRequest(BaseModel):
    username: str
    password: Optional[str] = None
    new_password: str


class ExecuteRequest(BaseModel):
    token: str
    sql: str
    # A plain list/tuple for SQLite's '?' positional style (translated to
    # psycopg2's '%s'), or an object/dict for save_po()'s ':name' style
    # (translated to psycopg2's '%(name)s') -- see sql_translate.py's
    # _translate_placeholders() for exactly which of po_core.py's SQL
    # actually uses each style. Pydantic tries each member in order, so a
    # JSON array becomes a list and a JSON object becomes a dict, matching
    # whatever the client's _RemoteConnection.execute() actually sent.
    params: Union[list, dict] = []


class TokenRequest(BaseModel):
    token: str


def _new_connection():
    conn = psycopg2.connect(DATABASE_URL, cursor_factory=psycopg2.extras.RealDictCursor)
    conn.autocommit = False
    return conn


def _cleanup_idle_sessions():
    now = time.time()
    with _sessions_lock:
        stale = [t for t, s in _sessions.items() if now - s["last_used"] > SESSION_IDLE_TIMEOUT_SECONDS]
        for t in stale:
            try:
                _sessions[t]["conn"].close()
            except Exception:
                pass
            del _sessions[t]


def _get_session(token):
    with _sessions_lock:
        session = _sessions.get(token)
        if session is None:
            raise HTTPException(status_code=401, detail="Not logged in, or the session has expired. Please log in again.")
        session["last_used"] = time.time()
        return session


def _start_session(user_id, username, is_admin):
    token = secrets.token_urlsafe(32)
    with _sessions_lock:
        _sessions[token] = {
            "conn": _new_connection(),
            "user_id": user_id,
            "username": username,
            "is_admin": is_admin,
            "last_used": time.time(),
        }
    return token


def _lookup_user_row(cur, username):
    cur.execute(
        "SELECT id, username, full_name, password_hash, password_salt, is_admin, active, must_set_password, "
        "dashboard_widgets, notification_role "
        "FROM users WHERE lower(username) = lower(%s)",
        (username,),
    )
    return cur.fetchone()


def _user_permissions(cur, user_id):
    cur.execute("SELECT permission_key FROM user_permissions WHERE user_id=%s", (user_id,))
    return sorted(r["permission_key"] for r in cur.fetchall())


def _public_user(row, permissions):
    return {
        "id": row["id"],
        "username": row["username"],
        "full_name": row["full_name"],
        "is_admin": bool(row["is_admin"]),
        "permissions": permissions,
        # Batch 171: without these two, a real logged-in account's own
        # saved dashboard layout (core.user_dashboard_widgets()) and
        # notification-role filter (used to decide which toasts/badges
        # this account sees) silently fell back to the global default
        # every time -- harmless (never a crash, since both callers use
        # dict.get() with a sensible fallback), but a real, if minor,
        # loss of a personal setting for anyone using the hosted backend.
        "dashboard_widgets": row["dashboard_widgets"],
        "notification_role": row["notification_role"],
    }


@app.get("/health")
def health():
    try:
        conn = _new_connection()
        conn.cursor().execute("SELECT 1")
        conn.close()
        db_ok = True
    except Exception:
        db_ok = False
    return {"status": "ok", "database_reachable": db_ok}


@app.post("/bootstrap_migrate_token")
def bootstrap_migrate_token():
    """Batch 168: the one-time local-to-shared-database migration
    (see po_core.py's precheck_backend_migration()/
    migrate_one_table_to_backend()) needs an authenticated session to
    call /execute with, same as everything else here -- but the very
    first migration against a brand new database runs while `users` is
    still completely empty, before any ordinary /login is even possible.
    Creating a throwaway real account just to log in as would leave a
    spurious row behind, or collide with the real `users` table being
    copied across row-for-row by that same migration a few requests
    later.

    Issues a session token tied to no real user row at all -- an
    in-memory session only, same _start_session() every real login
    already uses, just with no matching users row behind it -- but ONLY
    while `SELECT COUNT(*) FROM users` is genuinely zero. The moment even
    one real user exists (which migrating the users table itself
    produces, among other ways), this permanently refuses again, on this
    or any future call, so it can never be used to bypass authentication
    on a database that's actually in use. Nothing here checks WHICH
    tables are empty, only `users` -- po_core.py's own
    precheck_backend_migration() is what actually guarantees every table
    is empty before copying a single row; this endpoint only needs to
    guarantee it can't hand out access once real accounts exist."""
    conn = _new_connection()
    try:
        cur = conn.cursor()
        try:
            cur.execute("SELECT COUNT(*) AS c FROM users")
            count = cur.fetchone()["c"]
        except Exception as e:
            # Most likely cause, by far: schema_postgres.sql was never
            # actually run in Neon's SQL Editor (SETUP_GUIDE.md's Part A,
            # step 4), so there's no `users` table -- or any other table --
            # to count rows in at all. /health's own database_reachable
            # check only proves Postgres itself answers, not that the
            # schema exists, so this is the first place that gap would
            # ever actually surface. Surfaced here, directly, rather than
            # as a bare 500 -- the previous version of this endpoint did
            # exactly that, which is what silently sent a real person
            # (Yitzi) into a login screen that could never work instead of
            # telling him what was actually wrong.
            raise HTTPException(
                status_code=500,
                detail=(
                    "Couldn't read the database's tables. The most likely reason: "
                    "schema_postgres.sql hasn't been run yet in Neon's SQL Editor "
                    "(SETUP_GUIDE.md's Part A, step 4) -- do that first, then try again. "
                    f"({e})"
                ),
            )
    finally:
        conn.close()
    if count:
        raise HTTPException(
            status_code=403,
            detail=(
                "This database already has real user accounts on it -- bootstrap access is "
                "only for a brand new, completely empty database. Log in with a real account "
                "instead."
            ),
        )
    token = _start_session(0, "migration-bootstrap", True)
    return {"token": token}


@app.post("/login")
def login(req: LoginRequest):
    _cleanup_idle_sessions()
    conn = _new_connection()
    try:
        cur = conn.cursor()
        row = _lookup_user_row(cur, req.username)
        if row is None or not row["active"]:
            raise HTTPException(status_code=401, detail="Invalid username or password.")
        if row["must_set_password"]:
            raise HTTPException(status_code=403, detail="must_set_password")
        if not verify_password(req.password, row["password_salt"], row["password_hash"]):
            raise HTTPException(status_code=401, detail="Invalid username or password.")
        permissions = _user_permissions(cur, row["id"])
    finally:
        conn.close()

    token = _start_session(row["id"], row["username"], bool(row["is_admin"]))
    return {"token": token, "user": _public_user(row, permissions)}


@app.post("/set_password")
def set_password(req: SetPasswordRequest):
    _cleanup_idle_sessions()
    conn = _new_connection()
    try:
        cur = conn.cursor()
        row = _lookup_user_row(cur, req.username)
        if row is None or not row["active"]:
            raise HTTPException(status_code=401, detail="Invalid username or password.")
        if not row["must_set_password"]:
            if not req.password or not verify_password(req.password, row["password_salt"], row["password_hash"]):
                raise HTTPException(status_code=401, detail="Invalid username or password.")
        if not req.new_password:
            raise HTTPException(status_code=400, detail="A new password is required.")
        salt, pw_hash = hash_password(req.new_password)
        cur.execute(
            "UPDATE users SET password_hash=%s, password_salt=%s, must_set_password=0 WHERE id=%s",
            (pw_hash, salt, row["id"]),
        )
        conn.commit()
        permissions = _user_permissions(cur, row["id"])
        row = dict(row)
        row["must_set_password"] = 0
    finally:
        conn.close()

    token = _start_session(row["id"], row["username"], bool(row["is_admin"]))
    return {"token": token, "user": _public_user(row, permissions)}


@app.post("/execute")
def execute(req: ExecuteRequest):
    session = _get_session(req.token)
    conn = session["conn"]
    try:
        translated = translate_sql(req.sql)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    try:
        cur = conn.cursor()
        # Batch 170: psycopg2 treats ANY non-None second argument to
        # execute() -- including an empty list/tuple, which is exactly
        # what a parameterless po_core.py query sends -- as a signal to
        # do %-style substitution on the SQL text itself. A query with no
        # real placeholders can still contain a literal '%' character
        # (several do, e.g. "WHERE po_ref LIKE 'PO:%'"), and psycopg2
        # then tries to substitute against an empty argument list for
        # that '%', raising a confusing "list index out of range" that
        # has nothing to do with the query's actual logic. sqlite3 itself
        # never has this problem (its placeholder is '?', not '%s', so a
        # literal '%' is never special) -- only pass params through here
        # when there actually are some, so a parameterless query's SQL
        # text is sent to Postgres completely unmodified, same as sqlite3
        # already treats it.
        if req.params:
            cur.execute(translated, req.params)
        else:
            cur.execute(translated)
        rows = []
        lastrowid = None
        if cur.description is not None:
            rows = [dict(r) for r in cur.fetchall()]
            if len(rows) == 1 and "id" in rows[0] and translated.strip().upper().startswith("INSERT"):
                lastrowid = rows[0]["id"]
        rowcount = cur.rowcount
    except HTTPException:
        raise
    except Exception as e:
        # Batch 175: unlike sqlite3, a Postgres connection with autocommit
        # off (see _new_connection() above) marks its WHOLE transaction
        # "aborted" the instant any single statement inside it fails --
        # every later command on that same connection then fails too, with
        # a generic "current transaction is aborted, commands ignored
        # until end of transaction block" error that has nothing to do
        # with whatever query actually runs next. sqlite3 has no such
        # behavior at all (a failed statement never poisons later ones),
        # which is the assumption every one of po_core.py's ~700 read and
        # write call sites was written and tested against -- in
        # particular, an ordinary SELECT (the overwhelming majority of
        # them) has no rollback of its own, because under SQLite there's
        # never anything to roll back after a read failing. Yitzi hit this
        # for real: one genuine query bug (report_pos_needing_followup's
        # HAVING clause, fixed separately in po_core.py) cascaded into the
        # notification check, the launch backup, and every other query on
        # that same session failing with "current transaction is aborted"
        # for the rest of that session's life, until the app was closed
        # and reopened (a fresh login starts a brand new connection).
        # Rolling back HERE, immediately, resets Postgres's transaction
        # state so the very next query -- even a totally unrelated one
        # moments later -- starts clean instead of cascading, matching
        # what every existing caller already assumes a failed query does:
        # nothing lingers, whatever went wrong is now this one query's
        # problem alone. Best-effort (the connection could already be in
        # a state where even rollback() itself raises, e.g. if the
        # network dropped mid-statement) -- either way, the ORIGINAL
        # error below is what actually matters to report back.
        try:
            conn.rollback()
        except Exception:
            pass
        raise HTTPException(status_code=400, detail=f"query failed: {e}")
    return {"rows": rows, "lastrowid": lastrowid, "rowcount": rowcount}


@app.post("/commit")
def commit(req: TokenRequest):
    session = _get_session(req.token)
    session["conn"].commit()
    return {"status": "ok"}


@app.post("/rollback")
def rollback(req: TokenRequest):
    session = _get_session(req.token)
    session["conn"].rollback()
    return {"status": "ok"}


@app.post("/logout")
def logout(req: TokenRequest):
    with _sessions_lock:
        session = _sessions.pop(req.token, None)
    if session is not None:
        try:
            session["conn"].close()
        except Exception:
            pass
    return {"status": "ok"}

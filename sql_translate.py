"""Batch 157+: the one place that translates po_core.py's own SQL text
(written, and tested, against SQLite) into the equivalent Postgres SQL,
so none of po_core.py's ~700 existing, thoroughly-tested query strings
need to change for the move off Turso. See FEATURE_LOG.md's Batch 157
entry and TECHNICAL_REBUILD_SPEC.md's matching "how" section for the full
audit this is based on -- this module covers exactly the handful of
SQLite idioms that audit found still needing translation at query time
(everything else -- AUTOINCREMENT, PRAGMA, sqlite_master -- is a
schema-management concern the backend's own independent Postgres schema
(schema_postgres.sql) makes unnecessary to ever translate on the fly).

translate_sql(sql) is the one function callers need. Order of operations
inside it matters less than it looks -- date-function rewriting only ever
touches text around literal `date(`/`datetime(`/`strftime(` calls, `?`
placeholder rewriting only touches bare `?` characters outside quoted
string literals, and RETURNING-id appending only looks at the statement's
overall INSERT/table shape -- but it's applied in a fixed order (date
functions, then placeholders, then RETURNING) so the same input always
produces the same output, which matters for testing.
"""
import re

# Every other table in schema_postgres.sql has a single-column integer `id`
# primary key -- these five don't, so a blanket "RETURNING id" must never
# be appended to an INSERT into one of them.
TABLES_WITHOUT_ID = {
    "settings",
    "zoho_vendor_map",
    "product_merge_ignored",
    "supplier_merge_ignored",
    "user_permissions",
}

_INSERT_TABLE_RE = re.compile(r"^\s*INSERT\s+INTO\s+([A-Za-z_][A-Za-z0-9_]*)", re.IGNORECASE)
_RETURNING_RE = re.compile(r"\bRETURNING\b", re.IGNORECASE)

# strftime('fmt', expr) -- expr in this codebase is always a bare column
# reference (with or without a table alias) or a '?' placeholder, never
# something itself containing parentheses, so a non-greedy match up to the
# first ')' is safe and correct for every real call site.
_STRFTIME_RE = re.compile(r"strftime\(\s*'([^']*)'\s*,\s*([^()]+?)\s*\)", re.IGNORECASE)
_DATE_RE = re.compile(r"\bdate\(\s*([^()]+?)\s*\)", re.IGNORECASE)
_DATETIME_RE = re.compile(r"\bdatetime\(\s*([^()]+?)\s*\)", re.IGNORECASE)

# The only four format strings report_spend_over_time()/similar reporting
# queries ever actually pass to strftime() (see that function's own `fmt =
# {"day": ..., "week": ..., "year": ...}.get(granularity, "%Y-%m")` line) --
# deliberately a closed, hand-verified list rather than a general SQLite
# strftime-format interpreter, since a general one is real work this
# codebase doesn't need. translate_sql() raises loudly on anything else
# here rather than silently emitting wrong SQL, so a future new format
# string gets caught in testing, not in front of Yitzi.
#
# The week case is the one honest approximation: SQLite's %W is "week of
# year, Monday-based, 00-53"; Postgres's IW is the ISO-8601 week number,
# which can differ by one for the first/last few days of a year. Close
# enough for a weekly trend chart's bucketing, not byte-identical in every
# edge case -- flagged here rather than assumed away.
_STRFTIME_FORMAT_MAP = {
    "%Y-%m-%d": "YYYY-MM-DD",
    "%Y-%m": "YYYY-MM",
    "%Y": "YYYY",
    "%Y-W%W": 'IYYY-"W"IW',
}


def _translate_date_functions(sql):
    def strftime_repl(m):
        fmt, expr = m.group(1), m.group(2)
        pg_fmt = _STRFTIME_FORMAT_MAP.get(fmt)
        if pg_fmt is None:
            raise ValueError(
                f"sql_translate: no Postgres translation registered for strftime format {fmt!r} "
                f"-- add it to _STRFTIME_FORMAT_MAP after confirming the right to_char() equivalent"
            )
        return f"to_char(({expr})::timestamp, '{pg_fmt}')"

    sql = _STRFTIME_RE.sub(strftime_repl, sql)
    sql = _DATE_RE.sub(lambda m: f"({m.group(1)})::date", sql)
    sql = _DATETIME_RE.sub(lambda m: f"({m.group(1)})::timestamp", sql)
    return sql


_NAMED_PARAM_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_]*")


def _translate_placeholders(sql):
    """SQLite's two placeholder styles -> psycopg2's own equivalents,
    never touching one that's actually inside a single-quoted string
    literal (e.g. a notes field containing a literal '?' or ':'). po_core.py's
    SQL text never uses double-quoted string literals (double quotes are
    reserved for identifiers), so only single-quote state needs tracking;
    SQL's own escape for a quote inside a string is doubling it ('') -- a
    plain toggle-per-quote-character handles that correctly too, since two
    togglings in a row cancel out, leaving the state unchanged.

    - a bare '?' -> '%s' -- the overwhelming majority of po_core.py's
      conn.execute(sql, (...)) call sites.
    - ':name' -> '%(name)s' -- save_po()'s two statements are the only
      call sites in po_core.py using SQLite's OTHER placeholder style
      (a dict of values instead of a tuple); psycopg2 accepts a dict
      parameter natively via this exact '%(name)s' pyformat syntax.
    - a literal '::' is left completely alone -- _translate_date_functions()
      (run before this, see translate_sql() below) already introduced
      Postgres type casts like "(created_at)::timestamp" into the SQL
      text, and a second, unadorned ':' immediately following the first
      must never be mistaken for the start of a ':name' placeholder."""
    out = []
    in_string = False
    i, n = 0, len(sql)
    while i < n:
        ch = sql[i]
        if ch == "'":
            in_string = not in_string
            out.append(ch)
            i += 1
        elif in_string:
            out.append(ch)
            i += 1
        elif ch == "?":
            out.append("%s")
            i += 1
        elif ch == ":" and sql[i + 1:i + 2] == ":":
            out.append("::")  # a Postgres type cast, not a named placeholder
            i += 2
        elif ch == ":":
            m = _NAMED_PARAM_RE.match(sql, i + 1)
            if m:
                out.append(f"%({m.group(0)})s")
                i = m.end()
            else:
                out.append(ch)
                i += 1
        else:
            out.append(ch)
            i += 1
    return "".join(out)


def _add_returning_id(sql):
    """Postgres's psycopg2 has no cursor.lastrowid -- the backend reads the
    new row's id back via RETURNING instead, and reports it to the client
    as lastrowid so none of po_core.py's ~7 `cur.lastrowid` call sites need
    to change. Only appended to a plain INSERT, only when the statement
    doesn't already have its own RETURNING clause, and never for the five
    tables in TABLES_WITHOUT_ID (no `id` column for it to return)."""
    stripped = sql.strip()
    if stripped[:6].upper() != "INSERT":
        return sql
    if _RETURNING_RE.search(sql):
        return sql
    m = _INSERT_TABLE_RE.match(sql)
    if not m:
        return sql
    if m.group(1).lower() in TABLES_WITHOUT_ID:
        return sql
    return stripped.rstrip(";") + " RETURNING id"


def translate_sql(sql):
    sql = _translate_date_functions(sql)
    sql = _translate_placeholders(sql)
    sql = _add_returning_id(sql)
    return sql

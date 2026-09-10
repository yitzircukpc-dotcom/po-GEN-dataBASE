"""
YJ - Rose Communications Group Ltd PO Generator — core business/data layer
============================================================================
Version 3.9.14

This module has NO UI framework dependency at all (no Tkinter, no Qt). It is
the tested, working guts of the app: the SQLite schema and data access, the
one-time migration from the old v1 JSON files, PO business logic (parsing,
validation, email/PDF content building), the dependency-free PDF writer, the
Windows clipboard/Outlook integration, and backup/restore/export/import.

Both the classic Tkinter UI (po_generator_v2.py) and the new PySide6 UI
(po_generator_qt.py) import everything they need from here, unchanged. None
of this file changed when the UI was rebuilt — it was already tested against
real production data (81 migrated purchase orders, 17 suppliers, 4 addresses)
and there was no reason to touch code that works.

Sections:
    1. Paths / constants / logging
    2. Database layer (schema, CRUD)
    3. Migration from v1 (loose JSON files)
    4. Business logic — formatting, parsing, validation, email builders
    5. Clipboard / Outlook integration (Windows only)
    6. PDF writer (dependency-free)
    7. Backup / restore / export / import
"""

import csv  # noqa: F401  (available for callers doing CSV import/export, e.g. suppliers/reports)
import difflib
import hashlib
import json
import logging
import os
import re
import secrets
import shutil
import sqlite3
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import zipfile
import zlib
from datetime import datetime, timedelta
from pathlib import Path

# Batch 109: optional -- only needed once someone actually turns on the
# shared/hosted database (see load_turso_config below). A build that hasn't
# added the 'turso' package yet (or a computer where it failed to install)
# simply can't use that feature; everything else in the app is completely
# unaffected, since every other line in this file only ever talks to
# whatever `conn` it's handed via plain .execute()/.commit() calls, never
# sqlite3 directly.
#
# Batch 123: `import turso` on its own does NOT make `_turso.sync` usable --
# Yitzi hit "Couldn't connect: module 'turso' has no attribute 'sync'" the
# moment he actually had a working pyturso install (0.7.2) to test against.
# turso.sync is a real, separate submodule (turso/sync/__init__.py, with its
# own connect()/ConnectionSync -- confirmed by installing 0.7.2 and checking
# directly), but Python only attaches a submodule onto its parent package as
# an attribute once that submodule has actually been imported somewhere --
# plain `import turso` never touches it, so every `_turso.sync.connect(...)`
# call below raised this exact AttributeError on a clean install. The
# explicit `import turso.sync` line is what registers it.
try:
    import turso as _turso
    import turso.sync  # noqa: F401 -- registers _turso.sync (see comment above)
except Exception:
    _turso = None

# ============================================================
# 1. Paths / constants / logging
# ============================================================

APP_TITLE = "YJ - Rose Communications Group Ltd PO Generator"
APP_VERSION = "3.9.14"

# Batch 104: a separate counter from APP_VERSION, bumped only when a change
# actually alters what's expected to be in the database (a new required
# column, a new table something else now depends on, a changed meaning for
# an existing value) -- not on every release. Stored inside the database
# file itself (SQLite's built-in PRAGMA user_version) so it travels with the
# data no matter which computer opens it. See the version check inside
# get_connection() below: this is what lets an old copy of the app refuse to
# touch a database a newer copy has already upgraded, instead of silently
# misreading or corrupting it once several computers share one database.
APP_SCHEMA_VERSION = 1

# Batch 105/111: where the app looks for over-the-air update information --
# a small JSON file (see check_for_update() below) telling it the latest
# version, an optional minimum required version, and where to download it.
# Left blank, update checking is simply off; nothing else in the app
# depends on this. Points at manifest.json in the public
# yitzircukpc-dotcom/POGEN-UPDATES GitHub repo Yitzi set up (see
# SETUP_GUIDE.md) -- a plain raw.githubusercontent.com URL, reachable with
# no login/auth, which is what urllib.request.urlopen() in
# check_for_update() actually needs (there's no credential support built
# into that call). Safe to leave pointed here even before the file exists
# or before any build has actually shipped: check_for_update() never
# raises on a 404/unreachable host, it just returns None (no update
# available), same as if this were still blank.
UPDATE_MANIFEST_URL = "https://raw.githubusercontent.com/yitzircukpc-dotcom/POGEN-UPDATES/main/manifest.json"

APP_DIR = Path(os.path.expanduser("~")) / ".yj_po_generator"
APP_DIR.mkdir(parents=True, exist_ok=True)

DB_PATH = APP_DIR / "data.db"
LOG_PATH = APP_DIR / "app.log"
# Batch 109: the shared/hosted database's URL + auth token, when Yitzi has
# set one up (see SETUP_GUIDE.md and load_turso_config below) -- its own
# small file, deliberately outside both the database and the versioned
# source, since it's needed to even open the database and must never be
# copied into a git repo or a zipped build.
TURSO_CONFIG_PATH = APP_DIR / "turso_config.json"
BACKUP_DIR = APP_DIR / "backups"
BACKUP_DIR.mkdir(parents=True, exist_ok=True)
ZOHO_EXPORT_DIR = APP_DIR / "zoho_exports"
ZOHO_EXPORT_DIR.mkdir(parents=True, exist_ok=True)

# v1 legacy files this version migrates from, once, automatically.
LEGACY_SETTINGS_FILE = Path(os.path.expanduser("~")) / ".yj_po_generator_settings.json"
LEGACY_HISTORY_FILE = Path(os.path.expanduser("~")) / ".yj_po_generator_history.json"

BG = "#f3f0ea"
BLUE = "#cfe8f3"
RED = "#d40000"
BLACK = "#000000"

# ---- RCG brand palette for the redesigned emails (batch 85+) -- the same
# navy/blue/mauve/green/red already used throughout the app's own UI
# (ACCENT/BRAND_MAUVE/GREEN/RED in po_generator_qt.py) and the real RCG
# logo, so branded emails read as the same product as the app itself, not
# an invented new look. See section 2/14 of Yitzi's change request and the
# approved design sample (design_samples/rcg_email_redesign_sample.html). ----
# Batch 93: Yitzi sent the three exact brand hexes to use everywhere --
# "make sure you use these exact colors to replace all the branded colors
# you use across the system" -- these are the logo's own three droplet
# colors (see rcg_logo_full_transparent.png), so RCG_INK/RCG_ACCENT both
# now point at the same dark ink (#2a206f, replacing the old #241c4a/
# #4b728e guesses) and RCG_ACCENT_LIGHT/RCG_MAUVE are unchanged since they
# already matched exactly. No other hex should be introduced anywhere
# branded is drawn -- every PDF/email surface should trace back to one of
# these three.
RCG_INK = "#2a206f"
RCG_ACCENT = "#2a206f"
RCG_ACCENT_LIGHT = "#6498be"
RCG_MAUVE = "#b079ad"
RCG_GREEN = "#1e9e6b"
RCG_RED = "#d64545"
RCG_ROW_BLUE = "#eef4f9"
RCG_ROW_MAUVE = "#f8f1f7"
RCG_LINE = "#ece9f5"
RCG_MUTED = "#6b7280"
RCG_LOGO_CID = "rcg_logo"

CURRENCY_SYMBOLS = {"GBP": "£", "USD": "$", "EUR": "€"}

STATUS_CHOICES = ["Draft", "Sent", "Confirmed", "Received", "Cancelled"]

DEFAULT_SETTINGS = {
    "business_name": "ROSE COMMUNICATIONS GROUP LTD",
    "company_number": "4468350",
    "currency": "GBP",
    "po_prefix": "YJ",
    "invoice_name": "ROSE COMMUNICATIONS GROUP LTD",
    "invoice_address": "92-94 STAMFORD HILL\nLONDON\nN16 6XS",
    "bcc_emails": "Stock@rcuk.com\nAccounts@rcgroup.co.uk",
    "your_name": "",
    # Every Outlook draft this app opens (PO emails, price requests, the
    # periodic report, the Zoho export reminder) is sent from a shared
    # mailbox rather than your own -- set here once so it's applied
    # automatically instead of having to remember to switch the From
    # account by hand every time. Leave blank to use Outlook's own default
    # account instead.
    "outlook_shared_mailbox": "supply@rcgroup.co.uk",
    "auto_backup_on_launch": "1",
    "auto_backup_keep": "20",
    # Zoho Books integration: the status newly-imported historical POs are
    # given (most users only ever move a PO to "Sent" in this app and don't
    # track Received/Confirmed here, so this is a setting rather than fixed).
    "zoho_import_default_status": "Sent",
    # The "Purchase Order Status" value written into exported CSVs, i.e. what
    # status the order lands in inside Zoho Books after you upload it there.
    "zoho_export_default_status": "Draft",
    # When a Zoho import (POs or products) last actually ran -- shown in the
    # Import / Export page so it's obvious how stale the data might be if
    # imports aren't run for a while.
    "zoho_last_po_import_at": "",
    "zoho_last_item_import_at": "",
    "zoho_last_supplier_import_at": "",
    # Periodic PO summary report email (opens a ready-to-send Outlook draft
    # once the scheduled day/time arrives -- this app never sends on its own).
    # Weekly/monthly/quarterly/yearly are all independent schedules that can
    # be enabled in any combination at once (Batch 99: "it needs to remind
    # me for the monthly quater and yearly aswell" -- before this, only ONE
    # frequency was ever "the" schedule, so having weekly on meant monthly/
    # quarterly/yearly could never fire at all, however overdue they were).
    "po_report_enabled": "0",  # master switch -- off silences every frequency below at once
    "po_report_weekly_enabled": "0",
    "po_report_weekly_day": "0",  # weekday, 0=Monday
    "po_report_monthly_enabled": "0",
    "po_report_monthly_day": "1",  # day of the month
    "po_report_quarterly_enabled": "0",
    "po_report_quarterly_day": "1",  # day of the quarter's first month
    "po_report_yearly_enabled": "0",
    "po_report_yearly_day": "1",  # day of January
    "po_report_time": "09:00",
    # "this" = the full current week/month/quarter/year (even if not
    # finished yet); "last" = the most recently completed one instead.
    # Shared across all four frequencies -- one preference about what
    # content to include, not a per-schedule thing.
    "po_report_window": "this",
    "po_report_to": "Maxi.rose@rcgroup.co.uk",
    "po_report_cc": "",
    # First name (or however you'd like them addressed) used for the "Hi
    # ___," greeting at the top of the periodic report email. Leave blank
    # for a plain "Hi,". Multiple people can still be on po_report_to/cc --
    # this is just who the email is addressed to.
    "po_report_recipient_name": "Maxi",
    "po_report_last_sent_weekly": "",
    "po_report_last_sent_monthly": "",
    "po_report_last_sent_quarterly": "",
    "po_report_last_sent_yearly": "",
    # Legacy single-frequency settings -- no longer read anywhere except the
    # one-time _migrate_report_frequency_to_multi migration below, which
    # needs somewhere safe to read an old install's prior single choice
    # from. Never written to going forward; kept only for that migration.
    "po_report_frequency": "weekly",
    "po_report_day": "0",
    "po_report_last_sent_at": "",
    # Weekly stock team recap email -- separate from the periodic PO summary
    # report above. Goes to the stock team only, lists every order placed in
    # the period covered with its line items (normal turnaround is a couple
    # of days, so by the time this goes out most of it should already be in),
    # and asks them to flag anything missing/not arrived yet or any supplier
    # issues, returns, or RMAs. Same "opens a ready-to-send Outlook draft,
    # never sends on its own" pattern as the report above, on its own
    # independent schedule/recipients.
    "stock_recap_enabled": "0",
    "stock_recap_day": "0",  # weekday, 0=Monday
    "stock_recap_time": "09:00",
    # "this" = the current week so far; "last" = the most recently completed
    # week (Monday-Sunday) -- "last" is the natural fit for a Monday-morning
    # recap of what was ordered the week before.
    "stock_recap_window": "last",
    "stock_recap_to": "Stock@rcgroup.co.uk",
    "stock_recap_cc": "",
    "stock_recap_recipient_name": "Tommy",
    "stock_recap_subject": "Stock recap - orders placed w/c {date}",
    "stock_recap_prompt": (
        "Please have a look through the above and reply to this email if anything is "
        "missing, hasn't arrived yet, or if there's a supplier issue, return, or RMA to "
        "flag on any of it."
    ),
    "stock_recap_last_sent_at": "",
    # Batch 108: the reverse-direction report -- the stock team telling
    # Yitzi what hasn't arrived yet, instead of Yitzi telling the stock team
    # what was ordered. Triggered manually from a button on the Stock page
    # (stock.send_report permission), not on a schedule like the recap
    # above, so there's no enabled/day/time/window here -- just who it goes
    # to and the subject line.
    "stock_outstanding_report_to": "",
    "stock_outstanding_report_cc": "",
    "stock_outstanding_report_recipient_name": "",
    "stock_outstanding_report_subject": "Outstanding orders -- what hasn't arrived yet ({date})",
    "stock_outstanding_report_last_sent_at": "",
    # Batch 121: the stock team's own weekly "please order this" request to
    # Supply -- Yitzi: "emails from stock side send from Stock@rcgroup.co.uk
    # to Supply@ i want these all to be editable in settings aswell as
    # where the suplly ones send from" (the "supply ones" already have
    # their own From address -- outlook_shared_mailbox, above). Default day/
    # time match his own instruction ("defult should be set for a monday
    # morning at 9:30 am email"), with the same enabled/day/time/"send now"
    # override shape as stock_recap_* above -- see stock_request_is_due().
    #
    # Batch 151 correction: stock_request_enabled originally defaulted to
    # "1" here, seeded on ahead of the actual weekly-email feature (Batch
    # 146 built only the Settings card; nothing read this setting yet).
    # Now that stock_request_is_due() actually gates a real startup
    # "hasn't been sent yet -- open it now?" prompt (same shape as
    # po_report_enabled/stock_recap_enabled's own startup checks), leaving
    # this one alone defaulting to "1" would mean it's the only scheduled
    # email in the whole app that starts firing on a brand new install
    # with nobody having opted in -- every sibling schedule
    # (po_report_enabled, stock_recap_enabled) defaults to "0" specifically
    # so a fresh business doesn't start sending emails nobody asked for
    # yet. Flipped to match. The day/time defaults stay exactly what Yitzi
    # asked for (Monday 9:30am) -- only the on/off default changed, so
    # turning it on in Settings still lands on the schedule he wanted.
    "stock_request_from": "Stock@rcgroup.co.uk",
    "stock_request_to": "Supply@rcgroup.co.uk",
    "stock_request_cc": "",
    "stock_request_enabled": "0",
    "stock_request_day": "0",  # weekday, 0=Monday
    "stock_request_time": "09:30",
    "stock_request_subject": "Stock requests -- w/c {date}",
    "stock_request_last_sent_at": "",
    # Custom PDF branding: an alternate logo shown on generated PO PDFs
    # (top-right of page 1) and a standard terms/payment-terms/footer block
    # appended after the addresses on every PO PDF. A blank pdf_logo_path
    # means "use the bundled Rose Communications Group logo" (see
    # _effective_pdf_logo_path in po_generator_qt.py) unless pdf_logo_disabled
    # is set, in which case no logo is drawn at all.
    "pdf_logo_path": "",
    "pdf_logo_disabled": "0",
    "pdf_terms_text": "",
    # Scheduled automatic backup, independent of the safety backups already
    # taken before a Zoho import or a duplicate merge.
    "auto_backup_schedule_enabled": "0",
    "auto_backup_schedule_frequency": "weekly",  # currently only "weekly"
    "auto_backup_schedule_day": "0",  # weekday 0=Monday
    "auto_backup_last_scheduled_at": "",
    # Optional second copy of every backup, dropped into a folder the user
    # points at (e.g. a OneDrive/SharePoint-synced folder, or a network
    # share) -- off by default so nothing tries to write anywhere unexpected
    # until it's deliberately turned on and a folder is chosen.
    "secondary_backup_enabled": "0",
    "secondary_backup_folder": "",
    "secondary_backup_last_at": "",
    "secondary_backup_last_error": "",
    # Stock team receiving sync (Batch 96, stock_sync.py) -- Phase 1,
    # read-only: writes a snapshot of outstanding POs into a shared folder
    # for the separate stock companion program to read. Off by default,
    # same reasoning as the secondary backup above. These are just the
    # default settings values -- all the actual logic lives in
    # stock_sync.py, not here, so this feature stays cleanly removable.
    "stock_sync_enabled": "0",
    "stock_sync_folder": "",
    "stock_sync_last_at": "",
    "stock_sync_last_error": "",
    # Batch 97: per-filename line offsets (JSON-encoded dict) into each
    # companion device's own receiving-updates file, so a re-read only
    # looks at lines added since last time. See stock_sync.py.
    "stock_sync_event_offsets": "",
    # How long a backup is kept before it's eligible for pruning, applied to
    # both the local backups folder and the secondary/cloud folder. This is
    # a MINIMUM, not a target -- backups are only ever deleted once they're
    # older than this, never on a count basis, so there's no risk of a busy
    # week (e.g. the twice-daily scheduled backups below) crowding out
    # something older that's still within the window.
    "auto_backup_retention_days": "60",
    # Twice-daily weekday (Mon-Fri, 10:00 and 17:00) automatic backups via a
    # Windows Scheduled Task, mirroring reminder_tasks_registered_for below
    # -- registered once per exe path, not on every launch.
    "backup_tasks_registered_for": "",
    # Proactive price-change alerts: warn when adding a line item at a price
    # that differs from the product's last recorded price.
    "price_alert_enabled": "1",
    # Default period for the Dashboard's "Spend this period"/"Spend last
    # period" KPI cards -- "week" (Monday-Sunday) or "month" (calendar
    # month). The Dashboard also has its own on-screen selector that can
    # temporarily switch this per-view without changing the saved default.
    "dashboard_period": "week",
    # PO reference format: the DDMMYY date portion right after the prefix is
    # kept fixed (parse_date_from_po_ref and the Zoho date-preference/backfill
    # logic all depend on recognizing it), but what comes after the dash is
    # configurable -- "time" (HHMM, the original behaviour) or "sequence"
    # (001, 002, ... counting today's POs, so two POs made in the same
    # minute never collide).
    "po_ref_suffix_style": "time",
    # Wording used in individual PO emails (the periodic PO summary report
    # has its own separate template and isn't affected by these).
    "email_intro_line": "Please see below order",
    "email_outro_line": "Thank you",
    # A Sent PO with no activity (resent/copied/opened/marked followed-up)
    # for this many days shows up on the Dashboard's follow-up nudge list.
    "chase_followup_days": "5",
    # Price request emails ("what's your best price on X?") -- sent
    # individually to each chosen supplier from the Price Requests page.
    # Placeholders substituted at send time: {supplier_name}, {product},
    # {qty}, {your_name}.
    "price_request_email_subject": "Pricing enquiry: {product}",
    "price_request_email_body": (
        "Hi {supplier_name},\n\n"
        "Do you have the following in stock, and if so what would be your best price?\n\n"
        "Product: {product}\n"
        "Quantity: {qty}\n\n"
        "Please let me know at your earliest convenience.\n\n"
        "Thank you"
    ),
    # ---- Daily Zoho export reminder ----
    # The accounts team need every Sent PO in their Zoho account too, so the
    # Dashboard keeps a running count of Sent POs that haven't been exported
    # yet (however many days it's been -- nothing is dropped if a day is
    # missed) and a button that builds an Excel workbook and opens a ready
    # -to-send Outlook draft with it attached. Placeholders in the subject/
    # body: {count}, {date}.
    # Master on/off switch -- when "0" the Dashboard banner is hidden and
    # the 9:30 popup reminder (and the "reports/follow-ups outstanding"
    # check it also does) skips the Zoho part entirely.
    "zoho_export_enabled": "1",
    "zoho_export_recipient_to": "",
    "zoho_export_recipient_cc": "",
    "zoho_export_email_subject": "Purchase orders for Zoho import - {date}",
    "zoho_export_email_body": (
        "Hi Accounts team,\n\n"
        "Attached are {count_label} ready to import into Zoho Books ({date}).\n\n"
        "Thanks\n"
        "Procurement Team"
    ),
    "zoho_export_last_sent_at": "",
    # Time of day the Dashboard reminder (and, on Windows, the background
    # popup registered in Task Scheduler) checks in -- "09:30" by default.
    "zoho_reminder_check_time": "09:30",
    # Records what this install last registered with Windows Task Scheduler
    # (exe path) so it only re-registers when that's actually changed, e.g.
    # after an update -- not on every single launch.
    "reminder_tasks_registered_for": "",
    # ---- Zoho export column defaults ----
    # Zoho Books' own Purchase Order CSV has far more columns than this app
    # tracks per PO (cost centres, tax settings, ship-to details, and so
    # on) -- these are genuine accounting/business settings, not something
    # this app can guess correctly per order, so every one of them is a
    # plain editable default here (Settings > Zoho Export) rather than
    # hard-coded. Left blank means "leave that column blank on export".
    "zoho_col_vat_treatment": "uk",
    "zoho_col_is_inclusive_tax": "false",
    "zoho_col_exchange_rate": "1.000000000000",
    "zoho_col_template_name": "",
    "zoho_col_delivery_instructions": "",
    "zoho_col_terms_conditions": "",
    "zoho_col_shipment_preference": "",
    "zoho_col_account": "",
    "zoho_col_account_code": "",
    "zoho_col_usage_unit": "",
    "zoho_col_discount_type": "entity_level",
    "zoho_col_is_discount_before_tax": "true",
    "zoho_col_tax_id": "",
    "zoho_col_item_tax": "Standard Rate",
    "zoho_col_item_tax_pct": "20.00",
    "zoho_col_item_tax_type": "ItemAmount",
    "zoho_col_item_exemption_code": "",
    # A line item ticked "Margin VAT" on the PO (VAT-inclusive price, nothing
    # to reclaim -- e.g. second-hand/margin-scheme stock) exports with these
    # two values instead of the normal zoho_col_item_tax/_pct above, on that
    # line only. Zoho Books' UK edition doesn't ship a single standard preset
    # for this (its tax rates are set up per-account, and unlike some other
    # regional editions it has no built-in "Profit Margin Scheme" checkbox),
    # so this defaults to a plain 0% "No VAT" rate -- rename it here to match
    # whatever a 0%/no-VAT rate is actually called in your own Zoho account.
    "zoho_margin_vat_item_tax": "No VAT",
    "zoho_margin_vat_item_tax_pct": "0.00",
    "zoho_col_item_type": "goods",
    "zoho_col_acq_vat_name": "",
    "zoho_col_acq_vat_pct": "",
    "zoho_col_adjustment_description": "Adjustment",
    "zoho_col_discount_account": "",
    "zoho_col_discount_account_code": "",
    "zoho_col_cost_centre": "",
    "zoho_col_cost_allocation": "",
    "zoho_col_group_cost_allocation": "",
    "zoho_col_profit_centre": "",
    "zoho_col_location": "",
    "zoho_col_marketing_tag": "",
    "zoho_col_person": "",
    "zoho_col_budgets": "",
    "zoho_col_project_id": "",
    "zoho_col_project_name": "",
    "zoho_col_payment_terms": "",
    "zoho_col_payment_terms_label": "",
    "zoho_col_attention": "",
    "zoho_col_address": "92-94 Stamford Hill",
    "zoho_col_city": "London",
    "zoho_col_state": "",
    "zoho_col_country": "United Kingdom",
    "zoho_col_postcode": "N16 6XS",
    "zoho_col_phone": "",
    "zoho_col_deliver_to_customer": "",
    "zoho_col_purchase_owner": "",
    # ---- Procurement savings tracking ----
    # "YTD" respects a configurable financial year start month rather than
    # always meaning the plain calendar year -- 1=January (the calendar
    # year) by default, set higher here if the real financial year starts
    # some other month.
    "savings_fiscal_year_start_month": "1",
    "savings_dashboard_tile_enabled": "1",
    # Gentle, dismissible prompt offering to log a saving right when a PO is
    # sent -- that's the moment the final price is locked in, per Yitzi, so
    # it's the most natural time to capture it rather than reconstructing it
    # later. Purely optional, never blocks sending.
    "savings_nudge_enabled": "1",
    # ---- Supplier scorecards ----
    # How far back the "price competitiveness" comparison looks when
    # working out whether this supplier's prices beat everyone else's on
    # the same products.
    "scorecard_price_window_months": "12",
    # ---- Supplier credit limits ----
    # Off by default -- Zoho's PO export has no live payment/paid status
    # (only a billing status), so "used" can only ever be an approximation
    # (a running total since a manual reset), not a real balance. Left for
    # Yitzi to switch on only if that approximation is actually useful.
    "credit_limit_tracking_enabled": "0",
    # ---- Supplier naming consistency ----
    # Comma-separated words that "Clean up naming" (Suppliers page) always
    # renders fully uppercase rather than title-casing -- eg. "RVT" so a
    # supplier entered as "rvt" gets suggested as "RVT" rather than "Rvt".
    # Seeded once on first run from whatever's already written in all caps
    # in the supplier list, plus "RVT" itself (see
    # _migrate_seed_supplier_name_acronyms) -- editable afterwards from the
    # cleanup tool itself.
    "supplier_name_acronyms": "",
    "supplier_name_acronyms_seeded": "0",
    # ---- Dashboard widget visibility ----
    # Every Dashboard widget can be shown or hidden from Settings >
    # Dashboard -- see DASHBOARD_WIDGETS below (the YTD savings tile is
    # part of that same list too, reusing its own older
    # savings_dashboard_tile_enabled setting above rather than a new
    # dash_show_* one). All on by default so nothing changes for anyone
    # who's never touched this.
    "dash_show_kpi_total_pos": "1",
    "dash_show_kpi_active_pos": "1",
    "dash_show_kpi_spend_this": "1",
    "dash_show_kpi_spend_last": "1",
    "dash_show_kpi_suppliers": "1",
    "dash_show_kpi_products": "1",
    "dash_show_kpi_alerts": "1",
    "dash_show_kpi_price_requests": "1",
    "dash_show_recent_pos": "1",
    "dash_show_price_watch": "1",
    "dash_show_followup": "1",
    "dash_show_flagged": "1",
    "dash_show_price_requests_card": "1",
    "dash_show_kpi_open_stock_requests": "1",
    # Batch 117: the four new Dashboard tiles below, plus their own
    # threshold setting. On by default same as every widget added before
    # them, so an existing install's Dashboard just gains four extra tiles
    # rather than needing anyone to opt in.
    "dash_show_kpi_orders_due_in": "1",
    "dash_show_kpi_orders_overdue": "1",
    "dash_show_kpi_outstanding_value": "1",
    "dash_show_kpi_avg_order_value": "1",
    # A not-yet-fully-received PO counts as "overdue" once it's been this
    # many days since it was sent (see stock_sync.split_due_overdue) --
    # Yitzi asked for "orders due in"/"orders overdue" tiles specifically
    # "with a configurable overdue threshold." Configurable from Settings >
    # Dashboard, alongside the widget visibility checkboxes.
    "po_overdue_days": "7",
}


# Single source of truth for which Dashboard widgets can be shown/hidden
# from Settings > Dashboard, and what QWidget attribute on DashboardPage
# each one maps to -- shared by the Settings tab (which builds one checkbox
# per entry), the per-account picker on PermissionsDialog, and
# DashboardPage._refresh (which just loops this list and calls setVisible
# on the matching attribute), so none of the three can ever drift out of
# sync with each other.
#
# "permission" (Batch 123) is the viewing permission that widget's own
# figures actually come from -- Yitzi gave a non-admin ("Tommy") access to
# just Settings > Dashboard so he could manage the org's default tile
# layout, and found it listed every tile to tick, including ones tied to
# pages Tommy himself has no permission to open at all ("it shows him all
# the tiles to tick or untick even those that are not available to him").
# Both the Settings > Dashboard tab and PermissionsDialog's per-account
# picker now only offer a widget to tick when the relevant account (the
# person opening Settings for the former, the account being configured for
# the latter) actually holds this permission -- see
# core.user_has_permission. None means every dashboard.view-holding account
# can already see it (it's not gated behind any narrower page permission).
DASHBOARD_WIDGETS = [
    {"attr": "kpi_total_pos", "setting": "dash_show_kpi_total_pos", "label": "Total purchase orders", "permission": "po_manager.view"},
    {"attr": "kpi_active_pos", "setting": "dash_show_kpi_active_pos", "label": "Active (not deleted)", "permission": "po_manager.view"},
    {"attr": "kpi_this_month", "setting": "dash_show_kpi_spend_this", "label": "Spend this period", "permission": "po_manager.view"},
    {"attr": "kpi_last_month", "setting": "dash_show_kpi_spend_last", "label": "Spend last period", "permission": "po_manager.view"},
    {"attr": "kpi_suppliers", "setting": "dash_show_kpi_suppliers", "label": "Active suppliers", "permission": "suppliers.view"},
    {"attr": "kpi_products", "setting": "dash_show_kpi_products", "label": "Products in catalog", "permission": "products.view"},
    {"attr": "kpi_alerts", "setting": "dash_show_kpi_alerts", "label": "Price alerts", "permission": "reports.view"},
    {"attr": "kpi_price_requests", "setting": "dash_show_kpi_price_requests", "label": "Price requests awaiting replies", "permission": "price_requests.view"},
    {"attr": "recent_card", "setting": "dash_show_recent_pos", "label": "Recent purchase orders", "permission": "po_manager.view"},
    {"attr": "alerts_card", "setting": "dash_show_price_watch", "label": "Price watch (products that have gone up)", "permission": "reports.view"},
    {"attr": "followup_card", "setting": "dash_show_followup", "label": "Needs follow-up", "permission": "po_manager.view"},
    {"attr": "flagged_card", "setting": "dash_show_flagged", "label": "Flagged for review", "permission": "po_manager.view"},
    {"attr": "price_req_card", "setting": "dash_show_price_requests_card", "label": "Price requests waiting on replies", "permission": "price_requests.view"},
    # Batch 117: Yitzi asked to "put together some more dashboard icons and
    # widgets that may be useful for all our departments" -- these four are
    # backed by stock_sync.split_due_overdue/outstanding_order_value (the
    # first two) and this module's own average_order_value (the last one),
    # all computed in DashboardPage._refresh alongside the existing KPIs.
    {"attr": "kpi_orders_due_in", "setting": "dash_show_kpi_orders_due_in", "label": "Orders due in", "permission": "po_manager.view"},
    {"attr": "kpi_orders_overdue", "setting": "dash_show_kpi_orders_overdue", "label": "Orders overdue", "permission": "po_manager.view"},
    {"attr": "kpi_outstanding_value", "setting": "dash_show_kpi_outstanding_value", "label": "Outstanding order value", "permission": "po_manager.view"},
    {"attr": "kpi_avg_order_value", "setting": "dash_show_kpi_avg_order_value", "label": "Average order value", "permission": "po_manager.view"},
    # Batch 123: folded in from its own older, standalone
    # savings_dashboard_tile_enabled setting (still the same setting key,
    # still also editable from the Savings & Scorecards tab -- both places
    # just toggle the one underlying value) so it participates in the same
    # per-account personal-widget-list and permission-visibility handling
    # as every other tile instead of being a hardcoded special case.
    {"attr": "kpi_savings_ytd", "setting": "savings_dashboard_tile_enabled", "label": "YTD savings", "permission": "savings.view"},
    # Batch 147: mainly useful to Supply (a running count of what Stock is
    # waiting on), but gated on stock_requests.view like the page itself
    # rather than a Supply-only key, so a Stock account that wants it on
    # their own dashboard (as a reminder of what they're still waiting on
    # themselves) can add it too.
    {"attr": "kpi_open_stock_requests", "setting": "dash_show_kpi_open_stock_requests", "label": "Open stock requests", "permission": "stock_requests.view"},
]

DASHBOARD_WIDGET_ATTRS = {w["attr"] for w in DASHBOARD_WIDGETS}


def user_dashboard_widgets(user):
    """Batch 119: parses a user dict's raw `dashboard_widgets` column (a
    JSON array of DASHBOARD_WIDGETS attr strings, or NULL/None) into a
    Python list, or None. None means "no personal layout -- fall back to
    the global Settings > Dashboard toggles," which is also what every
    account gets before this feature and what a bare MainWindow() with no
    login context always gets (user is None there, and this function
    returns None for that too). Tolerates a malformed/corrupted value
    (falls back to None rather than raising) since a bad value here should
    never be able to break someone's Dashboard outright -- worst case they
    just see the same widgets everyone else does.

    Unrecognised attrs (e.g. a widget that's since been removed from
    DASHBOARD_WIDGETS) are silently dropped rather than left in -- they'd
    just be inert anyway (DashboardPage._refresh only ever looks up attrs
    it actually has), but keeping the parsed list clean makes it safe for
    a caller to feed straight back into set_user_dashboard_widgets."""
    if not user:
        return None
    raw = user.get("dashboard_widgets")
    if not raw:
        return None
    try:
        parsed = json.loads(raw)
    except (TypeError, ValueError):
        return None
    if not isinstance(parsed, list):
        return None
    return [a for a in parsed if a in DASHBOARD_WIDGET_ATTRS]


def set_user_dashboard_widgets(conn, user_id, attrs):
    """Saves this account's personal Dashboard widget list -- attrs is a
    list of DASHBOARD_WIDGETS attr strings (only recognised ones are kept,
    same filtering as user_dashboard_widgets). Pass None specifically to
    clear the override entirely and go back to "same as everyone else" --
    an empty list [] is a deliberate, different thing (a real, saved
    customization that happens to show zero widgets), so it's kept as a
    real (empty) JSON array rather than being treated the same as None."""
    if attrs is None:
        value = None
    else:
        value = json.dumps([a for a in attrs if a in DASHBOARD_WIDGET_ATTRS])
    conn.execute("UPDATE users SET dashboard_widgets=? WHERE id=?", (value, user_id))
    conn.commit()


def setup_logging():
    logger = logging.getLogger("po_generator")
    logger.setLevel(logging.INFO)
    if logger.handlers:
        return logger
    handler = logging.FileHandler(LOG_PATH, encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    logger.addHandler(handler)
    return logger


log = setup_logging()


def now_iso():
    return datetime.now().isoformat(timespec="seconds")


# ============================================================
# 2. Database layer
# ============================================================

SCHEMA = """
CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT
);

CREATE TABLE IF NOT EXISTS addresses (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    label TEXT NOT NULL,
    company TEXT DEFAULT '',
    address TEXT DEFAULT '',
    sort_order INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS suppliers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    company_name TEXT NOT NULL,
    contact_name TEXT DEFAULT '',
    email TEXT DEFAULT '',
    cc_emails TEXT DEFAULT '',
    phone TEXT DEFAULT '',
    address TEXT DEFAULT '',
    active INTEGER DEFAULT 1,
    notes TEXT DEFAULT ''
);

CREATE TABLE IF NOT EXISTS products (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    supplier_id INTEGER NOT NULL DEFAULT 0,
    code TEXT DEFAULT '',
    name TEXT NOT NULL,
    last_price REAL DEFAULT 0,
    times_ordered INTEGER DEFAULT 0,
    updated_at TEXT,
    UNIQUE(supplier_id, name)
);

CREATE TABLE IF NOT EXISTS purchase_orders (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    po_ref TEXT UNIQUE NOT NULL,
    status TEXT DEFAULT 'Draft',
    supplier_id INTEGER,
    supplier_company_name TEXT DEFAULT '',
    supplier_contact_name TEXT DEFAULT '',
    supplier_email TEXT DEFAULT '',
    supplier_cc_emails TEXT DEFAULT '',
    supplier_phone TEXT DEFAULT '',
    supplier_address TEXT DEFAULT '',
    delivery_label TEXT DEFAULT '',
    delivery_name TEXT DEFAULT '',
    delivery_address TEXT DEFAULT '',
    business_name TEXT DEFAULT '',
    company_number TEXT DEFAULT '',
    currency TEXT DEFAULT 'GBP',
    invoice_name TEXT DEFAULT '',
    invoice_address TEXT DEFAULT '',
    your_name TEXT DEFAULT '',
    bcc_emails TEXT DEFAULT '',
    total REAL DEFAULT 0,
    notes TEXT DEFAULT '',
    deleted INTEGER DEFAULT 0,
    created_at TEXT,
    updated_at TEXT
);

CREATE TABLE IF NOT EXISTS po_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    po_id INTEGER NOT NULL REFERENCES purchase_orders(id) ON DELETE CASCADE,
    position INTEGER DEFAULT 0,
    qty INTEGER DEFAULT 1,
    code TEXT DEFAULT '',
    product TEXT DEFAULT '',
    price REAL DEFAULT 0,
    margin_vat INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS po_events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    po_id INTEGER NOT NULL REFERENCES purchase_orders(id) ON DELETE CASCADE,
    event TEXT NOT NULL,
    detail TEXT DEFAULT '',
    at TEXT
);

CREATE INDEX IF NOT EXISTS idx_po_ref ON purchase_orders(po_ref);
CREATE INDEX IF NOT EXISTS idx_po_status ON purchase_orders(status);
CREATE INDEX IF NOT EXISTS idx_po_deleted ON purchase_orders(deleted);
CREATE INDEX IF NOT EXISTS idx_po_supplier ON purchase_orders(supplier_company_name);
CREATE INDEX IF NOT EXISTS idx_items_po ON po_items(po_id);
CREATE INDEX IF NOT EXISTS idx_events_po ON po_events(po_id);
CREATE INDEX IF NOT EXISTS idx_products_name ON products(name);
CREATE INDEX IF NOT EXISTS idx_products_supplier ON products(supplier_id);

CREATE TABLE IF NOT EXISTS product_merge_ignored (
    product_id_a INTEGER,
    product_id_b INTEGER,
    PRIMARY KEY(product_id_a, product_id_b)
);

CREATE TABLE IF NOT EXISTS supplier_merge_ignored (
    supplier_id_a INTEGER,
    supplier_id_b INTEGER,
    PRIMARY KEY(supplier_id_a, supplier_id_b)
);

CREATE TABLE IF NOT EXISTS zoho_vendor_map (
    zoho_vendor_name TEXT PRIMARY KEY,
    supplier_id INTEGER NOT NULL
);

CREATE TABLE IF NOT EXISTS supplier_issues (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    supplier_company_name TEXT NOT NULL,
    note TEXT NOT NULL,
    at TEXT
);
CREATE INDEX IF NOT EXISTS idx_supplier_issues_name ON supplier_issues(supplier_company_name);

-- Price requests ("what's your best price on X?" sent to several suppliers
-- at once) -- see section 10 below for the functions that use these.
CREATE TABLE IF NOT EXISTS price_requests (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    product_name TEXT NOT NULL,
    qty INTEGER DEFAULT 1,
    notes TEXT DEFAULT '',
    created_at TEXT
);

CREATE TABLE IF NOT EXISTS price_request_replies (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    request_id INTEGER NOT NULL REFERENCES price_requests(id) ON DELETE CASCADE,
    supplier_id INTEGER,
    supplier_company_name TEXT DEFAULT '',
    status TEXT DEFAULT 'pending',
    price REAL,
    notes TEXT DEFAULT '',
    sent_at TEXT,
    replied_at TEXT
);
CREATE INDEX IF NOT EXISTS idx_price_reply_request ON price_request_replies(request_id);
CREATE INDEX IF NOT EXISTS idx_price_requests_product ON price_requests(product_name);

-- Procurement savings tracking (section 11 below) -- a saving (or, with a
-- negative amount, a price increase) recorded against a PO, a product, or
-- both, so the YTD figure is built up as you go rather than reconstructed
-- at year end. Categories are a user-editable list rather than a fixed
-- enum, same idea as supplier_issue_categories below.
CREATE TABLE IF NOT EXISTS savings_categories (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE,
    sort_order INTEGER DEFAULT 0,
    active INTEGER DEFAULT 1
);

CREATE TABLE IF NOT EXISTS savings (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    category TEXT NOT NULL DEFAULT '',
    po_ref TEXT DEFAULT '',
    product TEXT DEFAULT '',
    supplier_company_name TEXT DEFAULT '',
    qty REAL DEFAULT 0,
    previous_price REAL,
    new_price REAL,
    amount REAL NOT NULL DEFAULT 0,
    note TEXT DEFAULT '',
    created_at TEXT,
    deleted INTEGER DEFAULT 0
);
CREATE INDEX IF NOT EXISTS idx_savings_created ON savings(created_at);
CREATE INDEX IF NOT EXISTS idx_savings_product ON savings(product);
CREATE INDEX IF NOT EXISTS idx_savings_po_ref ON savings(po_ref);
CREATE INDEX IF NOT EXISTS idx_savings_supplier ON savings(supplier_company_name);

-- User-editable categories for the supplier issues log (section 9b), so
-- the scorecard can show "3 late deliveries this year" instead of just a
-- pile of free-text notes.
CREATE TABLE IF NOT EXISTS supplier_issue_categories (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE,
    sort_order INTEGER DEFAULT 0,
    active INTEGER DEFAULT 1
);

-- Maps a PO reference's own prefix letters (e.g. "YJ", "MR") to the staff
-- member who raised it, for purchase-owner tracking (section 12 below).
-- Used whenever Zoho's own per-PO "purchase owner" field isn't available
-- -- which is always true for POs raised directly in this app, since this
-- app has never carried that field itself.
CREATE TABLE IF NOT EXISTS purchase_owner_prefixes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    prefix TEXT NOT NULL UNIQUE,
    name TEXT NOT NULL,
    sort_order INTEGER DEFAULT 0
);

-- Batch 97: receiving status events from the stock team's companion
-- programme (see stock_sync.py). Additive-only, part of the same
-- cleanly-removable stock sync subsystem as the Batch 96 snapshot writer
-- -- if ever abandoned, this table just sits unused or gets dropped in
-- one step, without touching purchase_orders/po_items at all.
-- source_event_id is globally unique (each companion install stamps its
-- own device id into every event it writes), which is what makes
-- re-applying the same file safe to do more than once -- INSERT OR
-- IGNORE on this column is how idempotent replay actually works.
CREATE TABLE IF NOT EXISTS po_receiving_events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    source_event_id TEXT NOT NULL UNIQUE,
    po_id INTEGER NOT NULL REFERENCES purchase_orders(id) ON DELETE CASCADE,
    po_item_id INTEGER NOT NULL REFERENCES po_items(id) ON DELETE CASCADE,
    status TEXT NOT NULL,
    qty REAL,
    note TEXT DEFAULT '',
    device_id TEXT DEFAULT '',
    reported_at TEXT,
    applied_at TEXT
);
CREATE INDEX IF NOT EXISTS idx_receiving_events_item ON po_receiving_events(po_item_id);
CREATE INDEX IF NOT EXISTS idx_receiving_events_po ON po_receiving_events(po_id);

-- Batch 103: user accounts, login, and per-user permissions. is_admin is a
-- separate bypass flag rather than just "every permission ticked" -- an
-- admin automatically has every permission that exists NOW and every one
-- added in a future batch, with nothing to re-tick each time a new
-- permission key is introduced, and it means Yitzi's own account can never
-- be accidentally locked out of a feature by an unticked box. Everyone
-- else's access is exactly whatever's listed in their own permissions row,
-- nothing implied. password_hash/password_salt use PBKDF2 (stdlib
-- hashlib, no third-party crypto dependency) -- see hash_password() below.
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT NOT NULL UNIQUE,
    full_name TEXT NOT NULL DEFAULT '',
    password_hash TEXT NOT NULL,
    password_salt TEXT NOT NULL,
    is_admin INTEGER DEFAULT 0,
    active INTEGER DEFAULT 1,
    -- Batch 114: 1 means this account has no password anyone actually
    -- knows yet (a brand new account, or one an admin just reset) -- the
    -- stored password_hash is a random, unusable placeholder nobody could
    -- practically guess (see create_user()/reset_user_password() below),
    -- and the login screen routes this account to "create your password"
    -- instead of an ordinary password check. Cleared back to 0 the moment
    -- a real password is actually set (set_user_password()), whoever set
    -- it -- the person themselves at first login, or an admin picking one
    -- directly.
    must_set_password INTEGER NOT NULL DEFAULT 0,
    created_at TEXT,
    updated_at TEXT
);
CREATE TABLE IF NOT EXISTS user_permissions (
    user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    permission_key TEXT NOT NULL,
    PRIMARY KEY (user_id, permission_key)
);
-- Batch 121: in-app notifications (bell icon). One row per recipient per
-- event -- a request that goes to three Supply-role accounts makes three
-- rows, each independently read/unread, so one person marking theirs read
-- never affects anyone else's copy. link_page/link_id (both optional) let
-- clicking a notification jump straight to the thing it's about (e.g.
-- link_page="stock_requests", link_id=the request id) -- see
-- MainWindow._open_notification. windows_shown tracks whether THIS row has
-- already triggered a native Windows toast on whichever machine happened to
-- notice it via pop_due_toasts() -- see that function's docstring for why
-- toasts are popped locally per running instance rather than at creation time.
CREATE TABLE IF NOT EXISTS notifications (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    type_key TEXT NOT NULL,
    title TEXT NOT NULL,
    body TEXT DEFAULT '',
    link_page TEXT DEFAULT '',
    link_id TEXT DEFAULT '',
    created_at TEXT,
    read_at TEXT DEFAULT '',
    windows_shown INTEGER DEFAULT 0
);
CREATE INDEX IF NOT EXISTS idx_notifications_user ON notifications(user_id, read_at);
-- Batch 145: the stock-to-supply product request feature itself -- the
-- NOTIFICATION_TYPES rows and stock_request_* settings for it were seeded
-- back in Batch 121, waiting for this table. One row per request, of two
-- kinds: "existing" (Stock asking Supply to order more of a product
-- that's already in the catalogue -- product_id points at it) or "new"
-- (Stock flagging a product that ISN'T in the catalogue yet -- product_id
-- stays 0 until Supply either matches it to an existing product or
-- approves it as genuinely new; see the dedupe/approval workflow still to
-- come). product_name is always stored directly on the row rather than
-- looked up fresh through product_id every time -- for an "existing"
-- request this is just a copy of the catalogue name at the moment it was
-- asked for, so a later rename or deletion of that product can't quietly
-- alter what a historical request says it was for; for a "new" request
-- it's simply the only name there is, since there's no catalogue row yet.
-- status moves open -> ordered -> resolved (or -> cancelled from open),
-- resolved covering both "Supply linked/approved a flagged new item" and,
-- eventually, "the reorder was added to a PO" -- po_id records which PO,
-- once that's built. resolved_by/resolved_note capture how a "new" kind
-- request was resolved (matched to an existing product, or approved as a
-- new one) for the Stock side's own "Supply linked or approved a flagged
-- new item" notification to have something concrete to say.
CREATE TABLE IF NOT EXISTS product_requests (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    kind TEXT NOT NULL DEFAULT 'existing',
    product_id INTEGER DEFAULT 0,
    product_name TEXT NOT NULL,
    supplier_id INTEGER DEFAULT 0,
    qty INTEGER DEFAULT 0,
    notes TEXT DEFAULT '',
    requested_by INTEGER NOT NULL REFERENCES users(id),
    status TEXT NOT NULL DEFAULT 'open',
    created_at TEXT,
    ordered_at TEXT DEFAULT '',
    po_id INTEGER DEFAULT 0,
    resolved_at TEXT DEFAULT '',
    resolved_by INTEGER DEFAULT 0,
    resolved_note TEXT DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_product_requests_status ON product_requests(status);
"""

_connection = None


def load_turso_config():
    """Reads the local, per-computer shared-database configuration Yitzi
    pastes in via Settings > Data & Backup once he's created a Turso
    database (see SETUP_GUIDE.md). Returns {"url", "auth_token"} or None --
    never raises -- if it's never been set up, or the file is missing,
    unreadable, or malformed. Deliberately its own file rather than a
    settings-table row: you need this to even open the database, so it
    can't live inside the thing it unlocks."""
    if not TURSO_CONFIG_PATH.exists():
        return None
    try:
        data = json.loads(TURSO_CONFIG_PATH.read_text(encoding="utf-8"))
        url = str(data.get("url") or "").strip()
        token = str(data.get("auth_token") or "").strip()
        if not url or not token:
            return None
        return {"url": url, "auth_token": token}
    except Exception:
        log.exception("Couldn't read turso_config.json -- treating shared database as not configured")
        return None


def save_turso_config(url, auth_token):
    """Saves the shared-database connection details, entered by hand in
    Settings -- never hardcoded anywhere in this source file. Takes effect
    on the next get_connection() call, i.e. the next app launch (the
    already-open connection for this run isn't swapped out live)."""
    url = (url or "").strip()
    auth_token = (auth_token or "").strip()
    if not url:
        raise ValueError("The database URL is required.")
    if not auth_token:
        raise ValueError("The auth token is required.")
    TURSO_CONFIG_PATH.write_text(json.dumps({"url": url, "auth_token": auth_token}), encoding="utf-8")


def clear_turso_config():
    """Switches back to local-only mode on the next launch. Does NOT touch
    DB_PATH itself -- whatever's in the local replica file stays exactly as
    it is, so nothing already saved is lost; it just stops being kept in
    sync with the shared database going forward."""
    try:
        TURSO_CONFIG_PATH.unlink()
    except FileNotFoundError:
        pass


# Batch 160+: the new backend-service shared-database mode (see FEATURE_LOG.md's
# Batch 157/159/160 entries) alongside the Turso one above, not replacing it
# yet -- Turso stays fully intact and untouched until the backend model is
# proven in real use (see TECHNICAL_REBUILD_SPEC.md's Batch 157 entry, "what's
# still ahead"). The two are deliberately NOT symmetrical: Turso's config file
# holds a real secret (an auth token that alone grants full access) because
# that's simply how Turso works; this file holds only the backend's own
# address, which is not a secret at all -- knowing it buys nothing without a
# real per-user login. Nothing this app ever persists to disk for the backend
# path can be used to reach the shared database on its own, which is the
# entire point of building it this way (see Batch 157's public-repo security
# discussion).
BACKEND_CONFIG_PATH = APP_DIR / "backend_config.json"


def load_backend_config():
    """Reads the local, per-computer backend address Yitzi enters in
    Settings > Data & Backup once the backend service is deployed (see
    SETUP_GUIDE.md). Returns {"base_url", "last_username"} or None -- never
    raises -- if it's never been set up, or the file is missing, unreadable,
    or malformed. "last_username" is optional and purely a convenience (the
    login dialog pre-fills it) -- never a password, never a token, nothing
    that alone opens the database."""
    if not BACKEND_CONFIG_PATH.exists():
        return None
    try:
        data = json.loads(BACKEND_CONFIG_PATH.read_text(encoding="utf-8"))
        base_url = str(data.get("base_url") or "").strip()
        if not base_url:
            return None
        return {"base_url": base_url, "last_username": str(data.get("last_username") or "").strip()}
    except Exception:
        log.exception("Couldn't read backend_config.json -- treating the shared-database backend as not configured")
        return None


def save_backend_config(base_url, last_username=None):
    """Saves the backend's own address -- never a secret, see this section's
    own opening comment above. Takes effect on the next login (the already-
    open connection for this run, if any, isn't swapped out live -- same
    rule Turso's save_turso_config() already follows)."""
    base_url = (base_url or "").strip()
    if not base_url:
        raise ValueError("The shared-database service address is required.")
    existing = load_backend_config() or {}
    username = (last_username or "").strip() or existing.get("last_username", "")
    BACKEND_CONFIG_PATH.write_text(json.dumps({"base_url": base_url, "last_username": username}), encoding="utf-8")


def clear_backend_config():
    """Switches back to local-only mode (or back to Turso, if that's still
    separately configured) on the next launch. Does not touch DB_PATH."""
    try:
        BACKEND_CONFIG_PATH.unlink()
    except FileNotFoundError:
        pass


# Populated by the login UI (not by get_connection() itself -- po_core.py has
# no UI dependency and can't show a password prompt) calling
# set_remote_session() right after a successful remote_backend_login(), before
# the very first get_connection() call of this process. get_connection()'s own
# backend branch (below) simply uses whatever's here; if nothing's been set
# yet it raises SharedDatabaseUnreachableError telling the caller a login is
# needed first, rather than silently doing nothing or crashing on a None.
_remote_session = None


def set_remote_session(base_url, token):
    global _remote_session
    _remote_session = {"base_url": base_url, "token": token}


def clear_remote_session():
    """Called on log out, and by get_connection() itself when a remote
    session's token turns out to be no longer valid -- either way, the next
    get_connection() call must not try to reuse a session that's known to be
    dead; it should raise its own clear "please log in again" error instead."""
    global _remote_session
    _remote_session = None


# Batch 149: Yitzi first hit a shared-database sync failure at work that
# hadn't happened at home, which looked at the time like a network/firewall
# cause -- but he then reproduced the exact same failure on a phone hotspot
# too, on both home and work Wi-Fi, ruling that out directly ("on hot spot
# im getting this still i dont think its to do with fire wall"). A follow-up
# test settled it for real: deleting turso_config.json (forcing a plain
# local-only launch), then re-entering the same URL/token in Settings (which
# calls setup_shared_database() -- a brand new local replica file, pulled
# fresh from the remote) connected fine, synced a new PO both ways -- and
# then failed with this exact error again about two minutes later, while the
# app was already open and otherwise working, on a periodic background pull
# (see pull_latest()), not on startup at all. That rules out network,
# firewall, and "stale local file" all at once: this was a freshly rebuilt
# replica, already proven to work, failing on its own a couple of minutes
# in. It also ruled out the original Batch 140 theory (an in-process
# "checkpoint already in progress" guard) just as clearly, since this was a
# background operation on an already-successful connection, not a repeat
# attempt after an earlier failure in the same process.
#
# The CheckpointResult text itself changed shape too: earlier occurrences
# showed every counter at 0/false (no work attempted at all); this one
# showed wal_max_frame=66 and wal_total_backfilled=66 -- real progress --
# with wal_checkpoint_backfilled still 0 and every *_sent flag still false.
# That reads as the checkpoint step doing real work and then failing to
# finish, not refusing to start -- consistent with a genuine bug in the
# underlying sync engine's checkpoint state machine rather than anything
# this app controls (BUILD_INSTRUCTIONS.md installs it unpinned via `pip
# install pyturso`, a young, actively-developed Rust-backed package still
# on 0.x releases; the project's own GitHub issue tracker has several open
# reports of embedded-replica sync/page-consistency bugs in this same
# area). See _open_database_or_exit()'s docstring in po_generator_qt.py for
# what this batch actually does about it -- the automatic self-heal
# relaunch now forces a genuinely fresh local replica pull (the same fix
# that worked for Yitzi manually), not just a fresh process against the
# same possibly-wedged local file.
#
# _diagnose_turso_error() adds a short, best-effort category on top of the
# real underlying error text -- never replacing it, only prefixing it, so
# nothing is ever hidden behind a vaguer message. The category comes from
# pattern-matching common phrases in str(e) (this app has no visibility
# into pyturso/libsql's actual exception class hierarchy -- the same
# string-based approach _looks_like_stuck_sync_engine_checkpoint() already
# uses for one specific known error shape, generalized here to a handful
# of broader buckets: can't reach the host at all, an auth/token problem,
# or the remote server itself reporting an error). It's a heuristic, not a
# guarantee -- an error that doesn't match any known phrase just gets the
# raw text with no added guess, same as before this batch. The checkpoint
# error itself deliberately isn't one of these categories: it isn't a
# network, auth, or remote-server problem, and mislabelling it as one would
# send someone re-checking their internet connection or re-typing a token
# that was never the issue.
_TURSO_ERROR_CATEGORIES = [
    # (category label, hint shown to the person, phrases to look for in str(e).lower())
    (
        "network",
        "This looks like a network connectivity problem -- this computer couldn't reach the "
        "shared database's server at all. Check the internet connection, and whether a firewall, "
        "proxy, or VPN on this network is blocking outbound connections{host_clause}.",
        (
            "timed out", "timeout", "name or service not known", "nodename nor servname",
            "failed to resolve", "connection refused", "network is unreachable",
            "could not connect", "temporary failure in name resolution", "getaddrinfo",
            "no route to host", "connection reset", "unreachable",
        ),
    ),
    (
        "auth",
        "This looks like an authentication problem -- the saved auth token may be invalid, "
        "expired, or revoked. Re-enter the database URL and a fresh auth token in Settings.",
        ("unauthorized", "401", "403", "invalid token", "auth", "permission denied", "forbidden"),
    ),
    (
        "server",
        "The shared database's own server reported an error on its end -- not something wrong "
        "with this computer's setup. This is usually temporary; try again in a minute.",
        ("500", "502", "503", "504", "internal server error", "bad gateway", "service unavailable"),
    ),
]


def _diagnose_turso_error(operation, url, e):
    """Logs a detailed line to app.log (operation, host -- never the auth
    token, exception type, full text, and a traceback) and returns a
    human-readable message: a category hint (see _TURSO_ERROR_CATEGORIES
    above) if the error text matches one, prepended to the real underlying
    error -- or just the real underlying error alone if nothing matched.
    Every SharedDatabaseUnreachableError raised anywhere in this module
    routes through here rather than building its own ad hoc f-string, so
    logging and categorization can't drift out of sync between call sites."""
    try:
        host = urllib.parse.urlparse(url or "").hostname or "(unknown host)"
    except Exception:
        host = "(unknown host)"
    log.error(
        "Turso %s failed -- host=%s exception=%s: %s",
        operation, host, type(e).__name__, e, exc_info=True,
    )
    text = str(e).lower()
    for _category, hint, phrases in _TURSO_ERROR_CATEGORIES:
        if any(p in text for p in phrases):
            host_clause = f" to {host}" if host != "(unknown host)" else ""
            return f"{hint.format(host_clause=host_clause)}\n\nFull error: {e}"
    return str(e)


def test_turso_connection(url, auth_token, timeout=None):
    """A real, live connectivity check -- what Settings' "Test connection"
    button calls. Deliberately never touches DB_PATH or the app's actual
    connection: connects a brand new, throwaway embedded replica in a temp
    folder, tries a pull and a trivial query, then discards the whole
    thing. Returns (True, message) on success or (False, message) on any
    failure (bad URL, bad/expired token, unreachable host, package not
    installed) -- never raises, so a bad paste can't crash the dialog.

    Batch 138: the pull uses _pull_with_retry() (see its docstring), same
    as setup_shared_database() and get_connection() -- a transient sync
    engine hiccup shouldn't make "Test connection" flicker false when the
    URL/token/network are all genuinely fine.

    Batch 149: two more things, directly from Yitzi's original diagnostic
    request -- whether the failure text can be categorized (now routes
    through _diagnose_turso_error(), same as every other Turso failure in
    this file, so this gets a detailed app.log line and the same
    network/auth/server hints instead of a bare "Couldn't connect: {e}"),
    and whether tables the app actually expects are missing (a SELECT 1
    only proves the connection and sync engine work at all -- it says
    nothing about whether this remote actually holds a real, migrated copy
    of this app's schema, e.g. someone pointed it at a genuinely empty or
    unrelated database). After a successful pull, this now also runs a
    real read against the users table -- present in SCHEMA from the very
    first shared-database batch, so its absence means something more
    fundamental than a missing column -- and folds a failure there into
    the same (False, message) result rather than reporting a bare
    "Connected successfully" that overstates what was actually checked."""
    if _turso is None:
        return False, "The 'turso' package isn't installed in this build. See BUILD_INSTRUCTIONS.md."
    url = (url or "").strip()
    auth_token = (auth_token or "").strip()
    if not url or not auth_token:
        return False, "Enter both the database URL and the auth token first."
    import tempfile
    tmp_dir = tempfile.mkdtemp(prefix="yj_turso_test_")
    try:
        test_path = str(Path(tmp_dir) / "test_replica.db")
        db = _turso.sync.connect(test_path, remote_url=url, auth_token=auth_token)
        _pull_with_retry(db)
        db.execute("SELECT 1")
        try:
            db.execute("SELECT COUNT(*) FROM users").fetchone()
        except Exception:
            return False, (
                "Connected and synced, but this remote database doesn't look like a "
                "real YJ PO Generator database yet -- the 'users' table the app "
                "expects isn't there. If this is meant to be brand new, that's "
                "expected (it gets created the first time this computer switches to "
                "it); if it's meant to already have data, double-check the URL."
            )
        return True, "Connected successfully."
    except Exception as e:
        return False, f"Couldn't connect: {_diagnose_turso_error('test connection', url, e)}"
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


class _PushOnCommitConnection:
    """Wraps a turso.sync embedded-replica connection so every existing
    conn.commit() call across this file (there are hundreds, none of which
    should need to change for this feature) transparently also pushes to
    the shared database. A push failure (offline, server unreachable, a
    momentary network hiccup) is logged but never raised -- by the time
    push() runs, commit() has already succeeded against the LOCAL replica,
    so nobody's work is ever lost; it simply catches up on the next
    successful push. Everything else (execute, executescript, row_factory,
    close, ...) passes straight through to the real connection via
    __getattr__/__setattr__, so this wrapper is invisible to the other
    ~10,000 lines of this file.

    Batch 110: unlike a startup connection failure (which now blocks
    opening the app entirely -- see get_connection()), a push failure here
    happens to a session that was already open and already confirmed
    synced. Yitzi's call on this specific case (mid-session, not startup):
    don't interrupt whoever's already working over a momentary blip, but
    never let that go unnoticed either -- so this tracks whether the most
    recent push actually succeeded (sync_ok/sync_last_error), which the UI
    polls to show a persistent "not synced" warning banner (see
    MainWindow's sync status check) until a later commit successfully
    pushes again."""

    def __init__(self, inner):
        object.__setattr__(self, "_inner", inner)
        object.__setattr__(self, "sync_ok", True)
        object.__setattr__(self, "sync_last_error", "")
        # Batch 150: how many pushes/pulls in a row have failed, reset to 0
        # by either one succeeding. On its own this changes nothing -- it's
        # read by shared_database_sync_looks_stuck() below, which is what
        # actually decides when to stop treating this as "a normal blip"
        # and offer a real fix instead. See that function's own docstring
        # for why this exists.
        object.__setattr__(self, "consecutive_sync_failures", 0)

    def commit(self):
        self._inner.commit()
        try:
            self._inner.push()
            object.__setattr__(self, "sync_ok", True)
            object.__setattr__(self, "sync_last_error", "")
            object.__setattr__(self, "consecutive_sync_failures", 0)
        except Exception as e:
            object.__setattr__(self, "sync_ok", False)
            object.__setattr__(self, "sync_last_error", str(e))
            object.__setattr__(self, "consecutive_sync_failures", self.consecutive_sync_failures + 1)
            log.warning("Push to the shared database failed -- saved locally, will retry on the next commit", exc_info=True)

    def __getattr__(self, name):
        return getattr(self._inner, name)

    def __setattr__(self, name, value):
        setattr(self._inner, name, value)


class MustSetPasswordError(Exception):
    """Raised by remote_backend_login() (below) when the backend reports
    the account still needs its first real password chosen -- the same
    must_set_password state create_user() already gives a brand new
    account locally (see create_user()'s own docstring). The caller (the
    shared-database login flow -- see task tracking for the still-pending
    UI rework) should show the same "choose your password" screen it
    already shows for a local account in this state, then call
    remote_backend_set_password() instead of retrying the login."""

    def __init__(self, username):
        self.username = username
        super().__init__(f"{username} needs to choose a password before logging in.")


def _remote_post_json(base_url, path, payload, timeout=60):
    """POSTs a JSON body to the shared-database backend (server/app.py) and
    returns (status_code, parsed_json_body). Stdlib-only (urllib.request),
    matching how this file already talks to the OTA update-manifest URL
    (see check_for_update() above) -- no new packaged dependency for
    something this small. Raises SharedDatabaseUnreachableError for a
    genuine transport failure (host unreachable, timed out, TLS error,
    a response that isn't valid JSON at all) -- as opposed to a normal
    rejected request (a 401/403/400 with a real JSON body), which is
    returned to the caller to interpret, exactly like a real HTTP client
    library would.

    Batch 172: default bumped from 15 to 60 -- Render's own free tier
    spins the whole service down after 15 minutes with no traffic at
    all, and SETUP_GUIDE.md already tells Yitzi the very next request
    after that "can take about a minute to wake it back up." A real
    account hit this for real: a plain 15-second timeout on a totally
    ordinary query (checking a setting during startup) raised
    SharedDatabaseUnreachableError with nothing at all wrong other than
    the service being asleep at that exact moment. 60 seconds costs
    nothing on a warm connection (which answers in milliseconds either
    way) and gives a cold one a real chance to finish waking up within
    a single request instead of failing it outright."""
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        base_url.rstrip("/") + path, data=data, method="POST",
        headers={"Content-Type": "application/json"},
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return resp.status, json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        try:
            body = json.loads(e.read().decode("utf-8"))
        except Exception:
            body = {"detail": str(e)}
        return e.code, body
    except urllib.error.URLError as e:
        raise SharedDatabaseUnreachableError(f"Couldn't reach the shared database service: {e.reason}")
    except Exception as e:
        raise SharedDatabaseUnreachableError(f"Couldn't reach the shared database service: {e}")


def remote_backend_login(base_url, username, password):
    """Calls the shared-database backend's own POST /login (server/app.py,
    Batch 159) with the SAME username/password already checked against the
    local users table today -- there is no separate database secret for a
    person to learn or for this call to need. Returns (token, user_dict)
    on success. Raises MustSetPasswordError for a brand new account that
    hasn't chosen a real password yet, or SharedDatabaseUnreachableError
    for everything else that stops a normal login (wrong password, a
    disabled account, or the backend being unreachable at all) -- the
    caller shows that message directly, the same way it already does for
    Turso's own connection errors today."""
    status, body = _remote_post_json(base_url, "/login", {"username": username, "password": password})
    if status == 403 and body.get("detail") == "must_set_password":
        raise MustSetPasswordError(username)
    if status != 200:
        raise SharedDatabaseUnreachableError(body.get("detail") or "Couldn't log in to the shared database.")
    return body["token"], body["user"]


def remote_backend_set_password(base_url, username, password, new_password):
    """Calls POST /set_password -- used both for a must_set_password
    account's first real password (password=None/blank) and an ordinary
    "I know my current password and want to change it" case. Returns
    (token, user_dict) on success, same shape as remote_backend_login(),
    since the backend logs the account in as part of this same call."""
    status, body = _remote_post_json(base_url, "/set_password", {
        "username": username, "password": password, "new_password": new_password,
    })
    if status != 200:
        raise SharedDatabaseUnreachableError(body.get("detail") or "Couldn't set the new password.")
    return body["token"], body["user"]


def bootstrap_migration_token(base_url, timeout=60):
    """Calls POST /bootstrap_migrate_token (server/app.py, Batch 168) --
    the one way to get a session token for precheck_backend_migration()
    to use when the shared database has never had anyone log in before,
    since an ordinary remote_backend_login() is impossible before the
    migration itself has copied even one real user row across. Only ever
    succeeds while the backend's users table is genuinely empty --
    raises BackendAlreadySetUpError specifically when real accounts
    already exist (HTTP 403; the caller should fall back to an ordinary
    login with one), or the plain SharedDatabaseUnreachableError for
    every other failure (unreachable, timed out, or the database's
    tables not existing yet at all) -- see BackendAlreadySetUpError's own
    docstring for why that distinction matters."""
    status, body = _remote_post_json(base_url, "/bootstrap_migrate_token", {}, timeout=timeout)
    if status == 403:
        raise BackendAlreadySetUpError(body.get("detail") or "This database already has real accounts on it.")
    if status != 200:
        raise SharedDatabaseUnreachableError(
            body.get("detail") or "Couldn't start the migration."
        )
    return body["token"]


def test_backend_connection(base_url, timeout=60):
    """Batch 164: the hosted-backend equivalent of test_turso_connection()
    above -- what Settings' own "Test connection" button on the new
    "Shared database (hosted backend)" card calls. Deliberately needs no
    username/password: GET /health (server/app.py) needs no login by
    design, so this can prove the SERVICE itself is up and reachable
    before anyone has an account to log in with at all -- a real login is
    checked separately, by BackendLoginDialog actually logging in.
    Returns (True, message) on success or (False, message) on any failure
    -- never raises, so a bad paste can't crash the dialog. timeout
    defaults higher than most of this file's other network calls (15s
    elsewhere) since Render's free tier can take about a minute to wake a
    service back up from its own idle sleep -- still not the full minute
    (a "Test connection" click that just hangs looking stuck would be
    worse than an honest "try again shortly" failure), but longer than a
    normal already-awake round trip needs."""
    base_url = (base_url or "").strip()
    if not base_url:
        return False, "Enter the backend's address first."
    if not (base_url.startswith("http://") or base_url.startswith("https://")):
        return False, (
            "The address should start with http:// or https://, e.g. "
            "https://yj-po-generator-backend.onrender.com"
        )
    try:
        req = urllib.request.Request(base_url.rstrip("/") + "/health", method="GET")
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            body = json.loads(resp.read().decode("utf-8"))
    except Exception as e:
        return False, (
            "Couldn't reach this address. If it was just deployed, or has been idle a "
            f"while and is waking back up, this can take about a minute -- try again "
            f"shortly. ({e})"
        )
    if body.get("status") != "ok":
        return False, "Reached this address, but it didn't respond the way this app expects -- double-check it."
    if not body.get("database_reachable"):
        return False, "Connected to the service, but it can't reach its own database right now -- check its DATABASE_URL setting."
    return True, "Connected successfully."


# Every sqlite3 statement type this file ever expects to open an implicit
# transaction, mirroring sqlite3's own legacy (still-default) behaviour:
# BEGIN/INSERT/UPDATE/DELETE (and REPLACE, never actually used remotely --
# see TECHNICAL_REBUILD_SPEC.md's Batch 157 audit -- kept for parity) open
# one; a bare SELECT does not. _RemoteConnection.in_transaction (below)
# relies on exactly this rule to track transaction state client-side
# without an extra round trip to ask the server.
_DML_STATEMENT_STARTS = ("BEGIN", "INSERT", "UPDATE", "DELETE", "REPLACE")


class _RemoteCursor:
    """Mimics just the sqlite3.Cursor surface this file's ~700 conn.execute()
    call sites actually use (see TECHNICAL_REBUILD_SPEC.md's Batch 157
    audit): iteration, fetchone()/fetchall(), .lastrowid, .rowcount.
    Rows already arrive from the backend as plain JSON objects -- ordinary
    Python dicts support row["col"] and dict(row) exactly like sqlite3.Row
    does everywhere in this file already, so no separate row-wrapper class
    is needed at all."""

    def __init__(self, rows, lastrowid, rowcount):
        self._rows = list(rows)
        self._pos = 0
        self.lastrowid = lastrowid
        self.rowcount = rowcount

    def fetchone(self):
        if self._pos >= len(self._rows):
            return None
        row = self._rows[self._pos]
        self._pos += 1
        return row

    def fetchall(self):
        rest = self._rows[self._pos:]
        self._pos = len(self._rows)
        return rest

    def __iter__(self):
        while self._pos < len(self._rows):
            row = self._rows[self._pos]
            self._pos += 1
            yield row


class _RemoteConnection:
    """Talks to the small authenticated backend (server/app.py, Batch 159)
    over plain HTTPS instead of opening sqlite3 or a turso replica directly
    -- see FEATURE_LOG.md's Batch 157/159 entries for why (no database
    credential can ever ship inside the publicly-distributed installer
    once every request goes through this instead; the backend is the only
    thing that ever holds one). Deliberately mimics just the
    sqlite3.Connection/Cursor surface this file's existing ~700
    conn.execute(...) call sites actually use -- execute()/executemany(),
    commit()/rollback(), in_transaction, and cursor fetchone()/fetchall()/
    iteration/.lastrowid/.rowcount -- so none of those call sites need to
    change at all, the same design already proven for _PushOnCommitConnection
    above.

    Two things this file's SQL actually needs that go beyond a plain `?`
    placeholder, both handled transparently: save_po()'s two statements
    use SQLite's `:name` named-placeholder style with a dict of values
    (translate_sql() in server/sql_translate.py rewrites `:name` to
    psycopg2's own `%(name)s` pyformat style server-side, and this class
    simply forwards whatever params it's given -- list/tuple or dict --
    as JSON exactly as sqlite3.Connection.execute() already accepts
    either); and a bare `PRAGMA ...` statement (get_connection() runs
    "PRAGMA foreign_keys = ON" unconditionally after opening any
    connection) is answered locally as a harmless no-op with no network
    call at all, since Postgres has no equivalent pragma and always
    enforces foreign keys unconditionally already -- there is nothing for
    it to turn on.

    executescript() deliberately raises rather than silently doing
    nothing: schema management is the backend's own job now (its Postgres
    schema already exists, hand-built and verified table-for-table in
    Batch 158), never something a client pushes -- if this is ever called
    it means a future change wired schema/migration code into the remote
    path by mistake, and that should fail loudly in testing, not silently
    no-op in front of Yitzi."""

    def __init__(self, base_url, token):
        self._base_url = base_url
        self._token = token
        self._in_transaction = False
        # Accepted for API parity with sqlite3.Connection (get_connection()
        # sets this unconditionally on whichever connection type it opens)
        # but never consulted -- rows are already dict-like, see
        # _RemoteCursor's own docstring above.
        self.row_factory = None

    @property
    def in_transaction(self):
        return self._in_transaction

    def execute(self, sql, params=(), timeout=60):
        stripped = sql.strip()
        if stripped[:6].upper() == "PRAGMA":
            return _RemoteCursor([], None, -1)
        wire_params = dict(params) if isinstance(params, dict) else list(params)
        status, body = _remote_post_json(self._base_url, "/execute", {
            "token": self._token, "sql": sql, "params": wire_params,
        }, timeout=timeout)
        if status == 401:
            raise SharedDatabaseUnreachableError(
                body.get("detail") or "Your shared-database session has expired. Please log in again."
            )
        if status != 200:
            raise sqlite3.OperationalError(body.get("detail") or f"remote query failed (HTTP {status})")
        if stripped.upper().startswith(_DML_STATEMENT_STARTS):
            self._in_transaction = True
        return _RemoteCursor(body.get("rows") or [], body.get("lastrowid"), body.get("rowcount", -1))

    def executemany(self, sql, seq_of_params):
        cur = None
        for params in seq_of_params:
            cur = self.execute(sql, params)
        return cur if cur is not None else _RemoteCursor([], None, 0)

    def executescript(self, sql):
        raise NotImplementedError(
            "_RemoteConnection.executescript() should never be called -- the shared-database "
            "backend already owns its own Postgres schema (see server/schema_postgres.sql); "
            "schema/migration code must be skipped entirely whenever the connection is remote."
        )

    def commit(self):
        status, body = _remote_post_json(self._base_url, "/commit", {"token": self._token})
        if status != 200:
            raise sqlite3.OperationalError(body.get("detail") or f"remote commit failed (HTTP {status})")
        self._in_transaction = False

    def rollback(self):
        status, body = _remote_post_json(self._base_url, "/rollback", {"token": self._token})
        if status != 200:
            raise sqlite3.OperationalError(body.get("detail") or f"remote rollback failed (HTTP {status})")
        self._in_transaction = False

    def close(self):
        # Best-effort -- a network drop on the way out shouldn't stop the
        # app from closing, and the backend's own idle-session sweep
        # (server/app.py's SESSION_IDLE_TIMEOUT_SECONDS) cleans up an
        # un-logged-out session eventually either way.
        try:
            _remote_post_json(self._base_url, "/logout", {"token": self._token}, timeout=5)
        except Exception:
            pass

    def reauthenticate(self, token):
        """Batch 171: swaps in a freshly-issued session token on this SAME
        object, in place -- for the real-world case of a session dying
        while the app is just sitting there (server/app.py's own idle
        sweep, or Render's free tier spinning the whole service down
        after 15 minutes with no traffic at all -- either wipes every
        in-memory session instantly, regardless of how recently it was
        issued). get_connection() caches one _RemoteConnection per
        process and hands the SAME object to everything that asks for
        it, so replacing self._token here (rather than constructing a
        new _RemoteConnection nobody but the caller would ever see) is
        what actually makes a fresh login take effect for every other
        piece of code already holding this same connection -- MainWindow
        and its pages never hold their own separate reference, they all
        call core.get_connection() themselves each time."""
        self._token = token


class MigrationTargetNotEmptyError(Exception):
    """Raised by precheck_backend_migration() (Batch 168) when the shared
    database already has at least one row in some table it's about to
    receive this computer's local data into. Deliberately refuses rather
    than guessing whether that's a previous migration attempt, someone's
    already typed something in, or a genuine mistake -- copying on top of
    it either duplicates rows or silently masks whichever of those it
    actually is. `table` names the first non-empty table found (checked
    in BACKEND_MIGRATION_TABLE_ORDER, so this is always the same one for
    the same database)."""

    def __init__(self, table):
        self.table = table
        super().__init__(
            f"The shared database already has data in '{table}' -- this looks like it's "
            f"already been migrated, or already has real data in it. Refusing to copy on "
            f"top of it without checking further."
        )


# Batch 168: the order every table must be copied in for the one-time
# local-to-shared-database migration -- parents before children, matching
# the real REFERENCES constraints server/schema_postgres.sql actually
# enforces (unlike SQLite, a genuinely separate database engine, Postgres
# really does reject a child row whose parent id doesn't exist yet, so
# unlike everywhere else in this file, statement order here isn't just
# cosmetic). The five tables with no `id` column at all (mirroring
# sql_translate.py's own TABLES_WITHOUT_ID, which exists for exactly the
# same reason -- a blanket "RETURNING id" would be wrong for them) don't
# need a sequence caught up afterwards; every other table does (see
# migrate_one_table_to_backend()).
BACKEND_MIGRATION_TABLE_ORDER = [
    "settings", "addresses", "suppliers", "products",
    "users", "user_permissions",
    "purchase_orders", "po_items", "po_events", "po_receiving_events",
    "product_merge_ignored", "supplier_merge_ignored", "zoho_vendor_map",
    "supplier_issues",
    "price_requests", "price_request_replies",
    "savings_categories", "savings",
    "supplier_issue_categories", "purchase_owner_prefixes",
    "notifications", "product_requests",
]


class BackendMigration:
    """Everything migrate_one_table_to_backend() needs held open across
    every table it's called for -- one real local read-only connection
    and one real logged-in remote session, both opened once by
    precheck_backend_migration() rather than per table. finish() commits
    the whole migration as one transaction (server/app.py's /commit
    covers everything sent through this one session, table by table) and
    logs the session out; abort() rolls back instead, for a caller that
    hits a failure partway through and wants to leave the shared database
    exactly as empty as it found it, not partially filled."""

    def __init__(self, local_conn, remote_conn):
        self.local_conn = local_conn
        self.remote_conn = remote_conn

    def finish(self):
        self.remote_conn.commit()
        self.local_conn.close()
        self.remote_conn.close()

    def abort(self):
        try:
            self.remote_conn.rollback()
        except Exception:
            pass
        self.local_conn.close()
        self.remote_conn.close()


def precheck_backend_migration(base_url, token):
    """Step 1 of the one-time local -> shared-database migration (Batch
    168 -- Yitzi's own real Render URL was confirmed live, so this is the
    piece every earlier batch's "what's still ahead" note kept pointing
    at). Takes an already-logged-in session token -- the caller is
    expected to have just proven a real login via BackendLoginDialog (or,
    at the po_core.py level, remote_backend_login()) rather than this
    function taking a username/password and logging in a second time
    itself.

    Opens this computer's local database file read-only (raises
    DatabaseCorruptError/RuntimeError exactly as open_view_only_connection()
    does for a damaged or missing file -- deliberately reusing the exact
    same checks, not a separate, potentially-inconsistent copy of them),
    then checks every table in BACKEND_MIGRATION_TABLE_ORDER on the
    backend is currently empty, raising MigrationTargetNotEmptyError
    immediately -- before copying a single row -- if even one of them
    already has data.

    Returns a BackendMigration to pass, once per table in
    BACKEND_MIGRATION_TABLE_ORDER (order matters -- see that list's own
    docstring), to migrate_one_table_to_backend(), then to finish() once
    every one has succeeded, or abort() if any of them raised."""
    remote_conn = _RemoteConnection(base_url, token)

    ok, detail = check_db_integrity(DB_PATH)
    if not ok:
        remote_conn.close()
        raise DatabaseCorruptError(detail)
    if not DB_PATH.exists():
        remote_conn.close()
        raise RuntimeError("There's no local database on this computer to migrate yet.")
    local_conn = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True, check_same_thread=False)
    local_conn.row_factory = sqlite3.Row

    try:
        for table in BACKEND_MIGRATION_TABLE_ORDER:
            count = remote_conn.execute(f"SELECT COUNT(*) AS c FROM {table}").fetchone()["c"]
            if count:
                raise MigrationTargetNotEmptyError(table)
    except Exception:
        local_conn.close()
        remote_conn.close()
        raise

    return BackendMigration(local_conn, remote_conn)


def migrate_one_table_to_backend(migration, table, batch_size=200):
    """Step 2, called once per table in BACKEND_MIGRATION_TABLE_ORDER, in
    that order (see its own docstring for why order matters). Copies
    every local row into the backend in batches of `batch_size` -- one
    real multi-row INSERT per batch, not one network round trip per row,
    the difference between a few dozen requests and several thousand over
    a real connection to Render. A table with no rows locally, or that
    doesn't even exist in this computer's own database yet (an old
    database that predates whatever batch introduced a newer table), is
    skipped cleanly, not treated as an error.

    Column names for the INSERT are read directly off the local table's
    own cursor description rather than hardcoded here, so this doesn't
    silently go stale the next time a column gets added -- the two
    schemas (SQLite's SCHEMA constant + its migrations, and
    server/schema_postgres.sql) are hand-kept in sync already; this just
    trusts that, the same way _RemoteConnection's whole design already
    trusts every one of po_core.py's ~700 other query strings to work
    against either database unchanged.

    For any table with an `id` column, the Postgres SERIAL sequence
    behind it is explicitly caught up afterwards to the highest id just
    inserted -- inserting an explicit id value never advances a SERIAL
    sequence on its own, and every id column in this schema is a real
    SERIAL, so without this the very next ordinary INSERT anyone makes
    through the app once it's live would collide with a row this
    migration just wrote.

    Returns the number of rows actually copied for this table."""
    local_conn, remote_conn = migration.local_conn, migration.remote_conn
    exists = local_conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table,)
    ).fetchone()
    if exists is None:
        return 0

    cur = local_conn.execute(f"SELECT * FROM {table}")
    columns = [d[0] for d in cur.description]
    rows = cur.fetchall()
    if not rows:
        return 0

    col_list = ",".join(columns)
    row_placeholders = "(" + ",".join(["?"] * len(columns)) + ")"
    has_id = "id" in columns
    max_id = None

    for start in range(0, len(rows), batch_size):
        batch = rows[start:start + batch_size]
        sql = f"INSERT INTO {table} ({col_list}) VALUES " + ",".join([row_placeholders] * len(batch))
        params = []
        for row in batch:
            params.extend(row[c] for c in columns)
            if has_id and (max_id is None or row["id"] > max_id):
                max_id = row["id"]
        remote_conn.execute(sql, params, timeout=60)

    if has_id and max_id is not None:
        remote_conn.execute(
            "SELECT setval(pg_get_serial_sequence(?, 'id'), ?, true)",
            (table, max_id), timeout=30,
        )
    return len(rows)


def shared_database_sync_status():
    """None if this computer isn't using the shared database at all,
    otherwise {"ok": bool, "last_error": str, "consecutive_failures": int}
    reflecting whether the most recent push/pull actually succeeded.
    Batch 110: lets the UI show a persistent, hard-to-miss warning the
    moment syncing stops working, rather than someone working for hours
    without knowing their changes aren't reaching anyone else. Batch 150
    added consecutive_failures -- see shared_database_sync_looks_stuck()."""
    if not isinstance(_connection, _PushOnCommitConnection):
        return None
    return {
        "ok": _connection.sync_ok,
        "last_error": _connection.sync_last_error,
        "consecutive_failures": _connection.consecutive_sync_failures,
    }


# Batch 150: Yitzi's real app.log from a live incident showed the shared
# database can get stuck mid-session, not just at startup (see Batch 149's
# comment above _TURSO_ERROR_CATEGORIES for the full story) -- once the
# periodic pull starts failing with the "unable to checkpoint synced
# portion of WAL" signature, it kept failing identically, roughly every 20
# seconds, for over 9 hours straight overnight, never once clearing on its
# own. Before this batch, the only sign of that was the passive "not
# synced" banner (see MainWindow.check_shared_database_sync) sitting there
# indefinitely -- accurate, but it left the actual fix (reconnecting with
# a fresh local replica, exactly what Batch 149's startup self-heal now
# does automatically) something only Settings > Data & Backup could do,
# which meant noticing the banner, knowing what it meant, and knowing
# where to go.
#
# SYNC_STUCK_THRESHOLD is how many pushes/pulls in a row have to fail
# before this stops looking like an ordinary, temporary network blip (a
# momentary wifi drop, a laptop still finishing waking up) and starts
# looking like the same "the local replica needs a fresh pull, not just
# another retry" problem Batch 149 already has a real fix for. 3 was
# chosen deliberately low: pushes/pulls that fail this way in Yitzi's log
# never once succeeded again on their own however long they were left, so
# there's little value in waiting longer just to be "more sure" -- the
# earlier this becomes visible and actionable, the less time spent
# silently unsynced.
SYNC_STUCK_THRESHOLD = 3


def shared_database_sync_looks_stuck():
    """True once shared_database_sync_status()'s consecutive_failures has
    reached SYNC_STUCK_THRESHOLD -- see that constant's own comment. False
    (never "stuck") for a plain local install, same as
    shared_database_sync_status() returning None for that case."""
    status = shared_database_sync_status()
    return bool(status) and status["consecutive_failures"] >= SYNC_STUCK_THRESHOLD


def pull_latest(conn):
    """Batch 129: pulls down anything other computers have pushed to the
    shared database since this connection last synced -- called from
    MainWindow's own 20-second _sync_pull_timer (po_generator_qt.py), not
    from anywhere in this file, since before this batch the ONLY time an
    already-open session ever re-synced at all was get_connection()'s own
    one-time connect at startup. Yitzi: "if i place a PO it needs to show
    up on Stock PC instantly not they they need to wait for it to cathc
    up if they are not fully in scnc this may cause issues" -- exactly
    that gap.

    A safe no-op returning False for a plain local (non-shared)
    connection -- there's nothing to pull, and every existing caller of
    get_connection() (this whole file, every verify_batchN.py test) gets
    back an ordinary sqlite3.Connection with no .pull() method at all in
    that case, so this must never assume conn is a _PushOnCommitConnection
    just because it was given one.

    A failure here (offline, server unreachable, a momentary network
    blip) is logged and never raised -- mirrors _PushOnCommitConnection.
    commit()'s own push failure handling exactly, including updating the
    SAME sync_ok/sync_last_error the UI's "not synced" warning banner
    already polls via shared_database_sync_status(), so a periodic pull
    failing shows up through the exact same existing warning a push
    failure would, rather than needing a second, parallel indicator.

    Returns True only if a pull was actually attempted AND succeeded --
    callers (MainWindow._on_sync_pull_timer) use this to decide whether
    refreshing the screen is actually worth doing, rather than doing it
    unconditionally every 20 seconds regardless of whether anything could
    possibly have changed."""
    if not isinstance(conn, _PushOnCommitConnection):
        return False
    try:
        conn.pull()
    except Exception as e:
        object.__setattr__(conn, "sync_ok", False)
        object.__setattr__(conn, "sync_last_error", str(e))
        object.__setattr__(conn, "consecutive_sync_failures", conn.consecutive_sync_failures + 1)
        log.warning("Periodic pull from the shared database failed: %s", e)
        return False
    object.__setattr__(conn, "sync_ok", True)
    object.__setattr__(conn, "sync_last_error", "")
    object.__setattr__(conn, "consecutive_sync_failures", 0)
    return True


def _ensure_column(conn, table, column, coltype_and_default):
    """Adds a column to an existing table if it isn't there yet -- for
    existing installs whose database file predates a column that's now part
    of the schema. CREATE TABLE IF NOT EXISTS (used everywhere else in
    SCHEMA) only helps for brand new tables; a column added to a table that
    already exists on disk needs an explicit ALTER TABLE, guarded by this
    check so it's safe to call on every startup."""
    cols = {row["name"] for row in conn.execute(f"PRAGMA table_info({table})")}
    if column not in cols:
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {coltype_and_default}")
        conn.commit()


def _ensure_schema_ddl(conn):
    """Every DDL statement this app has ever needed -- every CREATE TABLE/
    INDEX in SCHEMA, plus every column ever added by a migration -- and
    NOTHING else, no DML at all. Split out of _prepare_and_migrate_schema()
    (below) in Batch 154/155 so setup_shared_database() can push this half
    on its own, fully landed on the remote, before ANY DML (a backfill
    UPDATE, default-seeding INSERTs, migration writes, or this computer's
    own real data) ever touches a table or column this creates in the SAME
    push batch -- see setup_shared_database()'s own comment and this
    batch's FEATURE_LOG entry for why that combination can fail against
    this version of pyturso's sync engine. Safe to re-run any number of
    times on the same connection (which setup_shared_database() does --
    once directly, then again inside _prepare_and_migrate_schema() right
    after): CREATE TABLE/INDEX IF NOT EXISTS and _ensure_column() are both
    already fully idempotent, so a second pass is a harmless no-op."""
    conn.executescript(SCHEMA)
    conn.commit()
    _ensure_column(conn, "purchase_orders", "flagged_for_review", "INTEGER DEFAULT 0")
    _ensure_column(conn, "purchase_orders", "zoho_exported", "INTEGER DEFAULT 0")
    _ensure_column(conn, "purchase_orders", "zoho_exported_at", "TEXT DEFAULT ''")
    _ensure_column(conn, "supplier_issues", "category", "TEXT DEFAULT ''")
    _ensure_column(conn, "supplier_issues", "subject", "TEXT DEFAULT ''")
    _ensure_column(conn, "purchase_orders", "purchase_owner", "TEXT DEFAULT ''")
    _ensure_column(conn, "purchase_orders", "fx_rate", "REAL DEFAULT 1.0")
    _ensure_column(conn, "purchase_orders", "base_total", "REAL DEFAULT 0")
    _ensure_column(conn, "purchase_orders", "zoho_po_status", "TEXT DEFAULT ''")
    _ensure_column(conn, "suppliers", "credit_limit", "REAL DEFAULT 0")
    _ensure_column(conn, "suppliers", "credit_reset_at", "TEXT DEFAULT ''")
    _ensure_column(conn, "po_items", "margin_vat", "INTEGER DEFAULT 0")
    _ensure_column(conn, "po_items", "original_qty", "REAL DEFAULT NULL")
    _ensure_column(conn, "purchase_orders", "created_by_user_id", "INTEGER DEFAULT NULL")
    _ensure_column(conn, "users", "must_set_password", "INTEGER NOT NULL DEFAULT 0")
    _ensure_column(conn, "users", "dashboard_widgets", "TEXT DEFAULT NULL")
    _ensure_column(conn, "users", "notification_role", "TEXT DEFAULT ''")
    # Batch 157: zoho_vendor_map's own primary key is the vendor NAME
    # (zoho_vendor_name), not a real id -- reverse_zoho_vendor_map() relies
    # on SQLite's implicit rowid to know which mapping for a given supplier
    # was added most recently, but Postgres has no equivalent implicit
    # rowid. seq is an explicit stand-in: set once, on insert, to one more
    # than the current highest seq (never touched again by
    # set_zoho_vendor_map()'s own ON CONFLICT ... DO UPDATE, exactly
    # mirroring how a real rowid never changes on an UPDATE either) --
    # see _migrate_zoho_vendor_map_seq() below for backfilling existing rows.
    _ensure_column(conn, "zoho_vendor_map", "seq", "INTEGER DEFAULT 0")
    conn.commit()


def _prepare_and_migrate_schema(conn):
    """Everything a raw connection -- turso replica or plain sqlite3 --
    needs before it's fit to use: bringing the schema up to date, every
    one-time column/data migration, and default seeding. This used to be
    inline in get_connection() only; Batch 125 pulled it out into its own
    function so setup_shared_database()'s one-time seed replica (see
    below) goes through the exact same preparation as every normal
    launch, rather than a hand-copied approximation of it that could
    silently drift out of sync over time. Raises SchemaTooNewError if the
    connection's PRAGMA user_version is ahead of what this build
    understands -- the caller decides what to do about that (get_connection()
    clears the global _connection; setup_shared_database() just lets it
    propagate, since a brand new seed file can never actually hit this)."""
    db_schema_version = conn.execute("PRAGMA user_version").fetchone()[0]
    if db_schema_version > APP_SCHEMA_VERSION:
        conn.close()
        raise SchemaTooNewError(db_schema_version, APP_SCHEMA_VERSION)
    _ensure_schema_ddl(conn)
    # Backfill base_total for rows saved before this column existed -- they
    # predate any per-PO currency picker, so they're all effectively already
    # in the home currency (1:1). Safe to re-run on every launch: every row
    # this feature creates always sets a non-zero base_total itself (unless
    # the PO is genuinely worth 0), so this only ever touches pre-existing
    # rows that still show the column default.
    conn.execute("UPDATE purchase_orders SET base_total = total WHERE base_total = 0 OR base_total IS NULL")
    conn.commit()
    _seed_defaults(conn)
    if not get_setting(conn, "po_report_to", ""):
        set_settings(conn, {
            "po_report_to": DEFAULT_SETTINGS["po_report_to"],
            "po_report_recipient_name": DEFAULT_SETTINGS["po_report_recipient_name"],
        })
    _migrate_from_legacy_json(conn)
    _migrate_po_ref_prefix(conn)
    _migrate_report_frequency_to_multi(conn)
    _migrate_seed_supplier_name_acronyms(conn)
    _seed_own_purchase_owner_prefix(conn)
    _migrate_zoho_vendor_map_seq(conn)
    if db_schema_version < APP_SCHEMA_VERSION:
        conn.execute(f"PRAGMA user_version = {APP_SCHEMA_VERSION}")
        conn.commit()


def _pull_with_retry(replica, attempts=3, delay_seconds=1.5):
    """Batch 138: calls replica.pull(), retrying a few times before giving
    up instead of failing on the very first hiccup.

    Yitzi hit this directly: the "Fix the shared database connection"
    dialog's "Test connection" button (test_turso_connection(), which also
    does a connect+pull against a brand new throwaway replica) reported
    "Connected successfully" -- proving the URL, auth token, and network
    path were all genuinely fine -- and then clicking "Save and retry"
    straight after failed with the exact same
    "Couldn't sync with the shared database: ... unable to checkpoint
    synced portion of WAL: result=CheckpointResult { wal_max_frame: 0,
    ... }, watermark=21" error the original blocking dialog was showing.

    Reading turso's own bundled sync engine (lib_sync.py's
    ConnectionSync.pull(), which calls wait_changes() then
    apply_changes()) confirmed that checkpoint error is raised from deep
    inside the compiled Rust sync engine's apply_changes() step, not from
    anything this app controls -- and that setup_shared_database() and
    test_turso_connection() both do the exact same "brand new local file,
    connect, pull" sequence against the same remote. Two functionally
    identical pulls, moments apart, one succeeding and one failing with a
    CheckpointResult where every counter is 0/false (meaning no actual
    checkpoint work happened before it errored) points at a transient
    hiccup inside the sync engine itself -- not a real, permanent problem
    with this particular remote or these particular credentials -- so the
    right fix is to let a pull that fails this way simply try again a
    couple of times before treating it as a real failure, rather than
    a single momentary blip locking Yitzi out of "Save and retry" (or,
    via get_connection()'s own replica.pull() call, out of opening the
    program at all).

    Used by both setup_shared_database() (below) and get_connection()'s
    shared-database branch, so both the "Fix connection" dialog and an
    ordinary launch benefit. Deliberately NOT used inside
    _PushOnCommitConnection.commit()'s push() -- a push failure there is
    already handled as a non-blocking "will retry on the next commit"
    background condition (see Batch 110 notes above), so retrying inline
    there would just add delay to every save for no benefit.

    Raises whatever the final attempt's pull() raised if every attempt
    fails. Returns pull()'s own return value on success."""
    last_error = None
    for attempt in range(1, attempts + 1):
        try:
            return replica.pull()
        except Exception as e:
            last_error = e
            if attempt < attempts:
                log.warning(
                    "Shared database pull failed (attempt %d of %d), retrying: %s",
                    attempt, attempts, e,
                )
                time.sleep(delay_seconds)
    raise last_error


def setup_shared_database(url, auth_token):
    """The one-time, careful process of turning this computer's shared
    database on -- called by Settings' "Save" button (_save_turso_config
    in po_generator_qt.py) instead of just writing turso_config.json
    directly.

    Batch 125: this replaces the original Batch 109/110 design, which just
    wrote turso_config.json and left the actual first connect (and the
    seeding it never really did) to the next app launch's get_connection().
    That design lost a real customer's live database: get_connection()
    handed _turso.sync.connect() the SAME file DB_PATH that already held
    everything Yitzi had saved -- a perfectly ordinary, never-synced plain
    sqlite3 file, no different in kind from any other pre-Batch-109
    install. turso.sync has no way to know a plain file like that is
    "full of real data to keep" rather than just an unformatted path to
    initialize -- its own bootstrap_if_empty logic judges "empty" by
    whether the file already carries turso's own sync bookkeeping, not by
    row count, and a plain sqlite3 file never does. Against a brand new,
    genuinely empty remote, that reset the local file to match, and
    get_connection()'s own executescript(SCHEMA) call right afterward
    quietly recreated an empty, schema-only database on top -- which is
    exactly what showed up as "asking me to make a new admin user."
    Nothing was actually destroyed on the remote (there was nothing there
    to lose), but the local file was reset with nothing ever having been
    uploaded first, despite the Settings confirm dialog's promise that it
    would be.

    This function fixes that by never handing turso.sync a file that
    already has data to lose. It always takes a real backup first no
    matter what happens next; then it connects a brand new, disposable
    temp file to the remote (the exact scenario turso.sync's own docs
    show -- nothing pre-existing for it to misjudge), and only if the
    remote turns out to be genuinely empty of real content does it copy
    this computer's existing data into that connection with ordinary SQL
    and push it up -- the documented "make local changes, commit, push"
    pattern, not a file-level trick. Only after that upload is verified to
    have actually landed on the remote does this freshly-seeded,
    now-properly-synced file get swapped into place as DB_PATH itself
    (the untouched original is kept alongside, timestamped, never
    deleted) -- so the very next ordinary launch's get_connection() finds
    a file turso.sync already recognizes as a real replica, not another
    unformatted file it might reset all over again. turso_config.json
    itself is only written at the very end, once every step above has
    already succeeded -- if anything fails along the way, nothing is
    saved, DB_PATH is completely untouched, and the caller gets a message
    explaining what went wrong.

    Raises SharedDatabaseUnreachableError, ValueError, or RuntimeError on
    any failure. Returns None on success."""
    if _turso is None:
        raise SharedDatabaseUnreachableError(
            "This copy of the program doesn't include the component the shared "
            "database needs (the 'turso' package). Update to a build that includes it."
        )
    url = (url or "").strip()
    auth_token = (auth_token or "").strip()
    if not url or not auth_token:
        raise ValueError("Enter both the database URL and the auth token first.")

    global _connection
    backup_path = backup_now("pre-shared-db-switch")
    if not backup_path:
        raise RuntimeError(
            "Couldn't take a safety backup, so the shared database wasn't set up. "
            "See the log file for details."
        )

    # Batch 127: the global _connection (the SAME object this app's own
    # already-open window is using as self.main.conn) used to get closed
    # right here, long before any of the risky work below even started.
    # Every ordinary read against DB_PATH below (copying rows out,
    # verifying counts) works perfectly well with it still open -- SQLite
    # allows multiple simultaneous connections to one file for that. The
    # only operation that genuinely needs every handle on DB_PATH released
    # is the rename right at the end, so closing it happens there instead,
    # only once everything else has already succeeded. Closing it this
    # early meant ANY failure after this point (a network hiccup, a failed
    # push, a failed upload verification, or -- exactly what actually
    # happened to Yitzi -- the rename itself failing on Windows because
    # two other, separately-fixed, unclosed connections were still
    # holding the file open) left this app's own already-open window
    # holding a dead, closed connection object, which then crashed with
    # "Cannot operate on a closed database" the next time anything
    # (its periodic refresh timer, any button) tried to use it.

    import tempfile
    work_dir = Path(tempfile.mkdtemp(prefix="yj_turso_seed_"))
    try:
        seed_path = work_dir / "seed.db"
        try:
            replica = _turso.sync.connect(str(seed_path), remote_url=url, auth_token=auth_token)
            replica.row_factory = _turso.Row
        except Exception as e:
            raise SharedDatabaseUnreachableError(
                f"Couldn't connect to the shared database: {_diagnose_turso_error('connect (setup)', url, e)}"
            )
        try:
            _pull_with_retry(replica)
        except Exception as e:
            try:
                replica.close()
            except Exception:
                pass
            raise SharedDatabaseUnreachableError(
                f"Couldn't sync with the shared database: {_diagnose_turso_error('pull (setup)', url, e)}"
            )

        # Batch 154/155: run EVERY piece of DDL this app has ever needed --
        # every CREATE TABLE/INDEX, and every column any migration has ever
        # added (_ensure_schema_ddl(), which _prepare_and_migrate_schema()
        # below also calls as its own first step) -- and push THAT on its
        # own, in its own completed push() call, before ANY DML runs
        # anywhere on this replica: not _prepare_and_migrate_schema()'s own
        # migrations/default-seeding below (a backfill UPDATE against
        # purchase_orders.base_total, INSERTs seeding default settings/
        # addresses rows, and more), and not a single row of this
        # computer's real data further down. Yitzi hit this directly,
        # twice, creating a genuinely new/empty shared database -- first
        # "... no such table: purchase_orders" (Batch 154, a brand-new
        # TABLE written to in the same push it was created in), then,
        # after that fix, "... no such column: base_total" (this table
        # already existed on the remote from the Batch 154 fix above, but
        # the ALTER TABLE that added base_total to it was still sitting in
        # the SAME push batch as the backfill UPDATE that writes to that
        # column -- the identical failure shape, one level down, at the
        # column instead of the table). This version of pyturso's sync
        # engine (Batch 153 pinned it to 0.7.2) replicates changes as a
        # logical operation log (its own docstring calls these "CDC
        # operations"), not raw pages -- and the evidence points at
        # push()'s remote-side apply step validating each queued statement
        # against the remote's schema as it stood before the batch
        # started, rather than against its own still-being-applied earlier
        # statements in the very same batch, for ALTER TABLE just as much
        # as CREATE TABLE. A push containing ONLY DDL (_ensure_schema_ddl()
        # is pure CREATE TABLE/CREATE INDEX/ALTER TABLE ADD COLUMN, nothing
        # else) has nothing to fail that validation against, so doing that
        # push alone first, before ANY DML happens anywhere on this
        # replica, sidesteps the whole class of failure regardless of the
        # exact mechanism -- by the time migrations, seed defaults, or
        # this computer's real data ever get written, every table AND
        # every column they touch already genuinely exists on the remote.
        try:
            _ensure_schema_ddl(replica)
            replica.push()
        except Exception as e:
            try:
                replica.close()
            except Exception:
                pass
            raise SharedDatabaseUnreachableError(
                f"Connected, but couldn't prepare the shared database's schema: {e}"
            )

        try:
            _prepare_and_migrate_schema(replica)
        except Exception as e:
            try:
                replica.close()
            except Exception:
                pass
            raise RuntimeError(f"Couldn't prepare the shared database: {e}")

        # Only seed from this computer's local data if the remote is
        # genuinely empty of real content -- if it already has purchase
        # orders, suppliers, or user accounts, someone/something already
        # seeded it (this computer on an earlier attempt, or a colleague's
        # computer), and this computer's local copy should defer to that
        # shared truth rather than overwrite it.
        remote_has_data = False
        for table in ("purchase_orders", "suppliers", "users"):
            try:
                row = replica.execute(f"SELECT COUNT(*) AS c FROM {table}").fetchone()
                if row and row["c"]:
                    remote_has_data = True
                    break
            except Exception:
                pass

        if not remote_has_data:
            # Every table this computer's real data could possibly touch
            # already exists on the remote at this point (pushed on its own,
            # above, before _prepare_and_migrate_schema() ever ran) -- so
            # copying real rows in below and pushing them, further down,
            # never repeats the "CREATE TABLE + DML in the same batch"
            # failure shape this batch fixed, no matter how much data there
            # is or how many tables it spans.
            #
            # Deliberately not "with sqlite3.connect(...) as local:" --
            # Python's sqlite3.Connection context manager only commits/
            # rolls back the transaction on exit, it does NOT close the
            # connection or release its file handle. On Windows (unlike
            # Linux, where this bug went unnoticed through local testing)
            # that left DB_PATH still open when the code further down
            # tried to rename it out of the way, failing with "[WinError
            # 32] The process cannot access the file because it is being
            # used by another process" -- reported directly by Yitzi the
            # first time he actually tried this on his own machine.
            # Explicit try/finally + .close() below actually releases it.
            local = sqlite3.connect(str(DB_PATH))
            try:
                local.row_factory = sqlite3.Row
                table_names = [
                    r["name"] for r in local.execute(
                        "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'"
                    )
                ]
                for table in table_names:
                    try:
                        cols = [r["name"] for r in local.execute(f"PRAGMA table_info({table})")]
                    except Exception:
                        continue
                    if not cols:
                        continue
                    col_list = ", ".join(cols)
                    placeholders = ", ".join("?" for _ in cols)
                    try:
                        rows = local.execute(f"SELECT {col_list} FROM {table}").fetchall()
                    except Exception:
                        continue
                    for r in rows:
                        # OR REPLACE, not OR IGNORE: _prepare_and_migrate_schema
                        # just seeded fresh default rows into a couple of
                        # tables (settings, addresses, ...) via _seed_defaults --
                        # this computer's real saved values need to win over
                        # those generic placeholders, not get silently skipped
                        # because a row with the same key already "exists."
                        # Safe here specifically because remote_has_data is
                        # False -- there's no real shared content this could
                        # clobber.
                        try:
                            replica.execute(
                                f"INSERT OR REPLACE INTO {table} ({col_list}) VALUES ({placeholders})",
                                tuple(r[c] for c in cols),
                            )
                        except Exception as e:
                            log.warning("Skipped a row copying %s into the shared database: %s", table, e)
            finally:
                local.close()
            replica.commit()
            try:
                replica.push()
            except Exception as e:
                try:
                    replica.close()
                except Exception:
                    pass
                raise SharedDatabaseUnreachableError(
                    f"Connected, but couldn't upload this computer's existing data: {e}"
                )
            try:
                replica.pull()
            except Exception:
                pass
            local_po_count = 0
            local = sqlite3.connect(str(DB_PATH))
            try:
                try:
                    local_po_count = local.execute("SELECT COUNT(*) FROM purchase_orders").fetchone()[0]
                except Exception:
                    local_po_count = 0
            finally:
                local.close()
            remote_po_count = 0
            try:
                row = replica.execute("SELECT COUNT(*) AS c FROM purchase_orders").fetchone()
                remote_po_count = row["c"] if row else 0
            except Exception:
                remote_po_count = 0
            if local_po_count and remote_po_count < local_po_count:
                try:
                    replica.close()
                except Exception:
                    pass
                raise RuntimeError(
                    "The upload didn't seem to complete properly -- the shared database "
                    "doesn't show all of this computer's purchase orders yet. Nothing has "
                    "been changed on this computer, and the shared database wasn't turned "
                    "on. Your data and your backup are both safe -- please try again, or "
                    "get in touch if this keeps happening."
                )

        replica.close()

        # Only now -- everything above has already succeeded -- close the
        # shared global connection, right before the one step that
        # actually requires DB_PATH to have no open handles left on it.
        if _connection is not None:
            try:
                _connection.commit()
                _connection.close()
            except Exception:
                pass
            _connection = None

        try:
            preserved_original = APP_DIR / f"data.db.pre-shared-db-{datetime.now().strftime('%Y%m%d-%H%M%S')}.bak"
            if DB_PATH.exists():
                DB_PATH.replace(preserved_original)
            for companion in work_dir.iterdir():
                if companion.name == seed_path.name:
                    dest_name = DB_PATH.name
                elif companion.name.startswith(seed_path.name):
                    dest_name = DB_PATH.name + companion.name[len(seed_path.name):]
                else:
                    continue
                shutil.copy2(str(companion), str(APP_DIR / dest_name))
        except Exception:
            # The swap itself failed -- e.g. the rename couldn't get an
            # exclusive handle on DB_PATH. Windows' rename is all-or-
            # nothing, so DB_PATH should still be exactly what it was.
            # Re-opening an ordinary connection to it here (best effort --
            # if this also fails, the original exception below is still
            # what the caller sees) puts _connection back into a normal,
            # working state instead of leaving it cleared, so this
            # process's own next get_connection() call recovers cleanly.
            # The UI layer (_save_turso_config in po_generator_qt.py) is
            # what refreshes the already-open window's own conn reference
            # on this same failure -- this only fixes the module global.
            try:
                get_connection()
            except Exception:
                pass
            raise
    finally:
        shutil.rmtree(str(work_dir), ignore_errors=True)

    save_turso_config(url, auth_token)


def get_connection():
    global _connection
    if _connection is None:
        # Batch 160: the new backend-service shared-database mode --
        # deliberately checked FIRST and handled as its own separate
        # branch, before any of the local-file logic below (integrity
        # check, PRAGMA, schema migration) even runs. There is no local
        # SQLite file involved in this mode at all -- DB_PATH is never
        # touched, nothing here is a local replica of anything -- so none
        # of that machinery applies; see _RemoteConnection's own docstring
        # for exactly which of it (PRAGMA, executescript) it already
        # refuses/no-ops on its own as a second line of defence.
        #
        # The login itself -- which is what actually produces the session
        # token _RemoteConnection needs -- has to happen BEFORE this
        # function is ever called: po_core.py has no UI dependency and
        # can't show a password prompt itself. The UI layer's login flow
        # (see po_generator_qt.py's Batch 160+ shared-database login
        # dialog) is expected to have already called
        # remote_backend_login()/remote_backend_set_password() and handed
        # the resulting token to set_remote_session() before its first
        # call to get_connection(). If that hasn't happened -- a
        # programming error, not something a real user should ever
        # trigger -- this raises SharedDatabaseUnreachableError rather
        # than silently doing nothing or crashing on a None token.
        backend_cfg = load_backend_config()
        if backend_cfg is not None:
            if _remote_session is None or _remote_session["base_url"] != backend_cfg["base_url"]:
                raise SharedDatabaseUnreachableError(
                    "Not logged in to the shared database yet. Please log in and try again."
                )
            _connection = _RemoteConnection(_remote_session["base_url"], _remote_session["token"])
            return _connection
        # Corrupt-database detection happens before anything else touches
        # the file -- a damaged database can otherwise fail deep inside
        # executescript/_ensure_column with a raw sqlite3 error, or (worse)
        # silently succeed at creating a fresh empty schema on top of a
        # file SQLite can no longer read as the real one. Checked with its
        # own short-lived connection first so a bad file is never even
        # opened via the shared global one.
        ok, detail = check_db_integrity(DB_PATH)
        if not ok:
            log.error("Database failed integrity check: %s", detail)
            raise DatabaseCorruptError(detail)
        # Batch 109/110: if Yitzi's entered shared-database details in
        # Settings, open the SAME local file as a turso embedded replica
        # instead of a plain sqlite3 connection -- it already holds
        # everything that's been saved locally, so it becomes the seed of
        # the shared database rather than starting empty. Every existing
        # call below this point (executescript, _ensure_column, plain
        # .execute(...), row["col"] access, .commit()) keeps working
        # unchanged either way -- that's the whole point of
        # _PushOnCommitConnection and turso.Row lining up with sqlite3's
        # own API.
        #
        # Batch 110 correction: Batch 109 originally fell back to a plain
        # local sqlite3 connection here if the shared database couldn't be
        # reached, reasoning that a broken setup should never lock Yitzi
        # out of his own data. Yitzi (correctly) pointed out that's wrong
        # for a SHARED database specifically: once other people are also
        # working off the same live database, quietly working from a
        # stale local copy means whatever gets saved here is invisible to
        # everyone else and can silently diverge -- "it needs to all be
        # online, we can only be working off one database." So once
        # shared mode is configured, opening the database now genuinely
        # requires reaching it -- no silent local fallback at startup.
        # Anything that stops that (host unreachable, bad/expired token,
        # or this particular build not even including the 'turso'
        # package) raises SharedDatabaseUnreachableError instead, which
        # po_generator_qt.py's main() turns into a blocking dialog with
        # exactly two honest ways forward: try again, or explicitly (not
        # silently) switch this computer back to local-only via
        # clear_turso_config() -- a deliberate, visible choice, not a
        # quiet default. A momentary connection drop *after* the app is
        # already open and synced is handled differently (see
        # _PushOnCommitConnection.commit() below) -- that's a normal part
        # of using a live network connection, not a reason to block someone
        # who was already working.
        turso_cfg = load_turso_config()
        if turso_cfg is not None:
            if _turso is None:
                raise SharedDatabaseUnreachableError(
                    "This copy of the program doesn't include the component the shared "
                    "database needs (the 'turso' package). Update to a build that includes "
                    "it, or switch this computer back to using its local database only."
                )
            try:
                replica = _turso.sync.connect(
                    str(DB_PATH), remote_url=turso_cfg["url"], auth_token=turso_cfg["auth_token"]
                )
                replica.row_factory = _turso.Row
            except Exception as e:
                raise SharedDatabaseUnreachableError(
                    f"Couldn't connect to the shared database: "
                    f"{_diagnose_turso_error('connect', turso_cfg['url'], e)}"
                )
            try:
                _pull_with_retry(replica)
            except Exception as e:
                try:
                    replica.close()
                except Exception:
                    pass
                raise SharedDatabaseUnreachableError(
                    f"Couldn't sync with the shared database: "
                    f"{_diagnose_turso_error('pull', turso_cfg['url'], e)}"
                )
            _connection = _PushOnCommitConnection(replica)
        else:
            _connection = sqlite3.connect(str(DB_PATH), check_same_thread=False)
            _connection.row_factory = sqlite3.Row
        _connection.execute("PRAGMA foreign_keys = ON")
        # Batch 104: refuse to open a database a newer copy of the app has
        # already upgraded. PRAGMA user_version defaults to 0 on both a
        # brand new file and every pre-Batch-104 database already out
        # there, so existing installs are unaffected; it only ever fires
        # once a build with a higher APP_SCHEMA_VERSION has actually
        # opened this same file. Batch 125: the actual schema/migration
        # work now lives in _prepare_and_migrate_schema() (shared with
        # setup_shared_database()'s one-time seed replica) -- this just
        # handles the one bit that's specific to the real, global
        # connection: clearing it out before raising.
        try:
            _prepare_and_migrate_schema(_connection)
        except SchemaTooNewError:
            _connection = None
            raise
    return _connection


# ============================================================================
# Batch 103: user accounts, login, and per-user permissions
# ============================================================================

# The full set of permission keys the app understands, grouped by the same
# sections as the sidebar (NAV_ITEMS in po_generator_qt.py) so the Users
# admin screen can render one clearly labelled group per section instead of
# a flat, unsorted wall of checkboxes. Each entry is
# (key, section_label, action_label) -- "section_label" repeats on purpose
# for every row in a section (used to group rows, not just as trivia).
# Adding a new gated feature later means adding one line here and one
# `main.can("...")` check at the point that feature is shown/used -- see
# user_has_permission() below. This is Batch 103's first pass -- page-level
# access for every section in the app, plus the first slice of finer,
# button-level control (PO Manager's PDF export, deleting a PO, and Price
# Requests' view-vs-create split) -- not yet every single button in the
# app; see FEATURE_LOG.md's Batch 103 entry for what's still to come.
PERMISSIONS = [
    ("dashboard.view", "Dashboard", "View the Dashboard"),
    ("new_po.view", "New PO", "Create purchase orders"),
    # Batch 107: New PO's own "save" vs "send" split -- Save PO/Save
    # draft/Generate PDF never leave the building, Open in Outlook/Copy for
    # Outlook do -- kept separate per Yitzi's own new_po example.
    ("new_po.save", "New PO", "Save a draft, save as placed, or generate a PDF"),
    ("new_po.send", "New PO", "Send a PO to Outlook (open or copy)"),
    ("po_manager.view", "PO Manager", "View the PO Manager list"),
    ("po_manager.delete", "PO Manager", "Delete / restore purchase orders"),
    ("po_manager.export_pdf", "PO Manager", "Export / open a PO as PDF"),
    # Batch 107: everything else PO Manager (and its Quick View popup) can
    # do to an already-saved PO beyond deleting/exporting it.
    ("po_manager.edit", "PO Manager", "Edit, duplicate, or adjust the quantity on a saved PO"),
    ("po_manager.set_status", "PO Manager", "Change a PO's status"),
    ("po_manager.flag", "PO Manager", "Flag / unflag a PO for review"),
    ("po_manager.resend", "PO Manager", "Resend a PO to its supplier"),
    ("reports.view", "Reports", "View Reports"),
    ("reports.export", "Reports", "Export a report's rows to CSV"),
    ("reports.send_email", "Reports", "Pull/generate a report or stock recap email"),
    ("suppliers.view", "Suppliers", "View and manage suppliers"),
    # Batch 107: add/edit vs delete vs import vs merge/dedupe vs credit vs
    # issue-logging -- the finer supplier actions Yitzi asked for by name.
    ("suppliers.edit", "Suppliers", "Add / edit a supplier"),
    ("suppliers.delete", "Suppliers", "Delete a supplier"),
    ("suppliers.import", "Suppliers", "Import suppliers from CSV"),
    ("suppliers.merge", "Suppliers", "Find duplicates / clean up supplier naming"),
    ("suppliers.manage_credit", "Suppliers", "View and reset a supplier's credit status"),
    ("suppliers.log_issue", "Suppliers", "Log, edit, or delete a supplier issue"),
    ("scorecards.view", "Compare Suppliers", "View supplier scorecards / comparison"),
    ("addresses.view", "Addresses", "View and manage delivery addresses"),
    ("addresses.edit", "Addresses", "Add / edit a delivery address"),
    ("addresses.delete", "Addresses", "Delete a delivery address"),
    ("products.view", "Products", "View the product catalog"),
    ("products.edit", "Products", "Add / edit a product"),
    ("products.delete", "Products", "Delete a product"),
    ("products.merge", "Products", "Find duplicates / merge products"),
    ("price_requests.view", "Price Requests", "View price requests"),
    ("price_requests.create", "Price Requests", "Create / send price requests"),
    # Batch 107: "respond" covers every way a request's own details/replies
    # get entered or changed (enter price, mark unavailable, undo, add more
    # suppliers, edit note, rename product) -- one key, not six, since
    # they're all facets of the same "work this request" action on the
    # same detail screen. Export/delete stay their own keys, matching
    # Yitzi's own examples for this page.
    ("price_requests.respond", "Price Requests", "Enter/edit supplier replies on a price request"),
    ("price_requests.delete", "Price Requests", "Delete a price request"),
    ("price_requests.export", "Price Requests", "Copy / send a price request's pricing summary"),
    ("savings.view", "Savings", "View and log savings"),
    ("savings.record", "Savings", "Record a saving"),
    ("savings.manage", "Savings", "Edit / delete a saving"),
    ("import_export.view", "Import / Export", "Zoho import / export"),
    ("import_export.import", "Import / Export", "Import from Zoho (POs, products, or suppliers)"),
    ("import_export.export", "Import / Export", "Export to Zoho (POs or products)"),
    ("settings.view", "Settings & Backup", "Open Settings & Backup"),
    ("settings.manage_users", "Settings & Backup", "Manage user accounts (Users tab)"),
    # Batch 107: one key per remaining Settings tab, same precedent as
    # settings.manage_users -- someone without a given key never sees that
    # tab at all, matching "every setting... configurable per user."
    ("settings.manage_company", "Settings & Backup", "Company tab"),
    ("settings.manage_defaults", "Settings & Backup", "Your defaults tab"),
    ("settings.manage_dashboard", "Settings & Backup", "Dashboard tab"),
    ("settings.manage_pdf", "Settings & Backup", "PDF & Branding tab"),
    ("settings.manage_reports", "Settings & Backup", "Email & Reports tab"),
    ("settings.manage_zoho", "Settings & Backup", "Zoho Export tab"),
    ("settings.manage_savings_scorecards", "Settings & Backup", "Savings & Scorecards tab"),
    ("settings.manage_purchase_owners", "Settings & Backup", "Purchase Owners tab"),
    ("settings.manage_backup", "Settings & Backup", "Data & Backup tab (includes Stock Sync folder)"),
    # Batch 121: who gets which in-app/Windows-popup notification, kept as
    # its own key rather than folded into manage_reports -- notification
    # routing isn't really an email/report setting, and someone might
    # reasonably be trusted with one but not the other.
    ("settings.manage_notifications", "Settings & Backup", "Notifications tab (who gets pinged for what)"),
    # Batch 106: the stock team's receiving screen, folded in from the
    # formerly-separate standalone companion program -- see stock_sync.py.
    ("stock.view", "Stock", "View the Stock / receiving page"),
    ("stock.mark_received", "Stock", "Mark items received, partially received, or returned"),
    ("stock.send_report", "Stock", "Send the outstanding orders report"),
    # Batch 145: the stock-to-supply product request feature -- asking
    # Supply to order more of something, or flagging a product that isn't
    # in the catalogue yet. Same view/create split as Price Requests;
    # what Supply does to resolve a request (mark ordered, add to a PO,
    # link/approve a flagged new item) gets its own key in a later batch
    # once that part of the page exists.
    ("stock_requests.view", "Stock Requests", "View stock requests"),
    ("stock_requests.create", "Stock Requests", "Ask Supply to order more of something, or flag a new product"),
    # Batch 147: the Supply side of the same page -- mark a request
    # ordered, or bundle several of them (sharing a supplier) onto one new
    # PO at once. Kept as one key rather than splitting mark-ordered from
    # add-to-PO, same "these are all facets of the same 'work this
    # request' action" reasoning price_requests.respond already used.
    # Batch 148 folds the flagged-new-item dedupe/approval action into this
    # same key too, for the same reason.
    ("stock_requests.resolve", "Stock Requests", "Mark stock requests ordered, add them to a new PO, or resolve a flagged new item"),
]
PERMISSION_KEYS = {key for key, _, _ in PERMISSIONS}
PERMISSION_SECTIONS = list(dict.fromkeys(section for _, section, _ in PERMISSIONS))  # first-seen order

# Batch 108: starting points for Settings > Users > Permissions, not a
# separate enforcement mechanism -- applying one just ticks/unticks the same
# checkboxes a person could tick by hand, and every box stays individually
# editable afterward. Modelled directly on how Yitzi actually described his
# three real accounts: himself (full access -- handled by the existing
# is_admin bypass, not a permission list at all, so there's no "owner" entry
# here), the stock team (view POs and the Stock page, but no report pulling
# or PO creation), and finance/the CEO (reports and PO visibility, but no
# stock or PO-creation access). Keys are validated against PERMISSION_KEYS
# by whatever UI applies one, so a future renamed/removed key can't silently
# leave a stale entry in a preset.
ROLE_PRESETS = [
    ("stock_team", "Stock team", [
        "dashboard.view",
        "po_manager.view", "po_manager.export_pdf",
        "stock.view", "stock.mark_received", "stock.send_report",
        "stock_requests.view", "stock_requests.create",
    ]),
    ("finance_reports", "Finance / Reports (CEO)", [
        "dashboard.view",
        "po_manager.view",
        "reports.view", "reports.export", "reports.send_email",
        "suppliers.view",
        "savings.view",
    ]),
    # Batch 147: the Supply side's own preset, alongside stock_team -- adds
    # the ordinary product-ordering permissions (new PO, PO manager, price
    # requests, the supplier/product catalog) plus this batch's own
    # stock_requests.resolve, but deliberately NOT stock_requests.create --
    # raising a request is Stock's own action; nothing stops manually
    # granting it to a Supply account too, this is just the sensible
    # starting point.
    ("supply_team", "Supply team", [
        "dashboard.view",
        "new_po.view", "new_po.save", "new_po.send",
        "po_manager.view", "po_manager.edit", "po_manager.export_pdf", "po_manager.resend",
        "suppliers.view",
        "products.view", "products.edit",
        "price_requests.view", "price_requests.create", "price_requests.respond",
        "stock_requests.view", "stock_requests.resolve",
    ]),
]

# Batch 119: companion to ROLE_PRESETS above, for the Dashboard widget list
# rather than permissions -- Yitzi: "i want to configure what widgets they
# see ... maybe we can add more options as well as a finance or stock
# person need different widgets." Same "starting point, not an enforcement
# mechanism" rule as ROLE_PRESETS itself: applying one just fills in that
# account's dashboard_widgets (see set_user_dashboard_widgets), which stays
# a fully independent, individually-editable list from then on -- nothing
# here is re-applied automatically later, and picking a permissions preset
# in Settings > Users applies both this and ROLE_PRESETS' own permission
# keys together as one convenience action. A preset key with no entry here
# (there isn't one currently, but a future ROLE_PRESETS addition might not
# need one) just means "don't touch the widget list" when that preset's
# applied -- PermissionsDialog._apply_preset only pre-fills widgets for a
# preset that actually has an entry in this dict.
DASHBOARD_WIDGET_ROLE_DEFAULTS = {
    # The stock team cares about what's moving in and out, not spend
    # totals or price-request chasing.
    "stock_team": [
        "kpi_orders_due_in", "kpi_orders_overdue", "kpi_outstanding_value",
        "recent_card", "flagged_card",
    ],
    # Finance/the CEO cares about spend and value, not day-to-day receiving.
    "finance_reports": [
        "kpi_total_pos", "kpi_active_pos", "kpi_this_month", "kpi_last_month",
        "kpi_avg_order_value", "kpi_alerts", "recent_card", "alerts_card",
    ],
}

# ============================================================
# 8b. In-app notifications (bell icon) -- Batch 121.
# Yitzi: "i want there to be notifications going back and forth between the
# apps ... maybe we need to make a notification icon with a bell or number
# of notifications so its not messy ... there should be a setting to manage
# notifications who gets what." and later, on the routing idea itself: "yes"
# (confirming the Stock/Supply account role field) plus "i want in settings
# to be able to decide which people receive what notifications and when."
#
# Design: NOTIFICATION_TYPES is the fixed catalogue of events the app can
# raise. Each has a default_roles tuple used the first time Settings >
# Notifications is opened for it (before anyone's touched the routing at
# all) -- "stock"/"supply" match users.notification_role, "admin" means
# every is_admin account regardless of that field (Yitzi himself, day one,
# with no setup required). get_notification_routes()/set_notification_
# routes() store the actual (possibly customized) routing as one JSON blob
# in a single settings row -- same "one settings key, not dozens of flat
# ones" choice as price_request_email_subject/_body, since this is a
# genuinely structured, per-type record rather than a handful of independent
# scalars. A type missing from the stored JSON (a brand new type added in a
# future batch, or an install that's never opened the Notifications tab)
# falls back to its own default_roles with in-app on and Windows popups on
# -- so nothing needs migrating just because NOTIFICATION_TYPES grows.
NOTIFICATION_TYPES = [
    ("stock_request_item", "Stock requested an item to order", ("supply",)),
    ("stock_request_new_item", "Stock flagged a product not in the catalogue yet", ("supply",)),
    ("stock_request_ordered", "Supply marked a requested item ordered", ("stock",)),
    ("stock_request_new_item_resolved", "Supply linked or approved a flagged new item", ("stock",)),
    ("supplier_issue_logged", "An issue was logged on Receiving", ("admin",)),
]
NOTIFICATION_TYPE_KEYS = {key for key, _, _ in NOTIFICATION_TYPES}
NOTIFICATION_ROLE_LABELS = {"stock": "Stock team", "supply": "Supply team", "admin": "Admins"}


def get_notification_routes(conn):
    """Every NOTIFICATION_TYPES key, always -- each mapped to
    {"roles": [...], "user_ids": [...], "inapp": bool, "windows": bool}.
    roles/user_ids are added together (a union), not narrowed by each
    other, so "everyone tagged Supply, plus Dave specifically" is a normal
    thing to configure, not a contradiction."""
    raw = get_setting(conn, "notification_routes", "")
    stored = {}
    if raw:
        try:
            stored = json.loads(raw)
            if not isinstance(stored, dict):
                stored = {}
        except (TypeError, ValueError):
            stored = {}
    out = {}
    for key, _label, default_roles in NOTIFICATION_TYPES:
        entry = stored.get(key) or {}
        out[key] = {
            "roles": [r for r in entry.get("roles", list(default_roles)) if r in NOTIFICATION_ROLE_LABELS],
            "user_ids": [int(u) for u in entry.get("user_ids", []) if str(u).isdigit()],
            "inapp": bool(entry.get("inapp", True)),
            "windows": bool(entry.get("windows", True)),
        }
    return out


def set_notification_routes(conn, routes):
    """Overwrites the whole routing table. Silently drops any key that
    isn't a real notification type and any role that isn't one of
    stock/supply/admin, so a stale/typo'd key from a future rename can
    never linger in here forever."""
    clean = {}
    for key, entry in (routes or {}).items():
        if key not in NOTIFICATION_TYPE_KEYS:
            continue
        entry = entry or {}
        clean[key] = {
            "roles": [r for r in entry.get("roles", []) if r in NOTIFICATION_ROLE_LABELS],
            "user_ids": [int(u) for u in entry.get("user_ids", []) if str(u).isdigit()],
            "inapp": bool(entry.get("inapp", True)),
            "windows": bool(entry.get("windows", True)),
        }
    set_settings(conn, {"notification_routes": json.dumps(clean)})


def _notification_recipients(conn, type_key):
    """(user_ids, inapp, windows) for one notification type -- user_ids is
    every active account that matches the routed roles, unioned with any
    explicitly-added accounts, deduplicated. Returns ([], False, False) for
    an unrecognized type_key rather than raising, since a caller passing a
    typo'd key should just quietly notify nobody, not crash whatever
    action triggered it."""
    if type_key not in NOTIFICATION_TYPE_KEYS:
        return [], False, False
    route = get_notification_routes(conn).get(type_key, {})
    roles = set(route.get("roles", []))
    ids = set(route.get("user_ids", []))
    if roles:
        for u in conn.execute("SELECT id, is_admin, notification_role FROM users WHERE active=1").fetchall():
            if "admin" in roles and u["is_admin"]:
                ids.add(u["id"])
            role = u["notification_role"] or ""
            if role in roles:
                ids.add(u["id"])
            elif not role and ("stock" in roles or "supply" in roles):
                # Batch 156 (Task #106): Yitzi -- "not sure that notifications
                # are working between users requests such as when a new
                # product is requested". notification_role defaults to '' and
                # is only ever set through Settings' separate "Notification
                # team" dropdown, completely disconnected from a user's real
                # permissions -- so a brand-new account, or anyone nobody's
                # gotten around to tagging yet, matched nothing at all here
                # even though they already have the exact stock_requests.*
                # permission this notification is about. Fall back to
                # inferring stock/supply membership from real permissions
                # whenever the manual tag is still blank, so a fresh install
                # (or an untagged user) is routed sensibly by default. This
                # only runs when notification_role is blank, so an admin who
                # HAS set the dropdown for someone always wins outright --
                # this never overrides an explicit choice, only fills the gap
                # where nobody's made one yet.
                #
                # Deliberately checks the raw permission set rather than
                # going through user_has_permission() -- that helper treats
                # is_admin as "yes" to every key unconditionally (by design,
                # for the UI's own permission gating), which would silently
                # sweep every untagged admin into "supply"/"stock" here too.
                # An admin already gets routed via the "admin" in roles
                # check above when a type is actually routed to admins; this
                # branch is only about a real, specifically-granted
                # stock_requests.* permission, nothing broader.
                perms = (get_user(conn, u["id"]) or {}).get("permissions") or set()
                if "supply" in roles and "stock_requests.resolve" in perms:
                    ids.add(u["id"])
                elif "stock" in roles and "stock_requests.create" in perms:
                    ids.add(u["id"])
    return sorted(ids), route.get("inapp", True), route.get("windows", True)


def create_notification(conn, user_id, type_key, title, body="", link_page="", link_id=""):
    conn.execute(
        "INSERT INTO notifications(user_id, type_key, title, body, link_page, link_id, created_at) "
        "VALUES (?, ?, ?, ?, ?, ?, ?)",
        (user_id, type_key, title, body, link_page or "", str(link_id) if link_id else "", now_iso()),
    )
    conn.commit()


def notify_event(conn, type_key, title, body="", link_page="", link_id="", exclude_user_id=None):
    """Raises a notification of type_key for whoever it's routed to right
    now (see _notification_recipients). exclude_user_id leaves out the
    person who just caused the event themselves -- e.g. Supply marking
    their own request ordered shouldn't notify Supply-role accounts about
    it, only the Stock side. Returns the list of user_ids actually
    notified (empty if the type's routing has in-app notifications turned
    off, or matches nobody). Windows popups aren't shown here -- see
    pop_due_toasts() for why that happens locally, per running instance,
    instead of at the moment the event occurs."""
    ids, inapp, _windows = _notification_recipients(conn, type_key)
    if exclude_user_id is not None:
        ids = [i for i in ids if i != exclude_user_id]
    if not inapp or not ids:
        return []
    for uid in ids:
        create_notification(conn, uid, type_key, title, body, link_page, link_id)
    return ids


def list_notifications(conn, user_id, unread_only=False, limit=50):
    q = "SELECT * FROM notifications WHERE user_id=?"
    if unread_only:
        q += " AND (read_at IS NULL OR read_at='')"
    q += " ORDER BY datetime(created_at) DESC LIMIT ?"
    return [dict(r) for r in conn.execute(q, (user_id, limit)).fetchall()]


def unread_notification_count(conn, user_id):
    row = conn.execute(
        "SELECT COUNT(*) c FROM notifications WHERE user_id=? AND (read_at IS NULL OR read_at='')",
        (user_id,),
    ).fetchone()
    return row["c"] if row else 0


def mark_notification_read(conn, notification_id):
    conn.execute(
        "UPDATE notifications SET read_at=? WHERE id=? AND (read_at IS NULL OR read_at='')",
        (now_iso(), notification_id),
    )
    conn.commit()


def mark_all_notifications_read(conn, user_id):
    conn.execute(
        "UPDATE notifications SET read_at=? WHERE user_id=? AND (read_at IS NULL OR read_at='')",
        (now_iso(), user_id),
    )
    conn.commit()


def pop_due_toasts(conn, user_id):
    """Called periodically (MainWindow's existing 2-minute due-check timer,
    via refresh_all) by whichever running copy of the app is actually
    logged in as user_id -- there is no way for the account that CREATED a
    notification to pop a toast on a different person's machine, since
    each person runs their own copy of this app on their own PC and only
    shares the database/sync folder, not a live connection to each other's
    screen. So instead, every running instance checks its own logged-in
    user's still-unshown rows on each tick and shows the toast locally,
    right here. Re-checks each row's type's current "windows" routing
    flag at pop time (not whatever it was when the row was created), so
    turning Windows popups off for a type in Settings immediately stops
    new toasts for it without needing to touch already-queued rows."""
    if not user_id:
        return
    routes = get_notification_routes(conn)
    rows = conn.execute(
        "SELECT * FROM notifications WHERE user_id=? AND windows_shown=0 ORDER BY datetime(created_at)",
        (user_id,),
    ).fetchall()
    shown_ids = []
    for row in rows:
        route = routes.get(row["type_key"], {})
        if route.get("windows", True):
            _show_windows_toast(row["title"], row["body"] or "")
        shown_ids.append(row["id"])
    if shown_ids:
        conn.executemany("UPDATE notifications SET windows_shown=1 WHERE id=?", [(i,) for i in shown_ids])
        conn.commit()


# ============================================================
# 8c. Stock-to-supply product requests -- Batch 145.
# Yitzi: "where we up to with the other changes, like stock request" plus
# the original Batch 121 request that put NOTIFICATION_TYPES/stock_request_*
# settings in place ahead of this. Two request "kinds": Stock asking Supply
# to order more of an existing catalogue product ("existing"), or flagging
# a product that isn't in the catalogue at all yet ("new"). Both raise the
# matching notify_event() type already routed to Supply-role accounts since
# Batch 121; the reverse direction (Supply marking a request ordered, or
# resolving a flagged new item) is a later batch's job and will call
# notify_event() with the other two already-seeded types
# (stock_request_ordered / stock_request_new_item_resolved) the same way.
PRODUCT_REQUEST_KIND_LABELS = {"existing": "Reorder", "new": "New item"}
PRODUCT_REQUEST_STATUSES = ["open", "ordered", "resolved", "cancelled"]


def create_product_request(conn, kind, requested_by, product_id=0, product_name="", supplier_id=0, qty=0, notes=""):
    """kind must be 'existing' or 'new' -- anything else is treated as
    'existing', same defensive-default pattern as set_user_notification_role.
    product_name is required either way: for 'existing' the caller passes
    the catalogue product's current name (a snapshot, not a live FK lookup
    later -- see the product_requests table's own schema comment for why),
    for 'new' it's whatever free-text name Stock typed. Raises the right
    notify_event() type (routed to Supply-role accounts by default, same
    routing Batch 121 already set up) and excludes the requester themselves
    from it, since they obviously already know they just made the request."""
    kind = kind if kind in ("existing", "new") else "existing"
    product_name = (product_name or "").strip()
    now = now_iso()
    cur = conn.execute(
        "INSERT INTO product_requests(kind, product_id, product_name, supplier_id, qty, notes, "
        "requested_by, status, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, 'open', ?)",
        (kind, product_id or 0, product_name, supplier_id or 0, qty or 0, notes or "", requested_by, now),
    )
    conn.commit()
    request_id = cur.lastrowid
    if kind == "new":
        title = f"Stock flagged a product not in the catalogue: {product_name}"
        type_key = "stock_request_new_item"
    else:
        title = f"Stock requested more stock: {product_name}"
        type_key = "stock_request_item"
    body = f"Qty: {qty}" if qty else ""
    if notes:
        body = f"{body} -- {notes}" if body else notes
    notify_event(conn, type_key, title, body, link_page="stock_requests", link_id=request_id, exclude_user_id=requested_by)
    return request_id


# Batch 147: shared by list_product_requests()/get_product_request() so the
# two never drift apart. Beyond the requester's name and preferred-supplier
# name, this also works out an effective_supplier_id/supplier_name -- a
# request can have its own supplier_id=0 ("no preference") while its linked
# catalogue product (product_id, for the "existing" kind) DOES have a real
# supplier -- COALESCE(NULLIF(pr.supplier_id, 0), pp.supplier_id, 0) falls
# back to that product's supplier whenever the request itself didn't pin
# one down. This is what the Supply-side "Add to PO" bulk action (Batch
# 147) groups selected requests by, and what its resulting PO's supplier
# gets set to -- a request with neither a supplier of its own nor a linked
# product with one lands at 0, which "Add to PO" treats as "needs a
# supplier chosen first," not a silent guess. product_last_price rides
# along too, purely as a convenience starting price for that same PO-
# building step -- nothing here decides an order's actual price.
_PRODUCT_REQUEST_SELECT_SQL = (
    "SELECT pr.*, u.full_name AS requester_name, "
    "COALESCE(NULLIF(pr.supplier_id, 0), pp.supplier_id, 0) AS effective_supplier_id, "
    "COALESCE(s.company_name, ps.company_name) AS supplier_name, "
    "pp.last_price AS product_last_price "
    "FROM product_requests pr "
    "LEFT JOIN users u ON u.id = pr.requested_by "
    "LEFT JOIN suppliers s ON s.id = pr.supplier_id "
    "LEFT JOIN products pp ON pp.id = pr.product_id "
    "LEFT JOIN suppliers ps ON ps.id = pp.supplier_id"
)


def list_product_requests(conn, status=None, kind=None):
    """Every product request, newest first, each carrying the requester's
    display name plus the effective-supplier/last-price columns
    _PRODUCT_REQUEST_SELECT_SQL documents -- joined in here rather than
    left for the UI to look up per-row, same convenience
    list_price_requests() and list_products() already give their own
    callers."""
    sql = _PRODUCT_REQUEST_SELECT_SQL + " WHERE 1=1"
    params = []
    if status:
        sql += " AND pr.status = ?"
        params.append(status)
    if kind:
        sql += " AND pr.kind = ?"
        params.append(kind)
    # Tie-broken by id DESC as well as created_at -- two requests raised in
    # the same second (easily done from this same dialog, or in a test)
    # would otherwise sort in whatever order SQLite happens to return
    # matching rows, not necessarily newest-first.
    sql += " ORDER BY datetime(pr.created_at) DESC, pr.id DESC"
    return [dict(r) for r in conn.execute(sql, params)]


def get_product_request(conn, request_id):
    row = conn.execute(_PRODUCT_REQUEST_SELECT_SQL + " WHERE pr.id = ?", (request_id,)).fetchone()
    return dict(row) if row else None


def cancel_product_request(conn, request_id):
    """Withdraws a still-open request -- a no-op (no row changes) if it's
    already past 'open', so a stale/double click can't un-cancel an
    already-ordered or already-resolved request or resurrect a cancelled
    one. The UI is expected to only ever offer this for the requester's
    own open rows (or an admin's); this function itself doesn't check who's
    calling, same division of responsibility as the rest of this module."""
    conn.execute("UPDATE product_requests SET status='cancelled' WHERE id=? AND status='open'", (request_id,))
    conn.commit()


def mark_product_request_ordered(conn, request_id, marked_by=None):
    """Batch 147: Supply's own "yes, this is being ordered" action -- a
    no-op on anything not currently 'open', same defensive rule
    cancel_product_request() already follows, so a stale/double click can't
    un-cancel a cancelled request or re-stamp an already-ordered one's
    ordered_at. Used both standalone (Supply already placed the order some
    other way -- by phone, already on an existing PO, etc.) and as the
    second half of the "Add to PO" bulk action once that's prefilled a new
    PO draft for review -- see StockRequestsPage._add_selected_to_po's own
    docstring for why marking happens at that point rather than waiting for
    the draft to actually be saved. Raises stock_request_ordered (routed to
    Stock-role accounts since Batch 121) so the original requester finds
    out without needing to keep checking back on this page themselves;
    marked_by excludes whoever clicked it, same exclude_user_id use
    create_product_request() already makes for the forward direction."""
    cur = conn.execute(
        "UPDATE product_requests SET status='ordered', ordered_at=? WHERE id=? AND status='open'",
        (now_iso(), request_id),
    )
    conn.commit()
    if cur.rowcount:  # only the call that actually flipped open -> ordered notifies -- a repeat click is silent
        req = get_product_request(conn, request_id)
        if req:
            notify_event(
                conn, "stock_request_ordered",
                f"Your stock request for {req['product_name']} is being ordered",
                "", link_page="stock_requests", link_id=request_id, exclude_user_id=marked_by,
            )


def resolve_flagged_product_request(conn, request_id, resolution, resolved_by=None, matched_product_id=0,
                                     new_product_supplier_id=0, new_product_code="", new_product_price=0):
    """Batch 148: the dedupe/approval half of a Stock-flagged "not in the
    catalogue" request -- Supply's answer to "is this actually something we
    already stock under a different name, or is it genuinely new." Only
    ever acts on a still-open 'new' kind request; anything else (already
    resolved, cancelled, or an 'existing' kind request -- there's nothing
    to dedupe there, it was already linked to a real product at creation)
    is a silent no-op, same defensive pattern cancel_product_request()/
    mark_product_request_ordered() already use, and returns False so the
    caller can tell nothing happened.

    resolution='matched' links the request straight to matched_product_id
    (an existing catalogue product Supply picked by hand, or the one
    someone chose out of ResolveFlaggedItemDialog's own close-match
    suggestions -- see that dialog's docstring for how it reuses the exact
    same find_close_matching_products()/CloseMatchWarningDialog pair the
    Products page's own "Add product" dedupe check already established,
    rather than this function inventing a second dedupe UI).
    resolution='new' actually creates the catalogue entry via the existing
    upsert_product_manual(), using the request's own product_name verbatim
    as the new product's name (that's the whole point -- it's the text
    Stock flagged in the first place) with whatever supplier/code/price
    Supply entered. upsert_product_manual() itself doesn't hand back the
    row it just touched, so the new/matched id is looked up straight
    afterward by exact (case-insensitive) name under that supplier -- same
    lookup-after-the-fact NewProductRequestDialog._send() already relies on
    to resolve a typed name back to a real product id.

    Either way, raises the already-seeded stock_request_new_item_resolved
    notification (routed to Stock-role accounts since Batch 121) so the
    original requester finds out what happened to the product they
    flagged, excluding whoever resolved it."""
    req = get_product_request(conn, request_id)
    if not req or req["kind"] != "new" or req["status"] != "open":
        return False
    if resolution == "matched":
        if not matched_product_id:
            return False
        matched = conn.execute("SELECT name FROM products WHERE id=?", (matched_product_id,)).fetchone()
        if not matched:
            return False
        product_id = matched_product_id
        note = f"Matched to the existing catalogue product '{matched['name']}'."
    elif resolution == "new":
        name = req["product_name"]
        upsert_product_manual(conn, None, new_product_supplier_id, new_product_code, name, new_product_price)
        candidates = list_products(conn, exact_supplier_id=new_product_supplier_id, query=name)
        found = next((p for p in candidates if p["name"].strip().lower() == name.strip().lower()), None)
        product_id = found["id"] if found else 0
        note = "Approved as a new catalogue product."
    else:
        return False
    cur = conn.execute(
        "UPDATE product_requests SET status='resolved', product_id=?, resolved_at=?, resolved_by=?, resolved_note=? "
        "WHERE id=? AND status='open'",
        (product_id, now_iso(), resolved_by or 0, note, request_id),
    )
    conn.commit()
    if cur.rowcount:
        notify_event(
            conn, "stock_request_new_item_resolved",
            f"Your flagged product '{req['product_name']}' has been resolved",
            note, link_page="stock_requests", link_id=request_id, exclude_user_id=resolved_by,
        )
        return True
    return False


def set_user_notification_role(conn, user_id, role):
    """role must be 'stock', 'supply', or '' (blank/neither) -- anything
    else is silently treated as blank, same defensive pattern as
    set_user_dashboard_widgets' attr filtering."""
    role = role if role in ("stock", "supply") else ""
    conn.execute("UPDATE users SET notification_role=? WHERE id=?", (role, user_id))
    conn.commit()


def hash_password(password, salt=None):
    """PBKDF2-HMAC-SHA256, stdlib only -- no bcrypt/argon2 dependency to
    add for a small in-house app. Returns (salt_hex, hash_hex); a fresh
    random salt is generated unless one's passed in (re-hashing to check a
    login attempt against an already-stored salt)."""
    salt = salt or secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), bytes.fromhex(salt), 200_000)
    return salt, digest.hex()


def verify_password(password, salt, expected_hash):
    _, actual_hash = hash_password(password, salt)
    return secrets.compare_digest(actual_hash, expected_hash)


def _row_to_user(row, conn=None, permissions=None):
    d = dict(row)
    d.pop("password_hash", None)
    d.pop("password_salt", None)
    if permissions is not None:
        d["permissions"] = permissions
    elif conn is not None:
        d["permissions"] = {r["permission_key"] for r in conn.execute(
            "SELECT permission_key FROM user_permissions WHERE user_id=?", (d["id"],)
        )}
    return d


def count_users(conn):
    return conn.execute("SELECT COUNT(*) c FROM users").fetchone()["c"]


def list_users(conn, include_inactive=True):
    sql = "SELECT * FROM users"
    if not include_inactive:
        sql += " WHERE active=1"
    sql += " ORDER BY full_name COLLATE NOCASE, username COLLATE NOCASE"
    return [_row_to_user(r, conn) for r in conn.execute(sql)]


def get_user(conn, user_id):
    row = conn.execute("SELECT * FROM users WHERE id=?", (user_id,)).fetchone()
    return _row_to_user(row, conn) if row else None


def get_user_by_username(conn, username):
    row = conn.execute("SELECT * FROM users WHERE username=? COLLATE NOCASE", (username,)).fetchone()
    return _row_to_user(row, conn) if row else None


def _unusable_placeholder_password_hash():
    """A genuinely random password hash nobody knows or could practically
    guess -- used for both a brand new "must set password" account and a
    just-reset one (see create_user()/reset_user_password() below), so
    the account is actually unusable, not just supposed to be, until a
    real password gets set. Never logged, shown, or stored anywhere else."""
    return hash_password(secrets.token_hex(32))


def create_user(conn, username, password=None, full_name="", is_admin=False, permissions=None, active=True):
    """Creates a new account. Raises ValueError on a blank username, an
    empty (but explicitly given) password, or a username already taken
    (case-insensitively -- "Yitzi" and "yitzi" would otherwise be two
    confusingly different logins).

    Batch 114: password is now optional -- pass None (the default; every
    pre-Batch-114 caller in this codebase still passes a real one
    explicitly, so nothing about them changes) to create the account in
    "must set password" state instead of picking a password for them.
    That's what Settings > Users > "Add user" does now: Yitzi no longer
    invents a new teammate's password, they choose their own the first
    time they log in (see LoginDialog._try_login/CreatePasswordDialog in
    po_generator_qt.py). A blank string, as opposed to None, is still
    treated as a mistake and raises -- only explicitly asking for the new
    account to start passwordless (None) gets that treatment."""
    username = (username or "").strip()
    full_name = (full_name or "").strip()
    if not username:
        raise ValueError("Username is required.")
    must_set_password = password is None
    if password is not None and not password:
        raise ValueError("Password is required.")
    if get_user_by_username(conn, username):
        raise ValueError(f"The username '{username}' is already taken.")
    if must_set_password:
        salt, pw_hash = _unusable_placeholder_password_hash()
    else:
        salt, pw_hash = hash_password(password)
    now = now_iso()
    cur = conn.execute(
        "INSERT INTO users(username, full_name, password_hash, password_salt, is_admin, active, must_set_password, created_at, updated_at) "
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
        (username, full_name, pw_hash, salt, 1 if is_admin else 0, 1 if active else 0, 1 if must_set_password else 0, now, now),
    )
    user_id = cur.lastrowid
    for key in (permissions or []):
        if key in PERMISSION_KEYS:
            # Batch 157: "ON CONFLICT ... DO NOTHING" instead of SQLite's own
            # "INSERT OR IGNORE" shorthand -- both mean the same thing here
            # (user_permissions' own PRIMARY KEY is (user_id, permission_key),
            # so this is a plain "skip it if already there") but ON CONFLICT
            # is the one spelling both SQLite (3.24+) and PostgreSQL actually
            # understand identically, part of moving the shared-database
            # feature off Turso onto a real hosted Postgres -- see the Batch
            # 157 FEATURE_LOG entry for the full why.
            conn.execute(
                "INSERT INTO user_permissions(user_id, permission_key) VALUES (?, ?) "
                "ON CONFLICT (user_id, permission_key) DO NOTHING",
                (user_id, key),
            )
    conn.commit()
    return get_user(conn, user_id)


def _count_active_admins(conn, exclude_user_id=None):
    sql = "SELECT COUNT(*) c FROM users WHERE is_admin=1 AND active=1"
    params = []
    if exclude_user_id is not None:
        sql += " AND id != ?"
        params.append(exclude_user_id)
    return conn.execute(sql, params).fetchone()["c"]


def update_user(conn, user_id, username=None, full_name=None, is_admin=None, active=None):
    """Edits an existing account's details. Username, if given, must stay
    unique; refuses to remove admin status from, or deactivate, the very
    last active admin account -- there must always be at least one account
    that can get back into Settings > Users to fix a mistake, so nobody
    can ever lock the whole team out by accident."""
    user = conn.execute("SELECT * FROM users WHERE id=?", (user_id,)).fetchone()
    if not user:
        raise ValueError("No such user.")
    updates = {}
    if username is not None:
        username = username.strip()
        if not username:
            raise ValueError("Username is required.")
        existing = get_user_by_username(conn, username)
        if existing and existing["id"] != user_id:
            raise ValueError(f"The username '{username}' is already taken.")
        updates["username"] = username
    if full_name is not None:
        updates["full_name"] = full_name.strip()
    if is_admin is not None:
        if not is_admin and user["is_admin"] and _count_active_admins(conn, exclude_user_id=user_id) == 0:
            raise ValueError("Can't remove admin from the last remaining admin account.")
        updates["is_admin"] = 1 if is_admin else 0
    if active is not None:
        if not active and user["is_admin"] and _count_active_admins(conn, exclude_user_id=user_id) == 0:
            raise ValueError("Can't deactivate the last remaining active admin account.")
        updates["active"] = 1 if active else 0
    if not updates:
        return get_user(conn, user_id)
    updates["updated_at"] = now_iso()
    set_clause = ", ".join(f"{k}=?" for k in updates)
    conn.execute(f"UPDATE users SET {set_clause} WHERE id=?", (*updates.values(), user_id))
    conn.commit()
    return get_user(conn, user_id)


def set_user_password(conn, user_id, new_password):
    """Sets a real password directly -- used both by an admin who wants to
    pick a password for someone themselves, and by CreatePasswordDialog
    when an account holder chooses their own at first login. Always
    clears must_set_password (Batch 114), whichever of those two set it:
    a real password now exists, so the "create your password" routing in
    LoginDialog._try_login has nothing left to route to."""
    if not new_password:
        raise ValueError("Password is required.")
    salt, pw_hash = hash_password(new_password)
    conn.execute(
        "UPDATE users SET password_hash=?, password_salt=?, must_set_password=0, updated_at=? WHERE id=?",
        (pw_hash, salt, now_iso(), user_id),
    )
    conn.commit()


def reset_user_password(conn, user_id):
    """Batch 114: an admin's fix for "I forgot my password" on someone
    else's account. Deliberately takes no new password at all -- clears
    whatever's currently set (replaced with the same kind of random,
    unusable placeholder create_user() gives a brand new account) and
    flips must_set_password back on, so the very next time this account
    tries to log in it's routed straight to "create your password"
    instead of being handed whatever the admin might otherwise have
    picked for them. Raises ValueError if the user doesn't exist."""
    if not get_user(conn, user_id):
        raise ValueError("No such user.")
    salt, pw_hash = _unusable_placeholder_password_hash()
    conn.execute(
        "UPDATE users SET password_hash=?, password_salt=?, must_set_password=1, updated_at=? WHERE id=?",
        (pw_hash, salt, now_iso(), user_id),
    )
    conn.commit()


def set_user_permissions(conn, user_id, permission_keys):
    """Replaces this user's entire permission set with exactly the given
    keys (unknown keys silently ignored, so a stray typo can't grant
    something that doesn't exist). Has no effect on an admin account's
    actual access -- is_admin bypasses this table entirely, see
    user_has_permission -- but the ticked boxes are still saved, so
    they're already in place if that account is ever demoted from admin."""
    conn.execute("DELETE FROM user_permissions WHERE user_id=?", (user_id,))
    for key in permission_keys or []:
        if key in PERMISSION_KEYS:
            # Batch 157: same ON CONFLICT swap as create_user() above -- the
            # DELETE just above already clears any old rows for this user,
            # so DO NOTHING here only ever matters if the same key appears
            # twice in permission_keys itself.
            conn.execute(
                "INSERT INTO user_permissions(user_id, permission_key) VALUES (?, ?) "
                "ON CONFLICT (user_id, permission_key) DO NOTHING",
                (user_id, key),
            )
    conn.commit()


def delete_user(conn, user_id):
    user = conn.execute("SELECT * FROM users WHERE id=?", (user_id,)).fetchone()
    if not user:
        return
    if user["is_admin"] and _count_active_admins(conn, exclude_user_id=user_id) == 0:
        raise ValueError("Can't delete the last remaining admin account.")
    conn.execute("DELETE FROM users WHERE id=?", (user_id,))
    conn.commit()


def verify_login(conn, username, password):
    """Returns the full user dict (permissions included) on a correct,
    active login; None on a wrong password, an unknown username, or an
    account that's been deactivated -- deliberately the same None either
    way, so a login screen can't be used to fish for which usernames
    exist."""
    row = conn.execute("SELECT * FROM users WHERE username=? COLLATE NOCASE", (username,)).fetchone()
    if not row or not row["active"]:
        return None
    if not verify_password(password, row["password_salt"], row["password_hash"]):
        return None
    return _row_to_user(row, conn)


def user_has_permission(user, key):
    """The one function every permission check in the UI goes through, so
    the "admin bypasses everything" rule lives in exactly one place.
    `user` is a dict as returned by verify_login/get_user (needs a
    "permissions" set and an "is_admin" flag) -- or None, meaning "nobody's
    logged in," which never has permission for anything."""
    if not user:
        return False
    if user.get("is_admin"):
        return True
    return key in (user.get("permissions") or set())


def _migrate_po_ref_prefix(conn):
    """One-time, safe-to-re-run migration: strips the literal "PO:" that
    used to be baked into every PO reference this app generated itself
    (see make_po_ref, which no longer adds it) out of the actual stored
    data -- purchase_orders.po_ref (the primary reference) and the
    denormalized copy of it on savings.po_ref (used to link a saving back
    to the order it came from). References imported from Zoho never had
    this prefix in the first place (Zoho's own Reference# field is used
    as-is -- see _normalize_ref), so this only ever touches rows this app
    itself created the old way; it's never invented data.

    Naturally idempotent even without the settings flag below -- the WHERE
    po_ref LIKE 'PO:%' clause matches nothing once every row's already
    been fixed, so running this again (or even removing the flag) can
    never re-damage already-correct data. The flag just avoids re-scanning
    both tables on every single launch once there's genuinely nothing left
    to do. A backup is taken first, but only if there's actually something
    to migrate, so a fresh install or one that's already been migrated
    never gets an extra backup it doesn't need."""
    if get_setting(conn, "po_ref_prefix_migration_done", "0") == "1":
        return
    to_fix = conn.execute("SELECT COUNT(*) AS c FROM purchase_orders WHERE po_ref LIKE 'PO:%'").fetchone()["c"]
    to_fix += conn.execute("SELECT COUNT(*) AS c FROM savings WHERE po_ref LIKE 'PO:%'").fetchone()["c"]
    if to_fix:
        backup_now(reason="pre-po-ref-prefix-migration")
        conn.execute("UPDATE purchase_orders SET po_ref = substr(po_ref, 4) WHERE po_ref LIKE 'PO:%'")
        conn.execute("UPDATE savings SET po_ref = substr(po_ref, 4) WHERE po_ref LIKE 'PO:%'")
        conn.commit()
        log.info("Migrated %s off the old stored 'PO:' prefix", pluralize(to_fix, "PO reference"))
    set_settings(conn, {"po_ref_prefix_migration_done": "1"})


def _migrate_zoho_vendor_map_seq(conn):
    """One-time backfill for zoho_vendor_map.seq (Batch 157). Every row that
    predates the seq column still shows the ensure_column default of 0 --
    genuinely new rows never land on 0 (set_zoho_vendor_map() always writes
    at least 1, one more than the current highest seq), so "seq=0" reliably
    means "never backfilled yet" and this needs no separate settings flag
    the way _migrate_po_ref_prefix() above does.

    Ranks the still-zero rows by SQLite's own rowid, the exact thing seq is
    replacing -- this is the one and only place seq is ever allowed to read
    rowid, since this is a one-time snapshot of history that already
    happened under SQLite; every future write goes through the portable
    MAX(seq)+1 subquery in set_zoho_vendor_map() instead, which needs no
    rowid at all and runs identically on Postgres."""
    conn.execute(
        "UPDATE zoho_vendor_map SET seq = "
        "(SELECT COUNT(*) FROM zoho_vendor_map z2 WHERE z2.rowid <= zoho_vendor_map.rowid) "
        "WHERE seq = 0 OR seq IS NULL"
    )
    conn.commit()


def _migrate_report_frequency_to_multi(conn):
    """One-time migration from the old single-frequency periodic PO report
    schedule (exactly one of weekly/monthly/quarterly/yearly ever "the"
    active schedule, chosen via po_report_frequency) to the current
    multi-frequency one, where any combination can be enabled at once.

    Prompted by real use, in Yitzi's own words: "we are now in sept when i
    opend the program it reminded me to send this weeks report and stock
    recap but did not remind me to send august report it needs to remind
    me for the monthly quater and yearly aswell." The report had been set
    to weekly, which faithfully reminded him every week -- but monthly/
    quarterly/yearly never fired even once, because only one frequency was
    ever live at a time, no matter how overdue the others were.

    If the periodic report was already enabled under the old system, all
    four frequencies come out enabled here too -- "as well as" weekly, not
    "instead of" -- since that's the plain reading of the request and
    matches what he already has ON right now. The previously-active
    frequency keeps its exact day and already-sent state (so nothing that
    was already correctly handled suddenly re-fires), while the three that
    were never active before start with no history at all, so each simply
    becomes due on its own next scheduled occurrence -- the same as if it
    had just been freshly ticked on in Settings. If the report was off
    entirely, nothing gets auto-enabled here -- this only ever expands an
    already-active schedule, never switches one on behind the user's
    back."""
    if get_setting(conn, "po_report_freq_migration_done", "0") == "1":
        return
    old_enabled = get_setting(conn, "po_report_enabled", "0") == "1"
    old_freq = normalize_report_frequency(get_setting(conn, "po_report_frequency", "weekly"))
    old_day = get_setting(conn, "po_report_day", "")
    old_last_sent = get_setting(conn, "po_report_last_sent_at", "")

    updates = {f"po_report_{freq}_enabled": ("1" if old_enabled else "0") for freq in REPORT_FREQUENCIES}
    if old_day:
        updates[f"po_report_{old_freq}_day"] = old_day
    if old_last_sent:
        updates[f"po_report_last_sent_{old_freq}"] = old_last_sent
    updates["po_report_freq_migration_done"] = "1"
    set_settings(conn, updates)
    if old_enabled:
        log.info(
            "Migrated the periodic PO report from single-frequency (%s) to "
            "multi-frequency -- weekly/monthly/quarterly/yearly all enabled now",
            old_freq,
        )


def _migrate_seed_supplier_name_acronyms(conn):
    """One-time seed for the supplier_name_acronyms setting used by the
    Suppliers page's "Clean up naming" tool (see normalize_supplier_name /
    suggest_supplier_name_cleanup below). Prompted by Yitzi's own feedback
    on supplier data quality: "'rvt' should be 'RVT'" was given as the
    example. Title-casing alone can't turn a lowercase "rvt" into "RVT" --
    there's no way to tell an ordinary word from an acronym just by
    looking at it -- so this seeds the acronym list once from whatever's
    already written in all caps across the existing supplier list (so an
    already-correct name like "EGE" is never treated as something to fix)
    plus "RVT" itself, since that's the literal case flagged. Editable
    from the cleanup tool afterwards; safe to re-run (a no-op once done)."""
    if get_setting(conn, "supplier_name_acronyms_seeded", "0") == "1":
        return
    found = set()
    for s in list_suppliers(conn):
        for word in (s["company_name"] or "").split(" "):
            letters = "".join(ch for ch in word if ch.isalpha())
            if letters.isupper() and len(letters) >= 2:
                found.add(letters)
    found.add("RVT")
    existing = [a.strip() for a in get_setting(conn, "supplier_name_acronyms", "").split(",") if a.strip()]
    merged = sorted(set(existing) | found)
    set_settings(conn, {
        "supplier_name_acronyms": ", ".join(merged),
        "supplier_name_acronyms_seeded": "1",
    })


DEFAULT_SAVINGS_CATEGORIES = [
    "Negotiated savings", "Alternative supplier", "Improved pricing", "Avoided cost",
    "Bulk discount", "Avoided price increase",
]

DEFAULT_SUPPLIER_ISSUE_CATEGORIES = [
    "Late delivery", "Quality issue", "Pricing dispute", "Communication", "Other",
]


def _seed_defaults(conn):
    for key, value in DEFAULT_SETTINGS.items():
        row = conn.execute("SELECT 1 FROM settings WHERE key=?", (key,)).fetchone()
        if not row:
            conn.execute("INSERT INTO settings(key, value) VALUES (?, ?)", (key, value))
    if not conn.execute("SELECT 1 FROM addresses LIMIT 1").fetchone():
        conn.execute(
            "INSERT INTO addresses(label, company, address, sort_order) VALUES (?, ?, ?, ?)",
            ("N16", "ROSE COMMUNICATIONS GROUP LTD", "92-94 STAMFORD HILL\nLONDON\nN16 6XS", 0),
        )
    if not conn.execute("SELECT 1 FROM savings_categories LIMIT 1").fetchone():
        for i, name in enumerate(DEFAULT_SAVINGS_CATEGORIES):
            conn.execute(
                "INSERT INTO savings_categories(name, sort_order, active) VALUES (?, ?, 1)", (name, i)
            )
    if not conn.execute("SELECT 1 FROM supplier_issue_categories LIMIT 1").fetchone():
        for i, name in enumerate(DEFAULT_SUPPLIER_ISSUE_CATEGORIES):
            conn.execute(
                "INSERT INTO supplier_issue_categories(name, sort_order, active) VALUES (?, ?, 1)", (name, i)
            )
    conn.commit()


def _seed_own_purchase_owner_prefix(conn):
    """Seed one purchase_owner_prefixes row for this install's own
    prefix/name, if both are already set (nothing to guess if they aren't)
    -- everything else (other staff members' prefixes) has to be added in
    Settings, since there's no way to know their names automatically.

    Deliberately called AFTER _migrate_from_legacy_json, not from inside
    _seed_defaults -- on a brand new database, _seed_defaults runs before
    the legacy JSON settings (po_prefix/your_name) have been imported, so
    checking here (once those are in place) is what actually lets this
    seed correctly on first launch instead of silently seeding nothing."""
    if conn.execute("SELECT 1 FROM purchase_owner_prefixes LIMIT 1").fetchone():
        return
    own_prefix = get_setting(conn, "po_prefix", "") or DEFAULT_SETTINGS.get("po_prefix", "")
    own_name = get_setting(conn, "your_name", "")
    if own_prefix.strip() and own_name.strip():
        conn.execute(
            "INSERT INTO purchase_owner_prefixes(prefix, name, sort_order) VALUES (?, ?, 0)",
            (own_prefix.strip().upper(), own_name.strip()),
        )
        conn.commit()


# ---- settings ----

def get_setting(conn, key, default=""):
    row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    return row["value"] if row else default


def get_all_settings(conn):
    out = dict(DEFAULT_SETTINGS)
    for row in conn.execute("SELECT key, value FROM settings"):
        out[row["key"]] = row["value"]
    return out


def set_settings(conn, mapping):
    for key, value in mapping.items():
        conn.execute(
            "INSERT INTO settings(key, value) VALUES (?, ?) "
            "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (key, str(value)),
        )
    conn.commit()


# ---- addresses ----

def list_addresses(conn):
    return [dict(r) for r in conn.execute("SELECT * FROM addresses ORDER BY sort_order, id")]


def replace_addresses(conn, addresses):
    """Addresses are edited as a whole list in Backend Settings, like v1."""
    conn.execute("DELETE FROM addresses")
    for i, a in enumerate(addresses):
        conn.execute(
            "INSERT INTO addresses(label, company, address, sort_order) VALUES (?, ?, ?, ?)",
            (a.get("label", f"Option {i+1}"), a.get("company", ""), a.get("address", ""), i),
        )
    conn.commit()


# ---- suppliers ----

def list_suppliers(conn, active_only=False):
    q = "SELECT * FROM suppliers"
    if active_only:
        q += " WHERE active=1"
    q += " ORDER BY company_name COLLATE NOCASE"
    return [dict(r) for r in conn.execute(q)]


def get_supplier_by_name(conn, company_name):
    row = conn.execute(
        "SELECT * FROM suppliers WHERE company_name=?", (company_name,)
    ).fetchone()
    return dict(row) if row else None


def get_supplier_by_id(conn, supplier_id):
    """Batch 147: NewPOPage's own supplier_combo is keyed by company_name,
    not id (see get_supplier_by_name above, already used for that) -- this
    is the id-first lookup a caller starting from a stored supplier_id
    (e.g. a product_requests row's effective_supplier_id) needs before it
    can select that same combo entry."""
    if not supplier_id:
        return None
    row = conn.execute("SELECT * FROM suppliers WHERE id=?", (supplier_id,)).fetchone()
    return dict(row) if row else None


def _cascade_supplier_name_change(conn, supplier_id, old_name, new_name):
    """Relabels every place a supplier's name is denormalized as plain text
    once the supplier itself has been renamed (whether that's a genuine
    merge, or just a naming/casing fix like "rvt" -> "RVT") -- past
    purchase orders, supplier issues, savings, and price request replies --
    so the change doesn't leave a supplier's own history split across two
    spellings. Matches by supplier_id where that link exists (purchase_
    orders, price request replies) as well as by the old name text, since a
    handful of older rows may only ever have carried the name and never a
    real link. Runs the id-based reassignment even when old_name and
    new_name happen to be the same string -- a merge needs that regardless
    of naming (eg. merging into a supplier that already has the name being
    kept), it isn't only about relabelling text."""
    if not new_name:
        return
    conn.execute(
        "UPDATE purchase_orders SET supplier_id=?, supplier_company_name=? "
        "WHERE supplier_id=? OR supplier_company_name=?",
        (supplier_id, new_name, supplier_id, old_name),
    )
    conn.execute(
        "UPDATE supplier_issues SET supplier_company_name=? WHERE supplier_company_name=?",
        (new_name, old_name),
    )
    conn.execute(
        "UPDATE savings SET supplier_company_name=? WHERE supplier_company_name=?",
        (new_name, old_name),
    )
    conn.execute(
        "UPDATE price_request_replies SET supplier_id=?, supplier_company_name=? "
        "WHERE supplier_id=? OR supplier_company_name=?",
        (supplier_id, new_name, supplier_id, old_name),
    )


def replace_suppliers(conn, suppliers):
    """Suppliers are edited as a whole list in Backend Settings, like v1.
    Matches primarily by id (when the caller has one -- editing an existing
    supplier, see SuppliersPage._edit) so a rename is a genuine UPDATE of
    that same row rather than being treated as delete-the-old-row/insert-a-
    new-one under a fresh id, which used to silently orphan every product-
    catalog link (products.supplier_id) to that supplier and split its own
    purchase/issue/savings history across the old and new spellings. Falls
    back to matching by company_name for entries with no id (a fresh Add
    Supplier, or a CSV import row -- both never carry one). Any name change
    also cascades to every place that name is denormalized as plain text,
    via _cascade_supplier_name_change, so a naming cleanup (or any manual
    edit that changes the name) keeps a supplier's history reading as one
    consistent name throughout."""
    existing_rows = {r["id"]: dict(r) for r in conn.execute("SELECT id, company_name FROM suppliers")}
    existing_by_name = {r["company_name"]: sid for sid, r in existing_rows.items()}
    keep_ids = set()
    for s in suppliers:
        name = (s.get("company_name") or "").strip()
        if not name:
            continue
        try:
            credit_limit = float(s.get("credit_limit") or 0)
        except (TypeError, ValueError):
            credit_limit = 0
        sid = s.get("id")
        if sid is None or sid not in existing_rows:
            sid = existing_by_name.get(name)
        if sid is not None and sid in existing_rows:
            old_name = existing_rows[sid]["company_name"]
            conn.execute(
                "UPDATE suppliers SET company_name=?, contact_name=?, email=?, cc_emails=?, phone=?, address=?, "
                "active=?, credit_limit=? WHERE id=?",
                (
                    name, s.get("contact_name", ""), s.get("email", ""), s.get("cc_emails", ""),
                    s.get("phone", ""), s.get("address", ""), 1 if s.get("active", True) else 0,
                    credit_limit, sid,
                ),
            )
            if name != old_name:
                _cascade_supplier_name_change(conn, sid, old_name, name)
            keep_ids.add(sid)
        else:
            cur = conn.execute(
                "INSERT INTO suppliers(company_name, contact_name, email, cc_emails, phone, address, active, "
                "credit_limit) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                (
                    name, s.get("contact_name", ""), s.get("email", ""), s.get("cc_emails", ""),
                    s.get("phone", ""), s.get("address", ""), 1 if s.get("active", True) else 0, credit_limit,
                ),
            )
            keep_ids.add(cur.lastrowid)
    for sid in existing_rows:
        if sid not in keep_ids:
            conn.execute("DELETE FROM suppliers WHERE id=?", (sid,))
    conn.commit()


def reset_supplier_credit_balance(conn, company_name, at=None):
    """Marks 'now' (or a given timestamp) as the point credit usage is
    tracked from for this supplier -- the running total only counts orders
    from this point on. There's no live feed telling the app when Zoho
    actually gets paid, so this is how the person using the app tells it
    "that's been paid off, start counting again" by hand."""
    conn.execute(
        "UPDATE suppliers SET credit_reset_at=? WHERE company_name=?",
        (at or now_iso(), company_name),
    )
    conn.commit()


def get_supplier_credit_status(conn, company_name):
    """Credit limit vs. a running total of what's been ordered from this
    supplier since the last reset (Sent/Confirmed/Received, non-deleted,
    summed in the home currency via base_total). This is necessarily an
    approximation, not a live balance -- Zoho's PO export carries a billing
    status (Draft/Issued/Billed/Partially Billed) but no payment status, so
    there's no way to know from the data alone when an invoice has actually
    been paid and the limit freed back up. zoho_po_status is included per
    contributing order as the closest available context, not a source of
    truth for what's been paid."""
    supplier = get_supplier_by_name(conn, company_name)
    if not supplier:
        return None
    credit_limit = float(supplier.get("credit_limit") or 0)
    reset_at = supplier.get("credit_reset_at") or ""

    sql = (
        "SELECT po_ref, status, total, currency, base_total, zoho_po_status, created_at "
        "FROM purchase_orders WHERE deleted=0 AND supplier_company_name=? "
        "AND status IN ('Sent', 'Confirmed', 'Received')"
    )
    params = [company_name]
    if reset_at:
        sql += " AND created_at >= ?"
        params.append(reset_at)
    sql += " ORDER BY created_at DESC"
    contributing = [dict(r) for r in conn.execute(sql, params)]
    used = sum(float(r["base_total"] or 0) for r in contributing)
    remaining = (credit_limit - used) if credit_limit else None
    pct_used = round(used / credit_limit * 100, 1) if credit_limit else None
    return {
        "company_name": company_name,
        "credit_limit": credit_limit,
        "used": used,
        "remaining": remaining,
        "pct_used": pct_used,
        "reset_at": reset_at,
        "contributing_pos": contributing,
    }


# ---- product catalog ----

def record_product_price(conn, supplier_id, code, name, price, at=None):
    """Called every time a PO line item is saved, to keep the catalog's
    cached last-price/times-ordered up to date for that supplier+product.
    Matches the existing catalog row case/whitespace-insensitively (same
    _normalize_product_name comparison ensure_product_in_catalog already
    uses for price requests) before deciding whether this is a genuinely
    new product or just a differently-typed version of one already on
    file -- e.g. "iPhone 17 Pro" and "IPHONE 17  PRO" now update the SAME
    catalog row instead of silently creating a second, near-identical one.
    The first-seen spelling/casing stays canonical; a later differently-
    typed order just updates its price/times_ordered like normal."""
    name = (name or "").strip()
    if not name:
        return
    supplier_id = supplier_id or 0
    canonical_name = name
    for r in conn.execute("SELECT name FROM products WHERE supplier_id = ?", (supplier_id,)):
        if _product_names_match(r["name"], name):
            canonical_name = r["name"]
            break
    conn.execute(
        "INSERT INTO products(supplier_id, code, name, last_price, times_ordered, updated_at) "
        "VALUES (?, ?, ?, ?, 1, ?) "
        "ON CONFLICT(supplier_id, name) DO UPDATE SET "
        "code=excluded.code, last_price=excluded.last_price, "
        "times_ordered=times_ordered+1, updated_at=excluded.updated_at",
        (supplier_id, code or "", canonical_name, price, at or now_iso()),
    )


def check_price_alert(conn, supplier_id, product_name, entered_price, tolerance=0.01):
    """Returns the last recorded price for this product/supplier if it
    differs meaningfully from entered_price -- for warning the person when
    they're about to add a line item at a different price than last time --
    or None if there's nothing to flag (alerts turned off, no prior price on
    record, or the price hasn't actually changed)."""
    if get_setting(conn, "price_alert_enabled", "1") != "1":
        return None
    name = (product_name or "").strip()
    if not name:
        return None
    row = conn.execute(
        "SELECT last_price FROM products WHERE supplier_id=? AND name=?",
        (supplier_id or 0, name),
    ).fetchone()
    if not row:
        return None
    last_price = float(row["last_price"] or 0)
    if last_price <= 0:
        return None
    if abs(float(entered_price) - last_price) <= tolerance:
        return None
    return last_price


def get_product_catalog_live_info(conn):
    """For every product NAME that's ever actually been ordered, computed
    straight from po_items/purchase_orders (the same source of truth
    get_product_purchase_history uses) rather than from the products
    table's own cached supplier_id/last_price columns.

    Those cached columns are keyed by (supplier_id, name), so once a
    product's been bought from more than one supplier it doesn't have one
    row to summarize -- it has several. Merging two of those rows into one
    (a normal "these are the same product" cleanup) keeps only whichever
    single supplier_id/last_price the merge happened to keep, silently
    losing the fact it was also bought elsewhere, and sometimes leaving a
    stale/blank price even though a plainly-visible PO has one (this is
    the "Yealink UVC34 shows blank here but MR190126-OFFICEN16 clearly has
    a price" bug). Recomputing live sidesteps all of that -- it can never
    drift from what the PO history actually says.

    Returns {product_name: {"suppliers": [names, most-recently-bought
    first], "last_price": float or None, "last_ordered_at": iso or None}}.
    """
    rows = conn.execute(
        "SELECT pi.product AS product, po.supplier_company_name AS supplier_company_name, "
        "       pi.price AS price, po.created_at AS created_at "
        "FROM po_items pi JOIN purchase_orders po ON po.id = pi.po_id "
        "WHERE po.deleted = 0 AND pi.product != '' "
        # po.id DESC as a tiebreaker matters in practice -- Zoho PO dates
        # are often day-only, so a batch import can easily leave several
        # POs sharing the exact same created_at, and without a
        # deterministic tiebreaker "most recent price" could pick
        # whichever one SQLite happened to return first.
        "ORDER BY po.created_at DESC, po.id DESC"
    ).fetchall()
    info = {}
    for r in rows:
        name = r["product"]
        entry = info.setdefault(name, {"suppliers": [], "last_price": None, "last_ordered_at": None})
        if entry["last_price"] is None:
            entry["last_price"] = float(r["price"] or 0)
            entry["last_ordered_at"] = r["created_at"]
        supplier = r["supplier_company_name"] or ""
        if supplier and supplier not in entry["suppliers"]:
            entry["suppliers"].append(supplier)
    return info


def list_products(conn, supplier_id=None, query=None, limit=500, exact_supplier_id=None):
    """supplier_id (lenient) also includes supplier_id=0 "(any supplier)"
    catalog rows alongside the given supplier -- right for autocomplete
    (New PO's product-name suggestions should still find an unassigned
    product), but wrong for an explicit "filter to just this supplier"
    control, which is what exact_supplier_id is for (eg. the Product
    Catalog's supplier filter -- passing 0 there means "(no supplier)"
    specifically, not "any supplier"). Only one of the two is applied if
    both happen to be given -- exact_supplier_id wins."""
    sql = "SELECT * FROM products WHERE 1=1"
    params = []
    if exact_supplier_id is not None:
        sql += " AND supplier_id = ?"
        params.append(exact_supplier_id)
    elif supplier_id is not None:
        sql += " AND supplier_id IN (?, 0)"
        params.append(supplier_id)
    if query:
        sql += " AND (name LIKE ? OR code LIKE ?)"
        # Split on whitespace and join with a wildcard rather than matching
        # the query text verbatim -- so extra (or missing) spaces in either
        # the search box or the stored name don't stop an otherwise-matching
        # product from being found (e.g. "Apple  iPhone" still finds
        # "Apple iPhone").
        words = query.split()
        like = "%" + "%".join(words) + "%" if words else f"%{query}%"
        params += [like, like]
    sql += " ORDER BY times_ordered DESC, updated_at DESC LIMIT ?"
    params.append(limit)
    return [dict(r) for r in conn.execute(sql, params)]


def delete_product(conn, product_id):
    conn.execute("DELETE FROM products WHERE id=?", (product_id,))
    conn.commit()


def upsert_product_manual(conn, product_id, supplier_id, code, name, price, commit=True, rewrite_history=False):
    """rewrite_history=True (used by the Products page's "Edit selected")
    also rewrites any past PO line items that used the product's old exact
    name/code to the new one. Purchase history and the catalog's live
    Supplier/Last price columns are matched against po_items by exact name
    text (see get_product_purchase_history / get_product_catalog_live_info),
    not by this row's id -- so a plain rename here used to silently orphan
    a product's own history (it would just stop matching anything), the
    same class of bug merge_products already guards against with its own
    rewrite_history. Off by default so callers that aren't renaming an
    existing catalog entry by hand (creating a brand new product, or the
    Zoho import path, which only ever writes back the exact name it just
    matched on) are unaffected."""
    supplier_id = supplier_id or 0
    if product_id:
        existing = conn.execute(
            "SELECT supplier_id, code, name FROM products WHERE id=?", (product_id,)
        ).fetchone()
        if existing and (existing["supplier_id"], existing["name"]) != (supplier_id, name):
            clash = conn.execute(
                "SELECT id FROM products WHERE supplier_id=? AND name=? AND id != ?",
                (supplier_id, name, product_id),
            ).fetchone()
            if clash:
                raise ValueError(
                    f"Can't rename: another product already named '{name}' exists for this supplier."
                )
        if rewrite_history and existing and (existing["name"] != name or existing["code"] != code):
            old_name, old_code, old_supplier_id = existing["name"], existing["code"], existing["supplier_id"]
            # A po_items row is rewritten when its product text matches AND
            # either its own code matches the product's old code, the
            # product's old code was blank, or THAT ROW'S OWN code is blank
            # -- this last clause was missing until Batch 62 and is the fix
            # for a real orphaning bug Yitzi hit: a product ordered once
            # with no code entered and once with a code filled in ends up
            # with that code cached on its catalog row (record_product_price
            # always keeps the most recent code seen), so a later rename
            # only matched the coded po_items row and silently left the
            # blank-code one behind under the pre-rename text -- it kept
            # counting toward times_ordered (untouched by a rename) but
            # stopped showing up in Purchase History at all, since that
            # looks the product up by its current, now-mismatched name.
            code_guard = "(code=? OR ?='' OR code='')"
            if old_supplier_id == 0 or supplier_id == 0 or old_supplier_id != supplier_id:
                # Either side is the "any supplier" catalog, or the supplier
                # itself changed too -- old_name/old_code items could be
                # under any supplier's PO, so rewrite matching items
                # everywhere rather than restricting to one supplier's POs.
                conn.execute(
                    f"UPDATE po_items SET product=?, code=? WHERE product=? AND {code_guard}",
                    (name, code, old_name, old_code, old_code),
                )
            else:
                conn.execute(
                    f"UPDATE po_items SET product=?, code=? WHERE product=? AND {code_guard} "
                    "AND po_id IN (SELECT id FROM purchase_orders WHERE supplier_id=?)",
                    (name, code, old_name, old_code, old_code, old_supplier_id),
                )
        conn.execute(
            "UPDATE products SET supplier_id=?, code=?, name=?, last_price=?, updated_at=? WHERE id=?",
            (supplier_id, code, name, price, now_iso(), product_id),
        )
    else:
        # Creating a brand-new catalog entry (e.g. Products page "Add
        # product") -- match case/whitespace-insensitively (and, since the
        # tidy-up pass after Batch 151, variant-marker-aware -- see
        # _product_names_match's own docstring for why "+"/"Pro"/"Max"/
        # storage-size differences must never be treated as the same
        # product) against what's already on file for this supplier first,
        # so typing the same product with different capitalization or
        # spacing than an existing row updates that row instead of quietly
        # creating a near-duplicate one.
        existing = None
        for r in conn.execute("SELECT id, name FROM products WHERE supplier_id = ?", (supplier_id,)):
            if _product_names_match(r["name"], name):
                existing = r
                break
        if existing:
            conn.execute(
                "UPDATE products SET code=?, last_price=?, updated_at=? WHERE id=?",
                (code, price, now_iso(), existing["id"]),
            )
        else:
            conn.execute(
                "INSERT INTO products(supplier_id, code, name, last_price, times_ordered, updated_at) "
                "VALUES (?, ?, ?, ?, 0, ?)",
                (supplier_id, code, name, price, now_iso()),
            )
    if commit:
        conn.commit()


def get_product_purchase_history(conn, product_name, include_deleted=False):
    """Every PO line item that has ever ordered this exact product name,
    across every supplier -- not just whichever supplier the catalog entry
    itself is filed under. Used by the Products page's purchase-history
    view (double-click a product -> see every time it's been ordered, from
    who, on which PO, and jump straight to that PO).
    """
    sql = (
        "SELECT po.po_ref AS po_ref, po.supplier_company_name AS supplier_company_name, "
        "       po.status AS status, po.created_at AS created_at, po.deleted AS deleted, "
        "       po.purchase_owner AS purchase_owner, "
        "       pi.qty AS qty, pi.price AS price, pi.code AS code "
        "FROM po_items pi JOIN purchase_orders po ON po.id = pi.po_id "
        "WHERE pi.product = ?"
    )
    params = [product_name]
    if not include_deleted:
        sql += " AND po.deleted = 0"
    sql += " ORDER BY po.created_at DESC"
    return [dict(r) for r in conn.execute(sql, params)]


# ---- purchase orders ----

PO_FIELDS = [
    "po_ref", "status", "supplier_id", "supplier_company_name", "supplier_contact_name",
    "supplier_email", "supplier_cc_emails", "supplier_phone", "supplier_address",
    "delivery_label", "delivery_name", "delivery_address", "business_name", "company_number",
    "currency", "invoice_name", "invoice_address", "your_name", "bcc_emails", "total", "notes",
    "purchase_owner", "fx_rate", "base_total", "zoho_po_status", "created_by_user_id",
]


def _row_to_po(row):
    d = dict(row)
    return d


def get_po_full(conn, po_ref):
    row = conn.execute("SELECT * FROM purchase_orders WHERE po_ref=?", (po_ref,)).fetchone()
    if not row:
        return None
    po = _row_to_po(row)
    items = conn.execute(
        "SELECT id, qty, code, product, price, margin_vat, original_qty FROM po_items "
        "WHERE po_id=? ORDER BY position", (po["id"],)
    ).fetchall()
    po["items"] = [dict(i) for i in items]
    return po


def save_po(conn, po, items, status=None, event="saved", record_product_memory=True,
            created_at=None, updated_at=None, commit=True):
    """Insert or update a PO (matched by po_ref) plus its items, atomically.

    created_at/updated_at overrides are only used by the legacy-JSON migration,
    so imported PO history keeps its real original dates instead of being
    stamped with the migration run time.

    commit=False lets a caller batch many saves into one outer transaction
    (used by the Zoho PO import loop) instead of committing -- and fsync'ing
    -- after every single row, which is what made large imports feel like
    they'd frozen the app for a few seconds. The caller must have already
    opened a transaction (conn.execute("BEGIN")) and is responsible for the
    final conn.commit()/conn.rollback().
    """
    po_ref = po["po_ref"]
    total = sum(float(i["qty"]) * float(i["price"]) for i in items)
    existing = conn.execute("SELECT id, created_at, deleted FROM purchase_orders WHERE po_ref=?", (po_ref,)).fetchone()
    now = updated_at or now_iso()
    values = {k: po.get(k, "") for k in PO_FIELDS}
    values["total"] = total
    values["status"] = status or (po.get("status") or "Draft")

    # ---- currency: every PO gets a reliable home-currency equivalent ----
    # so cross-PO reports can always sum something consistent instead of
    # silently blending different currencies' raw numbers together. Resolved
    # centrally here (like purchase_owner above) rather than by every
    # caller. A caller in the PO's own currency (New PO's currency picker)
    # passes fx_rate and/or base_total after asking the user for one of
    # them; anything without that info (an older code path, or a Zoho
    # import -- Zoho's PO export has no exchange-rate data) falls back to
    # treating the currencies as 1:1 rather than blocking the save.
    home_currency = get_setting(conn, "currency", "GBP")
    if not str(values.get("currency") or "").strip():
        values["currency"] = home_currency
    if values["currency"] == home_currency:
        values["fx_rate"] = 1.0
        values["base_total"] = total
    else:
        try:
            fx_rate = float(po.get("fx_rate") or 0)
        except (TypeError, ValueError):
            fx_rate = 0
        try:
            base_total = float(po.get("base_total") or 0)
        except (TypeError, ValueError):
            base_total = 0
        if not fx_rate and not base_total:
            fx_rate, base_total = 1.0, total
        elif not base_total:
            base_total = round(fx_rate * total, 2)
        elif not fx_rate:
            fx_rate = (base_total / total) if total else 1.0
        values["fx_rate"] = fx_rate
        values["base_total"] = base_total
    # Batch 103: "_created_by_user" is a transient hint (never itself a
    # stored column) the UI passes through -- {"id": ..., "full_name": ...}
    # for whoever is logged in when this save happens, or nothing at all
    # for a path with no logged-in user yet (a Zoho import, a script).
    # created_by_user_id follows the exact same "only resolved while
    # blank" rule purchase_owner already used below, for the same reason:
    # a caller who already has the real value (loaded an existing PO via
    # get_po_full, which includes this column verbatim) keeps it as-is:
    # who originally raised a PO doesn't change just because someone else
    # later resaves it.
    creator = po.get("_created_by_user") or {}
    if not values.get("created_by_user_id"):
        values["created_by_user_id"] = creator.get("id") or None

    # Purchase owner is resolved here, centrally, rather than by every
    # caller of save_po -- so every path that ever creates/re-saves a PO
    # (New PO, duplicate, resend, a price-request turned into a PO, and the
    # Zoho import below) gets it automatically and consistently, with no
    # risk of one call site forgetting to set it. A caller can still pass
    # an explicit "purchase_owner" to keep it (e.g. someone corrected it by
    # hand). Priority: an explicit override wins outright (checked above);
    # otherwise the logged-in creator's own name (Yitzi logged in -> "Yitzi
    # Cukierman" on the PO, Maxy logged in -> his name, Batch 103 -- "if
    # it's Yitzy being logged in, the POs are on his name"); otherwise the
    # old prefix-mapping fallback, still needed for Zoho imports and any
    # path with no logged-in user at all. "_zoho_purchase_owner" is Zoho's
    # own per-PO field, passed through by import_zoho_po when present.
    if not str(values.get("purchase_owner") or "").strip():
        values["purchase_owner"] = (
            str(creator.get("full_name") or "").strip()
            or resolve_purchase_owner(conn, po_ref, po.get("_zoho_purchase_owner"))
        )

    try:
        if commit:
            # Tidy-up-pass safety net (post-Batch-151): if some earlier,
            # unrelated write on this same connection raised without its
            # own rollback, this connection could already be sitting in a
            # dangling open transaction -- see rename_savings_category's
            # own comment for the exact failure mode this guards against.
            # Harmless no-op the overwhelming majority of the time.
            if conn.in_transaction:
                conn.rollback()
            conn.execute("BEGIN")
        if existing:
            po_id = existing["id"]
            set_clause = ", ".join(f"{k}=:{k}" for k in PO_FIELDS)
            conn.execute(
                f"UPDATE purchase_orders SET {set_clause}, updated_at=:updated_at WHERE id=:id",
                {**values, "updated_at": now, "id": po_id},
            )
            conn.execute("DELETE FROM po_items WHERE po_id=?", (po_id,))
        else:
            cols = ", ".join(PO_FIELDS)
            placeholders = ", ".join(f":{k}" for k in PO_FIELDS)
            cur = conn.execute(
                f"INSERT INTO purchase_orders({cols}, created_at, updated_at, deleted) "
                f"VALUES ({placeholders}, :created_at, :updated_at, 0)",
                {**values, "created_at": created_at or now, "updated_at": now},
            )
            po_id = cur.lastrowid

        for i, item in enumerate(items):
            conn.execute(
                "INSERT INTO po_items(po_id, position, qty, code, product, price, margin_vat, original_qty) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                (po_id, i, item["qty"], item.get("code", ""), item["product"], item["price"],
                 1 if item.get("margin_vat") else 0, item.get("original_qty")),
            )
            if record_product_memory:
                record_product_price(conn, po.get("supplier_id") or 0, item.get("code", ""), item["product"], item["price"], at=now)

        conn.execute(
            "INSERT INTO po_events(po_id, event, detail, at) VALUES (?, ?, ?, ?)",
            (po_id, event, values["status"], now),
        )
        if commit:
            conn.commit()
    except Exception:
        if commit:
            conn.rollback()
        raise
    return get_po_full(conn, po_ref)


def list_pos(conn, deleted=False, supplier=None, status=None, query=None, flagged_only=False):
    sql = "SELECT * FROM purchase_orders WHERE deleted=?"
    params = [1 if deleted else 0]
    if supplier:
        sql += " AND supplier_company_name=?"
        params.append(supplier)
    if status:
        sql += " AND status=?"
        params.append(status)
    if flagged_only:
        sql += " AND flagged_for_review=1"
    if query:
        sql += " AND (po_ref LIKE ? OR supplier_company_name LIKE ?)"
        like = f"%{query}%"
        params += [like, like]
    sql += " ORDER BY updated_at DESC"
    rows = [dict(r) for r in conn.execute(sql, params)]
    if query:
        # also search inside item product names (needs a join, kept simple/explicit for clarity)
        # -- the supplier/status filters have to be re-applied here too, or a
        # product-name match on a PO that doesn't match those filters would
        # slip into the results.
        extra_ids = {
            r["po_id"] for r in conn.execute(
                "SELECT DISTINCT po_id FROM po_items WHERE product LIKE ?", (f"%{query}%",)
            )
        }
        if extra_ids:
            have = {r["id"] for r in rows}
            extra_sql = ("SELECT * FROM purchase_orders WHERE id IN (%s) AND deleted=?" %
                         ",".join("?" * len(extra_ids)))
            extra_params = [*extra_ids, 1 if deleted else 0]
            if supplier:
                extra_sql += " AND supplier_company_name=?"
                extra_params.append(supplier)
            if status:
                extra_sql += " AND status=?"
                extra_params.append(status)
            for r in conn.execute(extra_sql, extra_params):
                if r["id"] not in have:
                    rows.append(dict(r))
            rows.sort(key=lambda r: r["updated_at"], reverse=True)
    return rows


def set_po_deleted(conn, po_ref, deleted):
    conn.execute(
        "UPDATE purchase_orders SET deleted=?, updated_at=? WHERE po_ref=?",
        (1 if deleted else 0, now_iso(), po_ref),
    )
    row = conn.execute("SELECT id FROM purchase_orders WHERE po_ref=?", (po_ref,)).fetchone()
    if row:
        conn.execute(
            "INSERT INTO po_events(po_id, event, detail, at) VALUES (?, ?, ?, ?)",
            (row["id"], "deleted" if deleted else "restored", "", now_iso()),
        )
    conn.commit()


def set_po_status(conn, po_ref, status):
    conn.execute(
        "UPDATE purchase_orders SET status=?, updated_at=? WHERE po_ref=?",
        (status, now_iso(), po_ref),
    )
    row = conn.execute("SELECT id FROM purchase_orders WHERE po_ref=?", (po_ref,)).fetchone()
    if row:
        conn.execute(
            "INSERT INTO po_events(po_id, event, detail, at) VALUES (?, ?, ?, ?)",
            (row["id"], "status_changed", status, now_iso()),
        )
    conn.commit()


def set_po_flagged(conn, po_ref, flagged):
    """Marks a PO as flagged for a second opinion before it's actually sent
    (or unflags it) -- a lightweight nudge, not a hard approval gate: nothing
    stops the PO from being sent while flagged, it's just a way to mark
    "check this one before it goes out" and see those in one place."""
    conn.execute(
        "UPDATE purchase_orders SET flagged_for_review=? WHERE po_ref=?",
        (1 if flagged else 0, po_ref),
    )
    row = conn.execute("SELECT id FROM purchase_orders WHERE po_ref=?", (po_ref,)).fetchone()
    if row:
        conn.execute(
            "INSERT INTO po_events(po_id, event, detail, at) VALUES (?, ?, ?, ?)",
            (row["id"], "flagged" if flagged else "unflagged", "", now_iso()),
        )
    conn.commit()


def add_po_event(conn, po_ref, event, detail=""):
    """Logs a free-form history entry against a PO (e.g. 'resent') without
    touching its status -- see set_po_status for the status-change variant,
    which logs its own 'status_changed' event."""
    row = conn.execute("SELECT id FROM purchase_orders WHERE po_ref=?", (po_ref,)).fetchone()
    if not row:
        return
    conn.execute(
        "INSERT INTO po_events(po_id, event, detail, at) VALUES (?, ?, ?, ?)",
        (row["id"], event, detail, now_iso()),
    )
    conn.commit()


def adjust_po_item_quantity(conn, po_ref, item_id, new_qty, note=""):
    """Records a supplier coming back with less stock than ordered -- or
    none at all -- for one line on an already-placed PO. Yitzi's own
    request: "i order 100 phones but the supplier says he dose not have or
    less quantity i wnat to be able to log that so it removes from PO or
    updates quantity but still shows that it was once there."

    Reduces that single line's qty (and the PO's own total/base_total) to
    what's actually being fulfilled -- new_qty can be 0, for "they have
    none at all" -- while permanently remembering the amount actually
    ordered in original_qty. That's stamped only the first time a line is
    ever adjusted (a second, later adjustment doesn't overwrite the true
    original with an already-reduced figure), so the line stays visible on
    the PO showing both what was ordered and what's actually coming,
    rather than being silently deleted or having its history overwritten.
    Also logs a "qty_adjusted" po_event so it shows up in the PO's own
    history, same as every other status/edit event.

    Deliberately a narrow, direct UPDATE rather than routing through
    save_po -- save_po deletes and reinserts every line on the PO and
    stamps a fresh status event, which is the right shape for a full
    re-edit but far more than a one-line quantity correction needs (and
    would re-run product-memory tracking for every other untouched line
    too). Returns the updated po_items row."""
    po = conn.execute("SELECT * FROM purchase_orders WHERE po_ref=?", (po_ref,)).fetchone()
    if not po:
        raise ValueError(f"No such PO: {po_ref}")
    item = conn.execute("SELECT * FROM po_items WHERE id=? AND po_id=?", (item_id, po["id"])).fetchone()
    if not item:
        raise ValueError("That line item is no longer on this PO.")
    try:
        new_qty = float(new_qty)
    except (TypeError, ValueError):
        raise ValueError("Enter a valid quantity.")
    if new_qty < 0:
        raise ValueError("Quantity can't be negative.")
    old_qty = float(item["qty"])
    if new_qty == old_qty:
        return dict(item)  # nothing actually changed -- no event to log
    original_qty = item["original_qty"] if item["original_qty"] is not None else old_qty

    conn.execute("UPDATE po_items SET qty=?, original_qty=? WHERE id=?", (new_qty, original_qty, item["id"]))

    remaining = conn.execute("SELECT qty, price FROM po_items WHERE po_id=?", (po["id"],)).fetchall()
    total = sum(float(r["qty"]) * float(r["price"]) for r in remaining)
    home_currency = get_setting(conn, "currency", "GBP")
    if (po["currency"] or home_currency) == home_currency:
        base_total = total
    else:
        base_total = round(float(po["fx_rate"] or 1.0) * total, 2)
    conn.execute(
        "UPDATE purchase_orders SET total=?, base_total=?, updated_at=? WHERE id=?",
        (total, base_total, now_iso(), po["id"]),
    )

    detail = f"{item['product']}: qty {old_qty:g} -> {new_qty:g} (originally ordered {original_qty:g})"
    if note.strip():
        detail += f" -- {note.strip()}"
    add_po_event(conn, po_ref, "qty_adjusted", detail)
    return dict(conn.execute("SELECT * FROM po_items WHERE id=?", (item["id"],)).fetchone())


def get_po_events(conn, po_ref):
    row = conn.execute("SELECT id FROM purchase_orders WHERE po_ref=?", (po_ref,)).fetchone()
    if not row:
        return []
    return [dict(r) for r in conn.execute(
        "SELECT event, detail, at FROM po_events WHERE po_id=? ORDER BY at", (row["id"],)
    )]


_FOLLOWUP_TOUCH_EVENTS = ("resent", "followed_up", "opened in outlook", "copied for outlook", "pdf saved")


def mark_po_followed_up(conn, po_ref, note=""):
    """Resets a PO's chase-list clock -- logs a 'followed_up' event so it
    drops off report_pos_needing_followup() until the configured number of
    days has passed again without any further action on it."""
    add_po_event(conn, po_ref, "followed_up", detail=note)


def report_pos_needing_followup(conn, days=None):
    """Sent (not deleted) POs that haven't been touched -- sent, resent,
    copied for Outlook, or explicitly marked followed-up -- in at least
    `days` days. A lightweight "who do I need to chase?" nudge list, not a
    delivery/confirmation tracker (this app deliberately doesn't track
    whether goods actually arrived -- that's the stock team's job)."""
    if days is None:
        try:
            days = int(get_setting(conn, "chase_followup_days", "5") or 5)
        except ValueError:
            days = 5
    cutoff = (datetime.now() - timedelta(days=days)).isoformat(timespec="seconds")
    placeholders = ", ".join("?" * len(_FOLLOWUP_TOUCH_EVENTS))
    rows = conn.execute(
        "SELECT po.*, COALESCE(MAX(ev.at), po.created_at) AS last_touch "
        "FROM purchase_orders po "
        f"LEFT JOIN po_events ev ON ev.po_id = po.id AND ev.event IN ({placeholders}) "
        "WHERE po.status='Sent' AND po.deleted=0 "
        "GROUP BY po.id "
        "HAVING datetime(last_touch) <= datetime(?) "
        "ORDER BY last_touch ASC",
        (*_FOLLOWUP_TOUCH_EVENTS, cutoff),
    ).fetchall()
    out = []
    for r in rows:
        d = dict(r)
        try:
            d["days_since_touch"] = (datetime.now() - datetime.fromisoformat(d["last_touch"])).days
        except Exception:
            d["days_since_touch"] = None
        out.append(d)
    return out


def report_pos_needing_zoho_export(conn):
    """Sent (not deleted) POs the accounts team don't have in Zoho yet --
    however old, so nothing gets missed just because a day (or a week) was
    skipped. Oldest first, so exporting works through them in the order
    they were raised."""
    rows = conn.execute(
        "SELECT * FROM purchase_orders WHERE status='Sent' AND deleted=0 AND zoho_exported=0 "
        "ORDER BY created_at ASC"
    ).fetchall()
    return [dict(r) for r in rows]


def mark_pos_exported_to_zoho(conn, po_refs):
    """Marks the given POs as exported to Zoho -- called once the export
    email's draft has actually been opened (this app never sends email on
    its own, so "exported" means "ready and handed off", the same way a PO
    email itself is marked Sent as soon as its draft opens)."""
    now = now_iso()
    for po_ref in po_refs:
        conn.execute(
            "UPDATE purchase_orders SET zoho_exported=1, zoho_exported_at=? WHERE po_ref=?",
            (now, po_ref),
        )
        row = conn.execute("SELECT id FROM purchase_orders WHERE po_ref=?", (po_ref,)).fetchone()
        if row:
            conn.execute(
                "INSERT INTO po_events(po_id, event, detail, at) VALUES (?, ?, ?, ?)",
                (row["id"], "exported to zoho", "", now),
            )
    conn.commit()


def distinct_supplier_names_in_history(conn):
    return sorted({
        r["supplier_company_name"] for r in conn.execute(
            "SELECT DISTINCT supplier_company_name FROM purchase_orders WHERE deleted=0"
        ) if r["supplier_company_name"]
    })


# ============================================================
# 3. Migration from v1 (loose JSON files)
# ============================================================

def _archive_legacy_file(path):
    """Rename a legacy v1 file out of the way after it's been imported, without
    ever clobbering a previous archive. If the user kept using the old v1 app
    for a while after switching to v2, this file may get recreated more than
    once — each one gets its own uniquely-named archive copy."""
    dest = Path(str(path) + ".migrated")
    if dest.exists():
        dest = Path(str(path) + f".migrated-{datetime.now().strftime('%Y%m%d%H%M%S')}")
    try:
        path.rename(dest)
    except Exception:
        log.exception("Failed renaming legacy file %s", path)


def _migrate_from_legacy_json(conn):
    """Imports the old v1 JSON files into the database, then archives them.

    This intentionally runs on EVERY launch, not just once. Early versions of
    this function only ever ran a single time (gated by a 'migration_done'
    settings flag), which meant that if the old v1 app kept being used for a
    while after switching to v2 — even just to raise one more PO — that data
    would silently never make it into the new database. save_po() upserts by
    po_ref, so re-importing POs that are already in the database is a safe
    no-op; only genuinely new legacy activity actually changes anything.
    """
    imported = False
    try:
        if LEGACY_SETTINGS_FILE.exists():
            with open(LEGACY_SETTINGS_FILE, "r", encoding="utf-8") as f:
                old = json.load(f)
            mapping = {}
            for key in ("business_name", "company_number", "currency", "po_prefix",
                        "invoice_name", "invoice_address", "bcc_emails", "your_name"):
                if old.get(key):
                    mapping[key] = old[key]
            if mapping:
                set_settings(conn, mapping)
            addresses = old.get("addresses") or []
            if addresses:
                replace_addresses(conn, addresses)
            suppliers = old.get("suppliers") or []
            if suppliers:
                replace_suppliers(conn, suppliers)
            imported = True
            log.info("Migrated legacy settings from %s", LEGACY_SETTINGS_FILE)
            _archive_legacy_file(LEGACY_SETTINGS_FILE)
    except Exception:
        log.exception("Failed migrating legacy settings file")

    try:
        if LEGACY_HISTORY_FILE.exists():
            with open(LEGACY_HISTORY_FILE, "r", encoding="utf-8") as f:
                old_history = json.load(f)
            supplier_ids_by_name = {
                r["company_name"]: r["id"] for r in conn.execute("SELECT id, company_name FROM suppliers")
            }
            migrated_count = 0
            if isinstance(old_history, list):
                # Process oldest-first so that, once product price memory is built from
                # this history, "last price" really is the most recent price paid.
                old_history = sorted(
                    old_history, key=lambda r: r.get("updated_at") or r.get("created_at") or ""
                )
                for record in old_history:
                    po_ref = record.get("po_ref")
                    if not po_ref:
                        continue
                    items = record.get("items", [])
                    po = {k: record.get(k, "") for k in PO_FIELDS}
                    po["po_ref"] = po_ref
                    supplier_name = record.get("supplier_company_name", "")
                    po["supplier_id"] = supplier_ids_by_name.get(supplier_name)
                    try:
                        save_po(conn, po, items, status=record.get("status", "Draft"),
                                 event="migrated from v1", record_product_memory=True,
                                 created_at=record.get("created_at"), updated_at=record.get("updated_at"))
                        if record.get("deleted"):
                            set_po_deleted(conn, po_ref, True)
                        migrated_count += 1
                    except Exception:
                        log.exception("Failed migrating PO %s", po_ref)
                imported = True
                log.info("Migrated %d/%d legacy PO history rows", migrated_count, len(old_history))
            _archive_legacy_file(LEGACY_HISTORY_FILE)
    except Exception:
        log.exception("Failed migrating legacy history file")

    if imported:
        set_settings(conn, {"migration_done": "1", "last_migration_at": now_iso()})


# ============================================================
# 4. Business logic — formatting, parsing, validation, email builders
# ============================================================

def make_po_ref(prefix, conn=None, now=None):
    """Builds a new PO reference. The DDMMYY date portion is always fixed
    (parse_date_from_po_ref and everything downstream of it depends on
    recognizing that format), but what follows the dash is controlled by
    the "po_ref_suffix_style" setting: "time" (HHMM, the default -- what
    time of day the PO was created) or "sequence" (001, 002, ... -- the
    Nth PO created today, so two POs started in the same minute never end
    up with an identical reference)."""
    now = now or datetime.now()
    date_part = now.strftime("%d%m%y")
    style = get_setting(conn, "po_ref_suffix_style", "time") if conn is not None else "time"
    if style == "sequence" and conn is not None:
        today_str = now.strftime("%Y-%m-%d")
        row = conn.execute(
            "SELECT COUNT(*) AS c FROM purchase_orders WHERE date(created_at) = ?",
            (today_str,),
        ).fetchone()
        seq = (row["c"] if row else 0) + 1
        suffix = f"{seq:03d}"
    else:
        suffix = now.strftime("%H%M")
    return f"{prefix}{date_part}-{suffix}"


def display_po_ref(po_ref):
    """Display-only helper: adds the human-readable "PO: " label in front
    of a bare stored reference (e.g. "YJ140826-0933" -> "PO: YJ140826-0933"),
    for UI contexts that want that visual cue -- currently just the New PO
    screen's on-screen number. Never store the result of this anywhere --
    make_po_ref() above deliberately no longer bakes "PO:" into the
    reference itself, and every stored po_ref should stay bare."""
    po_ref = (po_ref or "").strip()
    return f"PO: {po_ref}" if po_ref else ""


def filename_po_ref(po_ref):
    return po_ref.replace(":", "")


def sanitize_filename(name):
    """Strips characters that are illegal in Windows filenames (\\ / : * ?
    " < > |), collapses runs of whitespace, and trims the result -- so any
    display string (a PO ref, a product name, a date range) can safely
    become part of a real filename on disk. Falls back to "file" if
    stripping leaves nothing behind."""
    cleaned = re.sub(r'[\\/:*?"<>|]', "", str(name or ""))
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned or "file"


def _named_pdf_temp_path(basename):
    """Creates a fresh, dedicated temp directory and returns a path inside
    it named after `basename` (sanitized) + ".pdf" -- unlike
    tempfile.NamedTemporaryFile, this lets a caller control the exact
    filename Outlook will display for the attachment, since Outlook's
    Attachments.Add derives the shown name from the file's actual basename
    on disk rather than anything set programmatically. A fresh directory
    per call avoids any risk of colliding with a previous attachment of the
    same name still sitting around from an earlier draft."""
    import tempfile
    d = tempfile.mkdtemp(prefix="yjpo_")
    return os.path.join(d, sanitize_filename(basename) + ".pdf")


_PO_REF_DATE_RE_8 = re.compile(r'^(?:PO:)?[A-Za-z]+[:\-]?(\d{2})(\d{2})(\d{4})-.+$')
_PO_REF_DATE_RE_6 = re.compile(r'^(?:PO:)?[A-Za-z]+[:\-]?(\d{2})(\d{2})(\d{2})-.+$')


def parse_date_from_po_ref(po_ref):
    """Purchase order references encode an order date right after the
    prefix letters, as either DDMMYY or DDMMYYYY -- e.g. YJ140826-0933,
    YJ:14082026-1241, YJ-14082026-1241, or MR140826-9999 all encode 14 Aug.
    The separator between the prefix and the date (none, ":", or "-") and
    the trailing part after the final dash (a real HHMM time for this app's
    own YJ... refs, just a filler/sequence number for older MR... refs and
    others carried over from Zoho) don't matter here -- only the date is
    being pulled out. Returns a 'YYYY-MM-DD' string, or None if the
    reference doesn't have this shape at all (e.g. a manually typed one-off
    ref like "ZERG0017")."""
    if not po_ref:
        return None
    ref = po_ref.strip()
    m = _PO_REF_DATE_RE_8.match(ref)
    if m:
        dd, mm, yyyy = (int(g) for g in m.groups())
        try:
            return datetime(yyyy, mm, dd).strftime("%Y-%m-%d")
        except ValueError:
            return None
    m = _PO_REF_DATE_RE_6.match(ref)
    if not m:
        return None
    dd, mm, yy = (int(g) for g in m.groups())
    try:
        return datetime(2000 + yy, mm, dd).strftime("%Y-%m-%d")
    except ValueError:
        return None


def html_escape(text):
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def pluralize(count, noun, plural=None):
    """Returns "{count} {noun}" with the noun correctly singular or plural --
    "1 order" / "8 orders" -- instead of the lazy "order(s)" this app used
    to write everywhere a count was shown. plural defaults to noun + "s",
    which covers every real word this app pluralizes this way (order, item,
    supplier, product, vendor, row, draft, time, month, reference...); pass
    plural explicitly for the rare irregular case (eg. "company"/"companies")
    if one ever comes up. count can be a float (eg. a quantity) -- only the
    exact value 1 (or 1.0) is treated as singular."""
    word = noun if count == 1 else (plural or f"{noun}s")
    return f"{count} {word}"


def plural_word(count, noun, plural=None):
    """Just the correctly singular/plural noun on its own, no count prefixed
    -- for the layouts (stat tiles, KPI cards) that already show the number
    separately, in its own bold/larger styling, with the word underneath or
    beside it rather than inline with the digits."""
    return noun if count == 1 else (plural or f"{noun}s")


def diff_highlight_html(a, b):
    """Compares two strings character-by-character and returns (html_a,
    html_b) with whatever differs between them wrapped in a highlighted
    <span>, so a near-identical pair like "TITANIUM GREY" vs "TITANIUM GRAY"
    -- where reading every character to spot the one that changed is slow --
    instead shows the single differing letter highlighted in both names at a
    glance. Matching stretches are left as plain escaped text either side."""
    a, b = str(a), str(b)
    matcher = difflib.SequenceMatcher(None, a, b, autojunk=False)
    out_a, out_b = [], []
    highlight = 'style="background:#fff3b0;color:#7a4b00;border-radius:2px;"'
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        seg_a = html_escape(a[i1:i2])
        seg_b = html_escape(b[j1:j2])
        if tag == "equal":
            out_a.append(seg_a)
            out_b.append(seg_b)
            continue
        if seg_a:
            out_a.append(f"<span {highlight}>{seg_a}</span>")
        if seg_b:
            out_b.append(f"<span {highlight}>{seg_b}</span>")
    return "".join(out_a), "".join(out_b)


def money(value, currency="GBP"):
    try:
        v = float(value)
    except Exception:
        return str(value)
    sym = CURRENCY_SYMBOLS.get(currency, "")
    if v.is_integer():
        return f"{sym}{int(v):,}"
    return f"{sym}{v:,.2f}"


def display_product(item):
    code = str(item.get("code", "")).strip()
    product = str(item.get("product", "")).strip()
    return f"{code} {product}".strip() if code else product


def display_product_label(item):
    """Like display_product, but with a visible "[Margin VAT]" tag appended
    for a line item ticked Margin VAT -- used everywhere a PO's goods are
    shown to a person (PDF, HTML/plain-text email, PO Quick View, the New PO
    items table) so the flag is never silently invisible on the document
    itself. Deliberately NOT used for the Zoho export's "Item Name" column
    (which still calls display_product directly) -- extra text there could
    stop Zoho matching the line to an existing catalog item, and the Zoho
    export already conveys Margin VAT properly via the Item Tax columns
    (see build_zoho_export_rows)."""
    text = display_product(item)
    if item.get("margin_vat"):
        text = f"{text}  [Margin VAT]"
    return text


def first_name_of(full_name):
    """First token of a person's name -- e.g. 'John Smith' -> 'John'. Used
    for email greetings ('Hi John,') so a supplier's Contact name field can
    still hold their full first and last name (shown as-is everywhere else
    -- the Suppliers list, PO details, etc.) while emails only address them
    by first name, not their surname too."""
    full_name = (full_name or "").strip()
    return full_name.split()[0] if full_name else ""


def supplier_display_name(po):
    return first_name_of(str(po.get("supplier_contact_name", "")).strip())


def validate_po_for_send(po, items):
    """Returns a list of human-readable problems; empty list means OK to send."""
    problems = []
    if not items:
        problems.append("Add at least one item before sending.")
    for i, item in enumerate(items, start=1):
        try:
            if int(item["qty"]) <= 0:
                problems.append(f"Item {i}: quantity must be greater than 0.")
        except Exception:
            problems.append(f"Item {i}: quantity is not a valid number.")
        try:
            if float(item["price"]) < 0:
                problems.append(f"Item {i}: price cannot be negative.")
        except Exception:
            problems.append(f"Item {i}: price is not a valid number.")
        if not str(item.get("product", "")).strip():
            problems.append(f"Item {i}: product name is missing.")
    # Deliberately NOT checking for a supplier email here -- saving or
    # sending a PO shouldn't be blocked just because a supplier's email
    # isn't on file yet (per Yitzi: it "should not require supplier to
    # have a email in backend"). Open in Outlook still fills whatever
    # email IS on file into the To: field; if there isn't one, the draft
    # just opens with To: blank, which is obvious enough to fill in by
    # hand rather than something worth hard-blocking on.
    if not po.get("delivery_address"):
        problems.append("No delivery address is selected.")
    return problems


def parse_item_line(line):
    line = " ".join(line.strip().split())
    if not line:
        return None, "blank"

    # price at end, allow trailing + and optional £ / @
    m_price = re.search(r"(?:@?\s*[£$€]?\s*)(\d+(?:\.\d{1,2})?)\s*[\+\)]*\s*$", line)
    if not m_price:
        return None, "no price found"
    price = float(m_price.group(1))
    head = line[:m_price.start()].strip()

    # qty at start: 3X / x3 / 3 x / 3
    m_qty = re.match(r"^(?:(\d+)\s*[xX]|[xX]\s*(\d+)|(\d+))\b\s*(.*)$", head)
    if not m_qty:
        return None, "no quantity found"
    qty = int(next(g for g in m_qty.groups()[:3] if g))
    rest = m_qty.group(4).strip()
    if not rest:
        return None, "no product found"

    # optional code token at start
    parts = rest.split()
    code = ""
    if parts:
        first = parts[0]
        if "/" in first or re.search(r"[A-Za-z].*\d|\d.*[A-Za-z]", first):
            if len(parts) > 1 and len(first) >= 6:
                code = first
                rest = " ".join(parts[1:])
    return {"qty": qty, "code": code, "product": rest, "price": price}, None


def build_email_intro_html(po, conn=None):
    supplier = html_escape(supplier_display_name(po))
    hello = f"Hi, {supplier}" if supplier else "Hi,"
    intro_line = get_setting(conn, "email_intro_line", "Please see below order") if conn is not None else "Please see below order"
    return (
        f'<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;margin:0 0 12px 0;">{hello}</div>'
        f'<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;margin:0 0 18px 0;">{html_escape(intro_line)}</div>'
    )


def build_email_outro_html(po, include_signature=False, conn=None):
    outro_line = get_setting(conn, "email_outro_line", "Thank you") if conn is not None else "Thank you"
    html = f'<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;margin:14px 0 0 0;">{html_escape(outro_line)}</div>'
    your_name = html_escape(str(po.get("your_name", "")).strip())
    if include_signature and your_name:
        html += (
            '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;margin:14px 0 0 0;">Kind regards</div>'
            f'<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;margin:2px 0 0 0;">{your_name}</div>'
        )
    return html


def build_plain_text(po, include_signature=False, conn=None):
    intro_line = get_setting(conn, "email_intro_line", "Please see below order") if conn is not None else "Please see below order"
    outro_line = get_setting(conn, "email_outro_line", "Thank you") if conn is not None else "Thank you"
    greeting = f"Hi, {supplier_display_name(po)}" if supplier_display_name(po) else "Hi,"
    lines = [
        greeting,
        "",
        intro_line,
        "",
        f"Business name: {po['business_name']}",
        f"Company registration number: {po['company_number']}",
        f"Currency: {po['currency']}",
        f"PO Ref: {po['po_ref']}",
        "",
        "Goods Ordered",
    ]
    for item in po["items"]:
        lines.append(f"{item['qty']}X  {display_product_label(item)}  {money(item['price'], po['currency'])}")
    lines += [
        "",
        f"Total Order Value: {money(po['total'], po['currency'])}",
        "",
        "Delivery Address",
        po["delivery_name"],
        po["delivery_address"],
        "",
        "Invoice Address",
        po["invoice_name"],
        po["invoice_address"],
        "",
        outro_line,
    ]
    your_name = str(po.get("your_name", "")).strip()
    if include_signature and your_name:
        lines += ["", "Kind regards", your_name]
    return "\n".join(lines)


_RCG_LOGO_CACHE = None


def _rcg_asset_path(*parts):
    """Resolves a path to a bundled asset (the RCG logo) both when running
    as a plain .py file and as a PyInstaller --onefile .exe, where bundled
    data files get unpacked into a temp folder at sys._MEIPASS rather than
    living next to the .exe itself -- the same resolution logic as
    po_generator_qt.py's own resource_path(), duplicated here rather than
    imported from there since this module deliberately has zero UI
    framework dependency and must never import po_generator_qt.py."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, *parts)


def rcg_logo_png_bytes():
    """The RCG logo, resized to a sensible email-header width (320px) and
    re-encoded as PNG bytes -- cached after the first call since the file
    never changes at runtime. Used as a cid: inline attachment (see
    RCG_LOGO_CID and _open_outlook_draft_windows's images= parameter) by
    every branded email EXCEPT the two that can also be pasted through
    Outlook's older Word-based clipboard converter (build_html, the PO
    email; build_price_request_summary, the price comparison table) --
    cid: references only ever resolve inside a real message with real
    attachments, which a clipboard paste never has, so those two use a
    plain coloured-text brand mark instead rather than risk a broken-image
    icon. Returns None (never raises) if the bundled asset is somehow
    missing, so a branded email still sends with the rest of its styling
    intact, just without the logo graphic."""
    global _RCG_LOGO_CACHE
    if _RCG_LOGO_CACHE is None:
        try:
            from PIL import Image as PILImage
            import io
            path = _rcg_asset_path("assets", "rcg_logo_full_transparent.png")
            im = PILImage.open(path)
            w, h = im.size
            new_w = 320
            new_h = int(h * new_w / w)
            im2 = im.resize((new_w, new_h), PILImage.LANCZOS)
            buf = io.BytesIO()
            im2.save(buf, format="PNG", optimize=True)
            _RCG_LOGO_CACHE = buf.getvalue()
        except Exception:
            log.exception("Could not load the RCG logo for branded emails")
            _RCG_LOGO_CACHE = False
    return _RCG_LOGO_CACHE or None


def email_brand_header_html(logo_cid=RCG_LOGO_CID):
    """The shared branded header every redesigned email (except the two
    clipboard-exposed ones -- see rcg_logo_png_bytes) opens with: the RCG
    logo referenced via cid:, then a thin three-colour stripe echoing the
    logo's own three droplet colours. Table-based, inline-styled, no
    border-radius/box-shadow -- plain, Outlook-safe HTML throughout."""
    return (
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">'
        f'<tr><td style="padding:26px 36px 18px 36px;">'
        f'<img src="cid:{logo_cid}" width="160" alt="Rose Communications Group" '
        'style="display:block;border:0;height:auto;"></td></tr>'
        '<tr><td style="padding:0;">'
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>'
        f'<td width="33.33%" style="height:4px;background:{RCG_ACCENT_LIGHT};font-size:0;line-height:0;">&nbsp;</td>'
        f'<td width="33.33%" style="height:4px;background:{RCG_INK};font-size:0;line-height:0;">&nbsp;</td>'
        f'<td width="33.34%" style="height:4px;background:{RCG_MAUVE};font-size:0;line-height:0;">&nbsp;</td>'
        '</tr></table></td></tr></table>'
    )


def email_text_brand_header_html(title_text):
    """A logo-free alternative to email_brand_header_html for the two
    builders (build_html, build_price_request_summary) that can also be
    pasted through Outlook's Word-based clipboard converter, where a cid:
    image never resolves. Still carries the brand -- the same three-colour
    stripe, and "ROSE COMMUNICATIONS GROUP" set as small bold navy caption
    text above the email's own title -- just without the raster logo."""
    title_html = (
        f'<h1 style="margin:0 0 4px 0;font-size:20px;color:{RCG_INK};font-weight:bold;">'
        f'{html_escape(title_text)}</h1>'
    ) if title_text else ""
    return (
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>'
        f'<td width="33.33%" style="height:4px;background:{RCG_ACCENT_LIGHT};font-size:0;line-height:0;">&nbsp;</td>'
        f'<td width="33.33%" style="height:4px;background:{RCG_INK};font-size:0;line-height:0;">&nbsp;</td>'
        f'<td width="33.34%" style="height:4px;background:{RCG_MAUVE};font-size:0;line-height:0;">&nbsp;</td>'
        '</tr></table>'
        # Bottom padding here used to be 0, so whatever text followed this
        # header (build_html's "Hi, Supplier" greeting, when title_text is
        # empty) sat right underneath the brand caption line with barely
        # any breathing room -- "when coppying PO to clipboard the spacing
        # needs to be better". 18px gives it real separation without
        # affecting the Zoho export header, which already has its own h1
        # margin doing similar work.
        f"""<div style="padding:20px 36px 18px 36px;font-family:Arial,'Segoe UI',Helvetica,sans-serif;">"""
        f'<div style="font-size:11px;font-weight:bold;letter-spacing:0.06em;color:{RCG_ACCENT};'
        f'text-transform:uppercase;margin:0 0 6px 0;">Rose Communications Group</div>'
        f'{title_html}'
        '</div>'
    )


def email_brand_footer_html():
    """The shared footer slot every redesigned email closes with. Yitzi
    asked for the "Sent by PO Generator on behalf of..." attribution line
    removed -- emails shouldn't advertise the internal software that
    generated them -- so this deliberately returns nothing rather than an
    empty divider; every call site (email_shell_html, build_html,
    build_zoho_export_reminder_email) still calls it so a real footer could
    be reintroduced here in one place later if wanted."""
    return ""


def email_shell_html(body_html, logo_cid=RCG_LOGO_CID, width=640):
    """Wraps a branded email's own body content (the stat strips/tables/etc.
    an individual builder assembles) in the shared outer shell -- branded
    header, a centred fixed-width white card on a soft background, branded
    footer. Every redesigned email (except build_html/build_price_request_
    summary, which use email_text_brand_header_html instead since they
    can't reference a cid: image -- see rcg_logo_png_bytes) is built by
    calling this once around its own assembled body_html."""
    return (
        f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" '
        f'style="background:#f3f1f9;">'
        f'<tr><td align="center" style="padding:24px 12px;">'
        f'<table role="presentation" width="{width}" cellpadding="0" cellspacing="0" border="0" '
        f'style="width:{width}px;max-width:{width}px;background:#ffffff;'
        'font-family:Arial,\'Segoe UI\',Helvetica,sans-serif;">'
        f'<tr><td>{email_brand_header_html(logo_cid)}</td></tr>'
        f'<tr><td style="padding:0 36px;">{body_html}</td></tr>'
        f'<tr><td>{email_brand_footer_html()}</td></tr>'
        '</table></td></tr></table>'
    )


def email_stat_strip_html(stats):
    """A row of KPI stat tiles -- stats is a list of (value_html, label)
    pairs -- alternating faint blue/mauve tints, matching the approved
    design sample. Used by every report-style branded email."""
    cells = []
    for i, (value_html, label) in enumerate(stats):
        bg = RCG_ROW_BLUE if i % 2 == 0 else RCG_ROW_MAUVE
        cells.append(
            f'<td align="center" style="background:{bg};padding:14px 6px;">'
            f'<div style="font-size:18px;font-weight:bold;color:{RCG_INK};">{value_html}</div>'
            f'<div style="font-size:11px;color:{RCG_MUTED};margin-top:2px;">{html_escape(label)}</div></td>'
        )
        if i < len(stats) - 1:
            cells.append('<td width="2" style="font-size:0;line-height:0;">&nbsp;</td>')
    return (
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" '
        f'style="margin-top:20px;">' + '<tr>' + ''.join(cells) + '</tr></table>'
    )


def email_table_html(headers, rows, align=None):
    """A branded data table -- solid navy header row (white uppercase
    text), body rows alternating white/faint-blue, matching the approved
    design sample. headers is a list of strings; rows is a list of lists of
    already-escaped cell HTML (same length as headers); align is an
    optional list of "left"/"right" per column (defaults to left)."""
    align = align or ["left"] * len(headers)
    head_cells = "".join(
        f'<th align="{a}" style="background:{RCG_INK};color:#ffffff;padding:8px 10px;font-size:11px;'
        f'letter-spacing:0.04em;text-transform:uppercase;">{html_escape(h)}</th>'
        for h, a in zip(headers, align)
    )
    body_rows = []
    for i, row in enumerate(rows):
        bg = "#ffffff" if i % 2 == 0 else RCG_ROW_BLUE
        cells = "".join(
            f'<td align="{a}" style="padding:7px 10px;color:#3a3164;background:{bg};'
            f'border-bottom:1px solid {RCG_LINE};">{cell}</td>'
            for cell, a in zip(row, align)
        )
        body_rows.append(f'<tr>{cells}</tr>')
    if not rows:
        body_rows.append(
            f'<tr><td colspan="{len(headers)}" style="padding:10px;color:{RCG_MUTED};background:#ffffff;">'
            'None</td></tr>'
        )
    return (
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" '
        f'style="font-size:13px;border-collapse:collapse;margin-top:8px;">'
        f'<tr>{head_cells}</tr>{"".join(body_rows)}</table>'
    )


def email_section_heading_html(text):
    return f'<h3 style="margin:22px 0 8px 0;font-size:14px;color:{RCG_INK};">{html_escape(text)}</h3>'


def build_html(po, conn=None, embed_logo=True):
    # Like build_price_request_summary, this can also be pasted through
    # Outlook's "Copy for Outlook" clipboard path (Word's own HTML
    # converter) as well as opened directly as a COM draft's .HTMLBody --
    # the clipboard path is the more aggressive, less forgiving of the two,
    # and has a long history of silently dropping a whole inline style
    # attribute it doesn't fully understand. <font color> is the one legacy
    # construct every version of Word's converter has always honored, so
    # every colour that carries real meaning here (the red order quantity,
    # most of all) is set with BOTH <font color> and plain (non-!important)
    # CSS as a redundant pair rather than CSS alone -- true regardless of
    # embed_logo, since the same HTML can end up on either path.
    #
    # embed_logo=True (the default -- matches every other redesigned email)
    # uses the real RCG logo via a cid: inline attachment; the caller
    # (open_outlook_email_windows) must pass the logo bytes through to
    # _open_outlook_draft_windows's images= for that cid: reference to
    # actually resolve. Pass embed_logo=False for any path where it can't --
    # the clipboard-paste copy (_copy_for_outlook, no real attachments at
    # all) and the in-app preview dialogs (a plain QTextEdit doesn't resolve
    # cid: references either) -- which instead get the lighter text-only
    # brand header.
    rows = []
    for i, item in enumerate(po["items"]):
        qty_html = (
            f'<font color="{RED}"><span style="color:{RED};font-weight:bold;">{item["qty"]}X</span></font>'
        )
        bg = "#ffffff" if i % 2 == 0 else RCG_ROW_BLUE
        rows.append(
            f'<tr>'
            f'<td align="center" style="padding:8px 6px;background:{bg};border-bottom:1px solid {RCG_LINE};white-space:nowrap;">{qty_html}</td>'
            f'<td style="padding:8px 6px;background:{bg};border-bottom:1px solid {RCG_LINE};word-break:break-word;color:{RCG_INK};">{html_escape(display_product_label(item))}</td>'
            f'<td align="right" style="padding:8px 6px;background:{bg};border-bottom:1px solid {RCG_LINE};white-space:nowrap;color:{RCG_INK};">{html_escape(money(item["price"], po["currency"]))}</td>'
            f'</tr>'
        )

    def info_card(label, value_html, accent=RCG_ACCENT_LIGHT):
        return (
            f'<td width="49%" style="background:#f9f8fc;border-left:3px solid {accent};'
            f'padding:12px 14px;vertical-align:top;">'
            f'<div style="font-size:10.5px;letter-spacing:0.04em;text-transform:uppercase;'
            f'color:{RCG_MUTED};margin-bottom:4px;">{html_escape(label)}</div>'
            f'<div style="font-size:13.5px;color:{RCG_INK};">{value_html}</div></td>'
        )

    def card_row(left, right):
        return (
            '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" '
            f'style="margin-top:10px;"><tr>{left}'
            '<td width="2%" style="font-size:0;line-height:0;">&nbsp;</td>'
            f'{right}</tr></table>'
        )

    def full_card(label, value_html, accent=RCG_GREEN):
        return (
            '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" '
            'style="margin-top:10px;"><tr>'
            f'<td style="background:#f9f8fc;border-left:3px solid {accent};padding:12px 14px;">'
            f'<div style="font-size:10.5px;letter-spacing:0.04em;text-transform:uppercase;'
            f'color:{RCG_MUTED};margin-bottom:4px;">{html_escape(label)}</div>'
            f'<div style="font-size:16px;font-weight:bold;color:{RCG_INK};">{value_html}</div></td>'
            '</tr></table>'
        )

    def html_address_block(name, address):
        lines = [str(name).strip()] + [str(x).strip() for x in str(address).splitlines() if str(x).strip()]
        parts = []
        for i, ln in enumerate(lines):
            safe = html_escape(ln)
            if i == 0:
                parts.append(f'<div style="margin:0;line-height:1.35;"><strong>{safe}</strong></div>')
            else:
                parts.append(f'<div style="margin:0;line-height:1.35;">{safe}</div>')
        return "".join(parts)

    delivery_html = html_address_block(po["delivery_name"], po["delivery_address"])
    invoice_html = html_address_block(po["invoice_name"], po["invoice_address"])

    intro_html = build_email_intro_html(po, conn=conn)
    outro_html = build_email_outro_html(po, include_signature=False, conn=conn)
    header_html = email_brand_header_html() if embed_logo else email_text_brand_header_html("")
    footer_html = email_brand_footer_html()

    body_html = f"""
    {intro_html}
    <h2 style="margin:6px 0 16px 0;font-size:20px;line-height:1.3;color:{RCG_INK};font-weight:bold;">Purchase Order</h2>
    <table role="presentation" cellspacing="0" cellpadding="0" width="100%" style="border-collapse:collapse;font-size:14px;margin-top:6px;">
      <tr><th align="left" style="background:{RCG_INK};color:#ffffff;padding:8px 6px;font-size:11px;letter-spacing:0.04em;text-transform:uppercase;width:64px;">Qty</th>
          <th align="left" style="background:{RCG_INK};color:#ffffff;padding:8px 6px;font-size:11px;letter-spacing:0.04em;text-transform:uppercase;">Product</th>
          <th align="right" style="background:{RCG_INK};color:#ffffff;padding:8px 6px;font-size:11px;letter-spacing:0.04em;text-transform:uppercase;width:96px;">Price</th>
      </tr>
      {''.join(rows)}
    </table>
    {card_row(
        info_card("PO Ref", f"<strong>{html_escape(po['po_ref'])}</strong>"),
        info_card("Currency", html_escape(po["currency"])),
    )}
    {card_row(
        info_card("Business name", html_escape(po["business_name"])),
        info_card("Company registration number", html_escape(po["company_number"])),
    )}
    {full_card("Total Order Value", f"<strong>{html_escape(money(po['total'], po['currency']))}</strong>")}
    {card_row(
        info_card("Delivery Address", delivery_html, accent=RCG_MAUVE),
        info_card("Invoice Address", invoice_html, accent=RCG_MAUVE),
    )}
    {outro_html}
    """

    return f"""
    <html>
    <body>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#f3f1f9;">
    <tr><td align="center" style="padding:24px 12px;">
    <table role="presentation" width="680" cellpadding="0" cellspacing="0" border="0" style="width:680px;max-width:680px;background:#ffffff;font-family:Arial,'Segoe UI',Helvetica,sans-serif;">
    <tr><td>{header_html}</td></tr>
    <tr><td style="padding:0 36px;">{body_html}</td></tr>
    <tr><td>{footer_html}</td></tr>
    </table>
    </td></tr>
    </table>
    </body>
    </html>
    """


# ============================================================
# 5. Clipboard / Outlook integration (Windows only)
# ============================================================

def _html_entity_escape_non_ascii(text):
    """Replaces every character outside plain ASCII with its numeric HTML
    entity (e.g. '£' -> '&#163;'). Browsers/Outlook render the entity
    back to the real character regardless of what byte encoding the
    surrounding text ends up stored in -- which matters for build_cf_html
    below, since .NET's DataFormats.Html clipboard format does not
    reliably preserve non-ASCII bytes (observed: the £ sign turning into a
    mangled replacement character once pasted into Outlook). Keeping every
    byte in the CF_HTML payload plain ASCII sidesteps that entirely."""
    return "".join(f"&#{ord(ch)};" if ord(ch) > 127 else ch for ch in text)


def build_cf_html(fragment_html):
    fragment_html = _html_entity_escape_non_ascii(fragment_html)
    html = "<html><body><!--StartFragment-->" + fragment_html + "<!--EndFragment--></body></html>"
    header = (
        "Version:0.9\r\n"
        "StartHTML:0000000000\r\n"
        "EndHTML:0000000000\r\n"
        "StartFragment:0000000000\r\n"
        "EndFragment:0000000000\r\n"
    )
    start_html = len(header.encode("utf-8"))
    start_fragment = start_html + html.index("<!--StartFragment-->") + len("<!--StartFragment-->")
    end_fragment = start_html + html.index("<!--EndFragment-->")
    end_html = start_html + len(html.encode("utf-8"))
    header = (
        "Version:0.9\r\n"
        f"StartHTML:{start_html:010d}\r\n"
        f"EndHTML:{end_html:010d}\r\n"
        f"StartFragment:{start_fragment:010d}\r\n"
        f"EndFragment:{end_fragment:010d}\r\n"
    )
    return (header + html).encode("utf-8")


def _run_powershell(script, timeout=25):
    import tempfile
    import subprocess

    # "utf-8-sig" (i.e. with a BOM) matters here, not just "utf-8": classic
    # Windows PowerShell (powershell.exe) only recognises a script file as
    # UTF-8 if it starts with a BOM. Without one it silently falls back to
    # the system's ANSI codepage, which mangles any non-ASCII character in
    # the script (e.g. the em dash in the weekly report's subject line
    # turning into "â€"").
    with tempfile.NamedTemporaryFile("w", suffix=".ps1", delete=False, encoding="utf-8-sig") as f:
        f.write(script)
        script_path = f.name

    # Batch 121: Yitzi noticed a PowerShell window flashing on screen every
    # time this runs -- a bare .exe (no installer) launching a console
    # subprocess normally does show one briefly unless explicitly told not
    # to. -WindowStyle Hidden asks PowerShell itself not to show a window;
    # CREATE_NO_WINDOW (and the STARTUPINFO fallback, belt-and-suspenders
    # for older Windows builds where CREATE_NO_WINDOW alone can still let
    # one flash through) tells Windows not to allocate a console for the
    # child process at all. Both only exist on win32 -- guarded so this
    # still runs (for local testing) everywhere else.
    creationflags = 0
    startupinfo = None
    if sys.platform == "win32":
        creationflags = subprocess.CREATE_NO_WINDOW
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE

    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-STA", "-WindowStyle", "Hidden",
             "-File", script_path],
            capture_output=True,
            text=True,
            timeout=timeout,
            creationflags=creationflags,
            startupinfo=startupinfo,
        )
    finally:
        try:
            Path(script_path).unlink(missing_ok=True)
        except Exception:
            log.exception("Could not remove temp script %s", script_path)

    if result.returncode != 0:
        stderr = (result.stderr or "").strip()
        stdout = (result.stdout or "").strip()
        raise RuntimeError(stderr or stdout or "PowerShell operation failed")
    return result


def set_clipboard_html_windows(html, plain_text):
    if sys.platform != "win32":
        raise RuntimeError("Windows only")

    cf_html = build_cf_html(html).decode("utf-8")

    def ps_escape(s):
        return s.replace("`", "``").replace("$", "`$")

    script = f"""
Add-Type -AssemblyName System.Windows.Forms
$plain = @'
{ps_escape(plain_text)}
'@
$html = @'
{ps_escape(cf_html)}
'@
$data = New-Object System.Windows.Forms.DataObject
$data.SetData([System.Windows.Forms.DataFormats]::Html, $html)
$data.SetData([System.Windows.Forms.DataFormats]::UnicodeText, $plain)
[System.Windows.Forms.Clipboard]::SetDataObject($data, $true)
"""
    _run_powershell(script, timeout=20)


def _effective_pdf_logo_path(conn):
    """The logo to draw on generated PDFs: whatever's explicitly chosen in
    Settings > PDF & Branding, or -- if nothing's been chosen and it hasn't
    been switched off -- the bundled Rose Communications Group logo, so
    PDFs are branded out of the box without needing manual setup. Same
    logic as po_generator_qt.py's own _effective_pdf_logo_path, duplicated
    here (using _rcg_asset_path rather than the UI's resource_path) since
    this module deliberately has zero UI framework dependency."""
    if conn is None:
        return None
    if get_setting(conn, "pdf_logo_disabled", "0") == "1":
        return None
    explicit = get_setting(conn, "pdf_logo_path", "").strip()
    if explicit:
        return explicit
    bundled = _rcg_asset_path("assets", "rcg_logo_full_transparent.png")
    return bundled if os.path.exists(bundled) else None


def _po_pdf_attachment_path(po, conn=None):
    """Builds the branded PO PDF (the same one "Export PDF" produces) to a
    temp file for attaching to the Outlook draft -- confirmed scope: PO
    emails are one of the "content" emails that get a matching PDF
    attachment alongside the new branded HTML design. Returns None (never
    raises) rather than blocking the email draft if PDF generation fails
    for any reason -- a missing attachment is far better than no draft at
    all.

    Uses _named_pdf_temp_path so the file's actual basename on disk --
    which is what Outlook's Attachments.Add displays, not anything set
    programmatically -- reads as the PO ref rather than a random temp name
    (Yitzi: "the PDF should be named with the PO number")."""
    try:
        pdf_path = _named_pdf_temp_path(filename_po_ref(po.get("po_ref", "") or "PO"))
        export_po_pdf(
            pdf_path, po, conn=conn,
            terms_text=get_setting(conn, "pdf_terms_text", "") or None if conn is not None else None,
        )
        return pdf_path
    except Exception:
        log.exception("Could not build the PO PDF attachment -- sending the draft without it")
        return None


def open_outlook_email_windows(po, conn=None):
    # A real Outlook draft can actually carry the logo as a genuine cid:
    # inline attachment, so this uses the same branded header every other
    # redesigned email uses (embed_logo=True, build_html's default) --
    # only the clipboard-paste copy and the in-app preview dialogs fall
    # back to the text-only header, since neither of those can resolve a
    # cid: reference at all.
    html = build_html(po, conn=conn).replace(
        build_email_outro_html(po, include_signature=False, conn=conn),
        build_email_outro_html(po, include_signature=True, conn=conn),
    )
    pdf_path = _po_pdf_attachment_path(po, conn=conn)
    logo_png = rcg_logo_png_bytes()
    images = [(RCG_LOGO_CID, logo_png)] if logo_png else None
    _open_outlook_draft_windows(
        to=po.get("supplier_email", ""), cc=po.get("supplier_cc_emails", ""),
        bcc=po.get("bcc_emails", ""), subject=po.get("po_ref", ""), html_body=html,
        images=images, attachments=[pdf_path] if pdf_path else None,
        from_address=get_setting(conn, "outlook_shared_mailbox", "") if conn is not None else None,
    )


def _open_outlook_draft_windows(to, cc, bcc, subject, html_body, images=None, attachments=None, from_address=None):
    """Shared Outlook-draft opener (Windows only, via COM automation).
    Used both for PO emails and the periodic PO summary report email --
    always opens a draft for the user to review and send themselves; this
    app never sends email on its own.

    images: optional list of (cid, png_bytes) tuples. Each is attached to
    the draft and made available to the HTML body as an inline image via
    <img src="cid:<cid>">, the standard way to embed a picture (e.g. a
    chart) directly in an Outlook email rather than as a data: URI, which
    Outlook's Word-based renderer does not reliably support.

    attachments: optional list of plain file paths (e.g. the daily Zoho
    export workbook) to attach as ordinary, visible attachments -- not
    inline-referenced in the body like images above.

    from_address: optional shared-mailbox email address to send this draft
    from instead of Outlook's own default account (see the
    outlook_shared_mailbox setting). Tries to match it to a real account
    already configured in Outlook first (so it sends as that mailbox
    outright); if it isn't set up as a full account, falls back to
    "on behalf of" that address, which works via ordinary Send As/Send On
    Behalf delegate permission on a shared mailbox without needing it added
    as a separate Outlook profile."""
    if sys.platform != "win32":
        raise RuntimeError("Windows only")

    import tempfile

    def ps_escape(s):
        return s.replace("`", "``").replace("$", "`$")

    with tempfile.NamedTemporaryFile("w", suffix=".html", delete=False, encoding="utf-8-sig") as hf:
        hf.write(html_body)
        html_path = hf.name

    image_paths = []
    attach_lines = []
    for cid, png_bytes in (images or []):
        img_file = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
        img_file.write(png_bytes)
        img_file.close()
        image_paths.append(img_file.name)
        var = f"att_{cid}".replace("-", "_")
        attach_lines.append(
            f"${var} = $mail.Attachments.Add('{ps_escape(img_file.name)}')\n"
            f"${var}.PropertyAccessor.SetProperty($imgPropertySchema, '{ps_escape(cid)}')"
        )
    for i, file_path in enumerate(attachments or []):
        attach_lines.append(f"$mail.Attachments.Add('{ps_escape(str(file_path))}') | Out-Null")
    attach_block = "\n".join(attach_lines)

    from_address = (from_address or "").strip()
    from_block = ""
    if from_address:
        from_block = f"""
$fromAddress = @'
{ps_escape(from_address)}
'@
$matchedAccount = $null
foreach ($acct in $outlook.Session.Accounts) {{
    if ($acct.SmtpAddress -and $acct.SmtpAddress.ToLower() -eq $fromAddress.ToLower()) {{
        $matchedAccount = $acct
        break
    }}
}}
if ($matchedAccount) {{
    $mail.SendUsingAccount = $matchedAccount
}} else {{
    $mail.SentOnBehalfOfName = $fromAddress
}}
"""

    script = f"""
$outlook = New-Object -ComObject Outlook.Application
$mail = $outlook.CreateItem(0)
$mail.To = @'
{ps_escape(to)}
'@
$mail.CC = @'
{ps_escape(cc)}
'@
$mail.BCC = @'
{ps_escape(bcc)}
'@
$mail.Subject = @'
{ps_escape(subject)}
'@
$imgPropertySchema = "http://schemas.microsoft.com/mapi/proptag/0x3712001E"
{attach_block}
{from_block}
$htmlPath = @'
{ps_escape(html_path)}
'@
$htmlBody = Get-Content -Raw -Encoding UTF8 $htmlPath
$mail.HTMLBody = $htmlBody
$mail.Display()
"""
    try:
        _run_powershell(script, timeout=25)
    finally:
        try:
            Path(html_path).unlink(missing_ok=True)
        except Exception:
            log.exception("Could not remove temp html %s", html_path)
        for p in image_paths:
            try:
                Path(p).unlink(missing_ok=True)
            except Exception:
                log.exception("Could not remove temp image %s", p)


def open_report_email_windows(to, cc, subject, html_body, images=None, attachments=None, conn=None):
    _open_outlook_draft_windows(
        to=to, cc=cc, bcc="", subject=subject, html_body=html_body, images=images, attachments=attachments,
        from_address=get_setting(conn, "outlook_shared_mailbox", "") if conn is not None else None,
    )


def open_price_request_email_windows(to, subject, html_body, images=None, conn=None):
    """One supplier's price-request draft -- no cc/bcc, just the supplier's
    own address, addressed to them individually. No PDF attachment -- this
    is the plain outgoing ask, before any pricing has come back, so there's
    nothing yet to summarise (confirmed scope, see build_price_request_summary_
    pdf_pages)."""
    _open_outlook_draft_windows(
        to=to, cc="", bcc="", subject=subject, html_body=html_body, images=images,
        from_address=get_setting(conn, "outlook_shared_mailbox", "") if conn is not None else None,
    )


def open_price_summary_email_windows(to, subject, html_body, images=None, attachments=None, conn=None):
    """The "Open in Outlook" summary draft for a price request's received
    pricing (see build_price_request_summary) -- addressed to whoever
    Yitzi wants to forward the pricing to internally (a staff member), not
    the supplier, so kept as its own function rather than reusing
    open_price_request_email_windows (which is specifically the outgoing
    ask-a-supplier-for-pricing draft)."""
    _open_outlook_draft_windows(
        to=to, cc="", bcc="", subject=subject, html_body=html_body, images=images, attachments=attachments,
        from_address=get_setting(conn, "outlook_shared_mailbox", "") if conn is not None else None,
    )


def open_zoho_export_email_windows(to, cc, subject, html_body, attachment_path, conn=None):
    """The daily Zoho export draft, with the Excel workbook attached. Also
    carries the real cid: logo (see build_zoho_export_reminder_email) as an
    inline image, the same way every other redesigned email's draft does."""
    logo_png = rcg_logo_png_bytes()
    images = [(RCG_LOGO_CID, logo_png)] if logo_png else None
    _open_outlook_draft_windows(
        to=to, cc=cc, bcc="", subject=subject, html_body=html_body,
        images=images, attachments=[attachment_path],
        from_address=get_setting(conn, "outlook_shared_mailbox", "") if conn is not None else None,
    )


def _show_windows_toast(title, message):
    """A simple native Windows notification balloon -- used for the daily
    reminders (Zoho export due, PO report due, suppliers to chase) so they
    show up even when the app itself isn't open. Uses a plain .NET
    NotifyIcon balloon tip via PowerShell, the same lightweight approach
    already used for Outlook automation elsewhere in this file -- no extra
    dependency to bundle, and it works on every supported Windows version."""
    if sys.platform != "win32":
        return

    def ps_escape(s):
        return s.replace("`", "``").replace("$", "`$")

    script = f"""
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$notify = New-Object System.Windows.Forms.NotifyIcon
$notify.Icon = [System.Drawing.SystemIcons]::Information
$notify.Visible = $true
$notify.BalloonTipTitle = @'
{ps_escape(title)}
'@
$notify.BalloonTipText = @'
{ps_escape(message)}
'@
$notify.ShowBalloonTip(15000)
Start-Sleep -Seconds 16
$notify.Dispose()
"""
    try:
        _run_powershell(script, timeout=25)
    except Exception:
        log.exception("Couldn't show a Windows notification: %s / %s", title, message)


def ensure_reminder_tasks_registered(conn):
    """Registers one daily Windows Scheduled Task that re-runs this same
    .exe with --reminder-check at the configured time, so the Zoho export /
    PO report / chase-up reminders can pop up even when the app itself
    isn't open. Only does anything on Windows, only when actually running
    as the packaged .exe (not a plain "python po_generator_qt.py" dev run,
    where scheduling a bare python interpreter wouldn't make sense), and
    only re-registers when the exe's own path has changed since it was
    last set up (e.g. after an update) -- not on every single launch."""
    if sys.platform != "win32" or not getattr(sys, "frozen", False):
        return
    exe_path = sys.executable
    if get_setting(conn, "reminder_tasks_registered_for", "") == exe_path:
        return
    check_time = get_setting(conn, "zoho_reminder_check_time", "09:30")
    try:
        import subprocess
        subprocess.run(
            [
                "schtasks", "/Create", "/F",
                "/SC", "DAILY",
                "/TN", "YJ PO Generator Reminders",
                "/TR", f'"{exe_path}" --reminder-check',
                "/ST", check_time,
            ],
            capture_output=True, text=True, timeout=15,
        )
        set_settings(conn, {"reminder_tasks_registered_for": exe_path})
    except Exception:
        log.exception("Couldn't register the Windows reminder scheduled task")


def run_reminder_check(conn):
    """The lightweight check invoked by the Windows Scheduled Task (see
    ensure_reminder_tasks_registered) -- fires a native popup for whichever
    daily reminder actually has something outstanding right now. Also safe
    to call directly (e.g. for testing) since it just reports what it
    found/fired rather than needing a live Outlook session to run."""
    result = {"zoho_count": 0, "report_due": False, "reports_due": [], "recap_due": False, "followup_count": 0, "notified": []}

    if get_setting(conn, "zoho_export_enabled", "1") == "1":
        zoho_needed = report_pos_needing_zoho_export(conn)
        result["zoho_count"] = len(zoho_needed)
        if zoho_needed:
            n = len(zoho_needed)
            _show_windows_toast(
                "Zoho export due",
                f"{pluralize(n, 'purchase order')} {'needs' if n == 1 else 'need'} exporting to Zoho.",
            )
            result["notified"].append("zoho")

    # Multi-frequency (Batch 99): weekly/monthly/quarterly/yearly are all
    # independent now, so more than one can genuinely be due at once (e.g.
    # right after several get turned on together and more than one has
    # already passed its scheduled day) -- fire one toast per due
    # frequency rather than assuming there's only ever one.
    due_reports = reports_due(conn)
    result["reports_due"] = [d["freq"] for d in due_reports]
    result["report_due"] = bool(due_reports)
    for d in due_reports:
        _show_windows_toast(
            "PO report due",
            f"Your {d['freq']} PO summary report hasn't been sent yet.",
        )
        result["notified"].append(f"report_{d['freq']}")

    recap_due, _rps, _rpe = stock_recap_is_due(conn)
    result["recap_due"] = recap_due
    if recap_due:
        _show_windows_toast(
            "Stock recap due",
            "This week's stock team recap hasn't been sent yet.",
        )
        result["notified"].append("recap")

    followups = report_pos_needing_followup(conn)
    result["followup_count"] = len(followups)
    if followups:
        n_followups = len(followups)
        _show_windows_toast(
            "Suppliers to chase",
            f"{pluralize(n_followups, 'purchase order')} {'needs' if n_followups == 1 else 'need'} chasing up.",
        )
        result["notified"].append("followup")

    return result


def ensure_backup_tasks_registered(conn):
    """Registers two Windows Scheduled Tasks -- one at 10:00, one at 17:00,
    both Monday-Friday only -- that each re-run this same .exe with
    --scheduled-backup-check, so the twice-daily automatic backups happen
    even when the app itself isn't open (exactly the gap
    ensure_reminder_tasks_registered already closes for the report/Zoho/
    chase-up reminders, mirrored here for backups). Always on, no settings
    toggle -- this is the twice-per-weekday backup Yitzi asked for outright,
    not an opt-in extra like the weekly schedule above. Only does anything
    on Windows, only when actually running as the packaged .exe, and only
    re-registers when the exe's own path has changed since it was last set
    up."""
    if sys.platform != "win32" or not getattr(sys, "frozen", False):
        return
    exe_path = sys.executable
    if get_setting(conn, "backup_tasks_registered_for", "") == exe_path:
        return
    try:
        import subprocess
        for task_name, when in (("YJ PO Generator Backup AM", "10:00"), ("YJ PO Generator Backup PM", "17:00")):
            subprocess.run(
                [
                    "schtasks", "/Create", "/F",
                    "/SC", "WEEKLY",
                    "/D", "MON,TUE,WED,THU,FRI",
                    "/TN", task_name,
                    "/TR", f'"{exe_path}" --scheduled-backup-check',
                    "/ST", when,
                ],
                capture_output=True, text=True, timeout=15,
            )
        set_settings(conn, {"backup_tasks_registered_for": exe_path})
    except Exception:
        log.exception("Couldn't register the Windows scheduled backup tasks")


def run_scheduled_backup_check(conn):
    """The lightweight action invoked by the Windows Scheduled Tasks
    registered in ensure_backup_tasks_registered -- just takes a backup and
    exits, no UI, no popup (unlike run_reminder_check, a backup succeeding
    isn't something that needs to interrupt anyone)."""
    return backup_now(reason="scheduled-weekday")


# ============================================================
# 6. PDF writer (dependency-free — no reportlab, so it can never
#    fail with "install reportlab", which was a recurring v1 bug)
# ============================================================

def pdf_escape(s):
    s = str(s)
    return s.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def pdf_currency_text(value, currency="GBP"):
    try:
        v = float(value)
    except Exception:
        return str(value)
    symbol = chr(163) if currency == "GBP" else ("$" if currency == "USD" else chr(128) if currency == "EUR" else "")
    if v.is_integer():
        return f"{symbol}{int(v):,}"
    return f"{symbol}{v:,.2f}"


def _load_image_rgb(path, max_dim=500):
    """Loads an arbitrary image file (PNG/JPG/etc) for embedding in a PDF.
    Returns (width, height, raw_rgb_bytes), or None if the file can't be
    read or Pillow isn't installed (a missing/broken logo should never stop
    a PO PDF from generating -- it just gets skipped). Transparency is
    flattened onto white since the PDF page background is white."""
    try:
        from PIL import Image as PILImage
    except ImportError:
        log.warning("Pillow not installed -- can't embed logo image %s in PDF", path)
        return None
    try:
        img = PILImage.open(path)
        img = img.convert("RGBA")
        bg = PILImage.new("RGB", img.size, (255, 255, 255))
        bg.paste(img, mask=img.split()[3])
        img = bg
        if max(img.size) > max_dim:
            img.thumbnail((max_dim, max_dim), PILImage.LANCZOS)
        return img.size[0], img.size[1], img.tobytes()
    except Exception:
        log.exception("Could not load logo image %s", path)
        return None


def _load_image_rgba(path, max_dim=900):
    """Like _load_image_rgb, but keeps the image's real alpha channel
    instead of flattening it onto white -- returns (width, height,
    raw_rgb_bytes, raw_alpha_bytes), or None on failure/missing Pillow.
    raw_alpha_bytes is a single-byte-per-pixel grayscale mask suitable for
    a PDF /SMask object (see _write_pdf_from_command_pages), so a logo like
    rcg_logo_full_transparent.png renders with its real transparent
    background intact rather than a flattened white card behind it --
    Yitzi: "the logo should not have white background" (batch 92)."""
    try:
        from PIL import Image as PILImage
    except ImportError:
        log.warning("Pillow not installed -- can't embed logo image %s in PDF", path)
        return None
    try:
        img = PILImage.open(path).convert("RGBA")
        if max(img.size) > max_dim:
            img.thumbnail((max_dim, max_dim), PILImage.LANCZOS)
        alpha = img.split()[3].tobytes()
        rgb = img.convert("RGB").tobytes()
        return img.size[0], img.size[1], rgb, alpha
    except Exception:
        log.exception("Could not load logo image %s", path)
        return None


def _whiten_rgba_image(img):
    """Given a (width, height, raw_rgb_bytes, raw_alpha_bytes) tuple from
    _load_image_rgba, returns the same shape (same alpha, so the exact
    same silhouette) with every pixel's RGB forced to solid white. Used to
    draw the RCG logo -- or a user's own custom PDF logo -- in white
    whenever it sits on one of the brand's colored banners instead of a
    white page background, so a logo with dark ink in it (like the RCG
    wordmark) doesn't wash out into near-invisibility against a similarly
    dark banner fill. Yitzi: "change the logo to white whenever it is on a
    blue background so it can be fully seen" (batch 93). Matches the
    pre-made white rcg_icon_only.png already used for the app's own navy
    sidebar -- this just does the same trick in code so it works for any
    logo, not only one pre-exported as a white asset."""
    if not img or len(img) < 4:
        return img
    w, h, rgb_bytes, alpha_bytes = img[0], img[1], img[2], img[3]
    return w, h, b"\xff" * len(rgb_bytes), alpha_bytes


def _load_image_rgb_from_bytes(image_bytes, max_dim=1400):
    """Same as _load_image_rgb, but for an image already in memory (e.g. a
    chart PNG rendered by _render_chart_png) instead of one on disk. Returns
    (width, height, raw_rgb_bytes), or None if Pillow isn't installed or the
    bytes can't be decoded -- a chart that fails to render should never stop
    the rest of the PDF from being generated, it just gets skipped."""
    try:
        from PIL import Image as PILImage
    except ImportError:
        log.warning("Pillow not installed -- can't embed chart image in PDF")
        return None
    try:
        import io
        img = PILImage.open(io.BytesIO(image_bytes))
        img = img.convert("RGBA")
        bg = PILImage.new("RGB", img.size, (255, 255, 255))
        bg.paste(img, mask=img.split()[3])
        img = bg
        if max(img.size) > max_dim:
            img.thumbnail((max_dim, max_dim), PILImage.LANCZOS)
        return img.size[0], img.size[1], img.tobytes()
    except Exception:
        log.exception("Could not load chart image from bytes")
        return None


def hex_to_rgb01(hex_color):
    """'#6498be' -> (0.392, 0.596, 0.745) -- the 0..1 float triples the raw
    PDF content-stream 'rg'/'RG' color operators expect."""
    h = str(hex_color).lstrip("#")
    return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))


def save_simple_multipage_pdf(path, po, logo_path=None, terms_text=None):
    logo_img = _load_image_rgb(logo_path) if logo_path else None
    pages = build_pdf_pages(po, logo_img=logo_img, terms_text=terms_text)
    if not pages:
        raise RuntimeError("No PDF pages were generated")
    _write_pdf_from_command_pages(path, pages, logo=logo_img)


def _write_pdf_from_command_pages(path, pages, page_size=(595, 842), logo=None, images=None):
    """Shared low-level PDF object writer. pages is a list of pages, each a
    list of PDF content-stream command strings (as produced by
    build_pdf_pages for a PO, or build_report_pdf_pages for a report table).
    page_size must match whatever coordinate system those commands were
    written in (build_pdf_pages uses the default A4 portrait; report PDFs
    are landscape). No external PDF library involved -- this writes the raw
    PDF structure directly, same as the original PO PDF writer always has.

    logo: optional (width, height, raw_rgb_bytes) tuple -- if given, it's
    embedded as an Image XObject named /Logo, available to any page's
    content stream via the 'q ... cm /Logo Do Q' pattern build_pdf_pages
    uses to place it.

    images: optional {name: (width, height, raw_rgb_bytes)} dict for
    embedding any number of additional images (e.g. the scorecard PDF's
    chart PNGs) -- each is available to any page's content stream as
    '/{name} Do', same placement pattern as the logo. An entry may instead
    be a 4-tuple (width, height, raw_rgb_bytes, raw_alpha_bytes) -- see
    _load_image_rgba -- in which case a real /SMask soft-mask object is
    embedded alongside it, so the image keeps its actual transparency
    instead of needing to be flattened onto a solid color first."""
    page_w, page_h = page_size
    objects = []

    def add_object(data_bytes):
        objects.append(data_bytes)
        return len(objects)

    font_regular = add_object(
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>"
    )
    font_bold = add_object(
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>"
    )

    all_images = dict(images or {})
    if logo:
        all_images["Logo"] = logo

    image_ids = {}
    for name, img in all_images.items():
        if not img:
            continue
        img_w, img_h, rgb_bytes = img[0], img[1], img[2]
        alpha_bytes = img[3] if len(img) > 3 else None
        smask_ref = ""
        if alpha_bytes:
            alpha_compressed = zlib.compress(alpha_bytes, 6)
            smask_obj = (
                f"<< /Type /XObject /Subtype /Image /Width {img_w} /Height {img_h} "
                f"/ColorSpace /DeviceGray /BitsPerComponent 8 /Filter /FlateDecode "
                f"/Length {len(alpha_compressed)} >>\nstream\n"
            ).encode("ascii") + alpha_compressed + b"\nendstream"
            smask_id = add_object(smask_obj)
            smask_ref = f" /SMask {smask_id} 0 R"
        compressed = zlib.compress(rgb_bytes, 6)
        img_obj = (
            f"<< /Type /XObject /Subtype /Image /Width {img_w} /Height {img_h} "
            f"/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode"
            f"{smask_ref} /Length {len(compressed)} >>\nstream\n"
        ).encode("ascii") + compressed + b"\nendstream"
        image_ids[name] = add_object(img_obj)

    content_ids = []
    for cmds in pages:
        content = ("\n".join(cmds) + "\n").encode("latin-1", "replace")
        stream = b"<< /Length " + str(len(content)).encode("ascii") + b" >>\nstream\n" + content + b"endstream"
        content_ids.append(add_object(stream))

    pages_obj = add_object(b"<< /Type /Pages /Kids [] /Count 0 >>")
    page_ids = []
    for cid in content_ids:
        resources = f"/Font << /F1 {font_regular} 0 R /F2 {font_bold} 0 R >>"
        if image_ids:
            xobjects = " ".join(f"/{name} {oid} 0 R" for name, oid in image_ids.items())
            resources += f" /XObject << {xobjects} >>"
        page_obj = (
            f"<< /Type /Page /Parent {pages_obj} 0 R /MediaBox [0 0 {page_w} {page_h}] "
            f"/Resources << {resources} >> "
            f"/Contents {cid} 0 R >>"
        ).encode("ascii")
        page_ids.append(add_object(page_obj))

    kids = "[ " + " ".join(f"{pid} 0 R" for pid in page_ids) + " ]"
    objects[pages_obj - 1] = f"<< /Type /Pages /Kids {kids} /Count {len(page_ids)} >>".encode("ascii")
    catalog_id = add_object(f"<< /Type /Catalog /Pages {pages_obj} 0 R >>".encode("ascii"))

    out = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets = [0]
    for i, obj in enumerate(objects, start=1):
        offsets.append(len(out))
        out += f"{i} 0 obj\n".encode("ascii") + obj + b"\nendobj\n"

    startxref = len(out)
    out += f"xref\n0 {len(objects)+1}\n".encode("ascii")
    out += b"0000000000 65535 f \n"
    for off in offsets[1:]:
        out += f"{off:010d} 00000 n \n".encode("ascii")
    out += (
        f"trailer\n<< /Size {len(objects)+1} /Root {catalog_id} 0 R >>\n"
        f"startxref\n{startxref}\n%%EOF"
    ).encode("ascii")

    with open(path, "wb") as f:
        f.write(out)


def build_pdf_pages(po, logo_img=None, terms_text=None):
    PAGE_W, PAGE_H = 595, 842
    margin = 18
    label_w = 120
    total_w = PAGE_W - margin * 2
    value_w = total_w - label_w

    def new_page():
        return {"cmds": [], "top": margin}

    pages = [new_page()]
    current = pages[-1]

    def use_page():
        return current

    def pdf_y(top, height=0):
        return PAGE_H - top - height

    def rect(x, top, w, h, line_w=1):
        p = use_page()
        y = pdf_y(top, h)
        p["cmds"].append(f"q {line_w} w 0 0 0 RG {x} {y} {w} {h} re S Q")

    def fill_rect(x, top, w, h, rgb):
        p = use_page()
        y = pdf_y(top, h)
        r, g, b = rgb
        p["cmds"].append(f"q {r:.3f} {g:.3f} {b:.3f} rg {x} {y} {w} {h} re f Q")

    def line(x1, top1, x2, top2, line_w=1):
        p = use_page()
        y1 = PAGE_H - top1
        y2 = PAGE_H - top2
        p["cmds"].append(f"q {line_w} w 0 0 0 RG {x1} {y1} m {x2} {y2} l S Q")

    def text(x, top, s, size=10, bold=False, color=(0, 0, 0)):
        p = use_page()
        r, g, b = color
        font = "/F2" if bold else "/F1"
        y = PAGE_H - top - size
        p["cmds"].append(
            f"BT {r:.3f} {g:.3f} {b:.3f} rg {font} {size} Tf 1 0 0 1 {x} {y} Tm ({pdf_escape(s)}) Tj ET"
        )

    def wrap_lines(s, width, size=10):
        words = str(s).split()
        if not words:
            return [""]
        limit = max(6, int(width / (size * 0.52)))
        lines = []
        cur = words[0]
        for w in words[1:]:
            test = cur + " " + w
            if len(test) <= limit:
                cur = test
            else:
                lines.append(cur)
                cur = w
        lines.append(cur)
        return lines

    def draw_wrapped(x, top, width, s, size=10, bold=False, color=(0, 0, 0), leading=12):
        lines = wrap_lines(s, width, size)
        for i, ln in enumerate(lines):
            text(x, top + i * leading, ln, size=size, bold=bold, color=color)
        return len(lines) * leading

    top = current["top"]

    def ensure_room(height_needed):
        nonlocal current, top
        if top + height_needed <= PAGE_H - margin:
            return
        current = new_page()
        pages.append(current)
        top = current["top"]

    if logo_img:
        logo_w, logo_h, _rgb = logo_img
        draw_h = min(60, 60 * logo_h / max(logo_w, 1))
        draw_w = draw_h * logo_w / max(logo_h, 1)
        lx = PAGE_W - margin - draw_w
        ly = top
        current["cmds"].append(
            f"q {draw_w:.2f} 0 0 {draw_h:.2f} {lx:.2f} {pdf_y(ly, draw_h):.2f} cm /Logo Do Q"
        )
        top += draw_h + 10

    header_rows = [
        ("Business name", po.get("business_name", "")),
        ("Company registration number", po.get("company_number", "")),
        ("Currency", po.get("currency", "")),
        ("PO Ref", po.get("po_ref", "")),
    ]
    for label, value in header_rows:
        h = 42 if label == "Company registration number" else 34
        ensure_room(h)
        rect(margin, top, label_w, h)
        rect(margin + label_w, top, value_w, h)
        if label == "Company registration number":
            draw_wrapped(margin + 6, top + 8, label_w - 12, label, size=9, bold=True, leading=11)
            text(margin + label_w + 10, top + 14, value, size=10, bold=True)
        else:
            draw_wrapped(margin + 6, top + 10, label_w - 12, label, size=10, bold=True)
            text(margin + label_w + 10, top + 11, value, size=10, bold=True)
        top += h

    items = list(po.get("items", []))

    # Geometry for the goods table is fixed regardless of how many rows end
    # up on a given page, so it can be computed once up front.
    goods_header_h = 38
    gx = margin + label_w + 10
    gw = value_w - 20
    q_w = 42
    price_w = 90
    product_w = gw - q_w - price_w - 4

    def measure_row_height(item):
        product_lines = wrap_lines(display_product_label(item), product_w - 12, 10)
        return max(28, 8 + len(product_lines) * 11 + 6)

    def draw_goods_block(goods_top, rows, continued):
        """Draws ONE self-contained 'Goods Ordered' box: outer label/value
        borders sized to exactly match the header bar plus every row inside
        them, so nothing overflows past the box the way a fixed-height outer
        border used to (rows used to be drawn inset and below a box that was
        only ever as tall as the header, making them look like they'd spilled
        out of the table). Returns the total height this block used."""
        total_h = goods_header_h + sum(h for _, h in rows)
        rect(margin, goods_top, label_w, total_h)
        rect(margin + label_w, goods_top, value_w, total_h)
        label = "Goods Ordered (continued)" if continued else "Goods Ordered"
        draw_wrapped(margin + 6, goods_top + 11, label_w - 12, label, size=10, bold=True)

        fill_rect(gx, goods_top + 6, gw, 26, (0.81, 0.91, 0.95))
        rect(gx, goods_top + 6, gw, 26)
        line(gx + q_w, goods_top + 6, gx + q_w, goods_top + 32)
        line(gx + q_w + product_w, goods_top + 6, gx + q_w + product_w, goods_top + 32)
        text(gx + 9, goods_top + 14, "QTY", 10, True)
        text(gx + q_w + 8, goods_top + 14, "PRODUCT", 10, True)
        text(gx + q_w + product_w + 18, goods_top + 14, "PRICE", 10, True)

        row_top = goods_top + goods_header_h
        for item, row_h in rows:
            rect(gx, row_top, gw, row_h)
            line(gx + q_w, row_top, gx + q_w, row_top + row_h)
            line(gx + q_w + product_w, row_top, gx + q_w + product_w, row_top + row_h)
            if item:
                qty_y = row_top + max(8, (row_h - 18) / 2)
                text(gx + 7, qty_y, f"{item.get('qty', '')}X", 10, True, (0.83, 0.0, 0.0))
                draw_wrapped(gx + q_w + 6, row_top + 8, product_w - 12, display_product_label(item), size=10, bold=False, leading=11)
                price_str = pdf_currency_text(item.get("price", 0), po.get("currency", "GBP"))
                price_x = gx + q_w + product_w + max(4, price_w - len(price_str) * 5 - 6)
                text(price_x, qty_y, price_str, 10, False)
            row_top += row_h
        return total_h

    if not items:
        ensure_room(goods_header_h + 28)
        block_h = draw_goods_block(top, [({}, 28)], continued=False)
        top += block_h
    else:
        idx = 0
        continued = False
        while idx < len(items):
            ensure_room(goods_header_h + 28)
            # Reserve roughly the same footer space (total + both address
            # blocks) the old fixed-row-count logic reserved, but now measure
            # each row's real wrapped height instead of assuming a flat 40px,
            # so the block this page draws is the block it actually fits.
            max_avail = PAGE_H - margin - (top + goods_header_h) - 170
            rows = []
            consumed = 0
            j = idx
            while j < len(items):
                row_h = measure_row_height(items[j])
                if rows and consumed + row_h > max_avail:
                    break
                rows.append((items[j], row_h))
                consumed += row_h
                j += 1
            if not rows:
                # Not even one row fits in the space left on this page —
                # start a fresh page rather than drawing a box with nothing
                # in it that would just get overwritten below.
                current = new_page()
                pages.append(current)
                top = current["top"]
                continued = True
                continue

            block_h = draw_goods_block(top, rows, continued=continued)
            top += block_h
            idx = j
            continued = idx < len(items)
            if continued:
                current = new_page()
                pages.append(current)
                top = current["top"]

    ensure_room(34)
    rect(margin, top, label_w, 34)
    rect(margin + label_w, top, value_w, 34)
    draw_wrapped(margin + 6, top + 10, label_w - 12, "Total Order Value", size=10, bold=True)
    text(margin + label_w + 10, top + 10, pdf_currency_text(po.get("total", 0), po.get("currency", "GBP")), 10, True)
    top += 34

    def address_block(label, name, address):
        nonlocal current, top
        lines = [str(name).strip()] + [ln.strip() for ln in str(address).splitlines() if ln.strip()]
        h = max(86, 18 + len(lines) * 16)
        ensure_room(h)
        rect(margin, top, label_w, h)
        rect(margin + label_w, top, value_w, h)
        draw_wrapped(margin + 6, top + h / 2 - 6, label_w - 12, label, size=10, bold=True)
        cur = top + 12
        for i, ln in enumerate(lines):
            text(margin + label_w + 10, cur, ln, 10, i == 0)
            cur += 16
        top += h

    address_block("Delivery Address", po.get("delivery_name", ""), po.get("delivery_address", ""))
    address_block("Invoice Address", po.get("invoice_name", ""), po.get("invoice_address", ""))

    if terms_text and terms_text.strip():
        ensure_room(24)
        top += 8
        line(margin, top, margin + total_w, top)
        top += 12
        for para in terms_text.splitlines():
            if not para.strip():
                top += 6
                continue
            lines = wrap_lines(para, total_w, 8)
            h = len(lines) * 10
            ensure_room(h)
            draw_wrapped(margin, top, total_w, para, size=8, bold=False, leading=10)
            top += h

    return [pg["cmds"] for pg in pages]


def build_report_pdf_pages(title, columns, rows, subtitle=""):
    """A simple paginated table PDF for exporting any report. columns is the
    same (header, key, formatter_or_None) shape used to fill the report
    tables in the UI; rows is a list of dicts. Landscape, since reports
    often have several columns. Reuses the same dependency-free PDF
    approach as the PO PDF writer."""
    PAGE_W, PAGE_H = 842, 595
    margin = 28
    pages = [{"cmds": []}]
    current = pages[-1]

    def text(x, top, s, size=9, bold=False):
        font = "/F2" if bold else "/F1"
        y = PAGE_H - top - size
        current["cmds"].append(f"BT 0 0 0 rg {font} {size} Tf 1 0 0 1 {x} {y} Tm ({pdf_escape(s)}) Tj ET")

    def hline(x1, top, x2, line_w=0.75):
        y = PAGE_H - top
        current["cmds"].append(f"q {line_w} w 0.6 0.6 0.6 RG {x1} {y} m {x2} {y} l S Q")

    usable_w = PAGE_W - margin * 2
    col_w = usable_w / max(len(columns), 1)
    row_h = 18
    state = {"top": margin}

    def new_page():
        nonlocal current
        current = {"cmds": []}
        pages.append(current)
        state["top"] = margin

    def draw_header_block():
        for i, (header, _key, _fmt) in enumerate(columns):
            text(margin + i * col_w, state["top"], str(header)[:40], size=9, bold=True)
        state["top"] += 12
        hline(margin, state["top"], margin + usable_w)
        state["top"] += 8

    text(margin, state["top"], title, size=14, bold=True)
    state["top"] += 20
    if subtitle:
        text(margin, state["top"], subtitle, size=9)
        state["top"] += 16
    state["top"] += 6
    draw_header_block()

    max_chars = max(6, int(col_w / 5.2))
    for row in rows:
        if state["top"] + row_h > PAGE_H - margin:
            new_page()
            draw_header_block()
        for i, (header, key, fmt) in enumerate(columns):
            raw = row.get(key, "")
            val = fmt(raw, row) if fmt else raw
            cell = str(val)
            if len(cell) > max_chars:
                cell = cell[: max_chars - 1] + "…"
            text(margin + i * col_w, state["top"], cell, size=8.5)
        state["top"] += row_h

    return [pg["cmds"] for pg in pages]


def export_report_to_pdf(path, title, columns, rows, subtitle=""):
    pages = build_report_pdf_pages(title, columns, rows, subtitle=subtitle)
    _write_pdf_from_command_pages(path, pages, page_size=(842, 595))


def build_supplier_scorecard_pdf_pages(company_name, data, currency="GBP", generated_at=None):
    """A branded, chart-led supplier scorecard PDF: a colour banner, headline
    numbers as stat cards, then each section (monthly spend, top products,
    price competitiveness, issues) with a chart for a quick visual read
    followed by the exact figures in a table underneath -- portrait,
    paginating wherever a section runs off the page. Reuses the same
    dependency-free PDF primitives as the PO and report PDF writers (see
    build_pdf_pages / build_report_pdf_pages) for text/shapes, and the same
    matplotlib chart renderer already used for the periodic email report
    (see _render_chart_png) for the charts, in the app's own brand colours.

    Returns (pages, images): pages is a list of pages, each a list of
    content-stream command strings; images is a {name: (w, h, rgb_bytes)}
    dict of every chart embedded, for the caller to pass through to
    _write_pdf_from_command_pages(..., images=images)."""
    PAGE_W, PAGE_H = 595, 842
    margin = 36
    content_w = PAGE_W - margin * 2

    # Batch 93: these used to be hardcoded hex literals of their own,
    # independent of RCG_INK/RCG_ACCENT/etc -- now sourced from the same
    # three module-level brand constants every other branded PDF/email
    # uses, so a color change (like Yitzi's exact-hex request) only ever
    # has to happen in one place.
    NAVY = hex_to_rgb01(RCG_INK)
    MUTED = hex_to_rgb01(RCG_MUTED)
    ACCENT = hex_to_rgb01(RCG_ACCENT_LIGHT)
    ACCENT_DARK = hex_to_rgb01(RCG_ACCENT)
    SIDEBAR = hex_to_rgb01(RCG_INK)
    # Tidy-up-pass cleanup (post-Batch-151, static-analysis finding): this
    # scorecard has no alternating-row list the way build_product_scorecard_
    # pdf_pages does (which is what actually uses its own MAUVE constant,
    # just below) -- this one was dead leftover from the shared boilerplate
    # both PDF builders start from, never referenced anywhere in this
    # function. Removed rather than left computing something nothing reads.
    BORDER = hex_to_rgb01(RCG_LINE)
    CARD_BG = hex_to_rgb01("#f4f6fa")
    WHITE = (1, 1, 1)
    LIGHT_ACCENT = hex_to_rgb01(RCG_ACCENT_LIGHT)

    pages = [{"cmds": []}]
    current = pages[-1]
    images = {}
    state = {"top": margin}

    def new_page():
        nonlocal current
        current = {"cmds": []}
        pages.append(current)
        state["top"] = margin

    def ensure_room(h):
        if state["top"] + h > PAGE_H - margin:
            new_page()

    def pdf_y(top, height=0):
        return PAGE_H - top - height

    def text(x, top, s, size=10, bold=False, color=NAVY):
        r, g, b = color
        font = "/F2" if bold else "/F1"
        y = PAGE_H - top - size
        current["cmds"].append(
            f"BT {r:.3f} {g:.3f} {b:.3f} rg {font} {size} Tf 1 0 0 1 {x} {y} Tm ({pdf_escape(s)}) Tj ET"
        )

    def fill_rect(x, top, w, h, color):
        r, g, b = color
        y = pdf_y(top, h)
        current["cmds"].append(f"q {r:.3f} {g:.3f} {b:.3f} rg {x:.2f} {y:.2f} {w:.2f} {h:.2f} re f Q")

    def hline(color=BORDER, w=0.75):
        r, g, b = color
        y = PAGE_H - state["top"]
        current["cmds"].append(f"q {w} w {r:.3f} {g:.3f} {b:.3f} RG {margin} {y} m {PAGE_W - margin} {y} l S Q")

    def heading(s):
        ensure_room(32)
        text(margin, state["top"], s, size=13, bold=True, color=ACCENT_DARK)
        state["top"] += 16
        hline(color=ACCENT, w=1.5)
        state["top"] += 12

    def row(cells, widths, size=9, bold=False, color=NAVY):
        ensure_room(16)
        x = margin
        for cell, w in zip(cells, widths):
            text(x, state["top"], str(cell)[:60], size=size, bold=bold, color=color)
            x += w
        state["top"] += 15

    def add_chart(kind, labels, values, colors=None, figsize=(7.4, 2.7), money=False):
        """Renders a chart via the shared matplotlib helper and registers it
        for embedding. Returns (name, aspect_h_over_w), or None if it
        couldn't be rendered (missing Pillow, bad data, etc) -- callers skip
        drawing it in that case rather than losing the rest of the PDF.
        money=True formats a bar chart's value axis with the currency
        symbol (e.g. "£25,000" instead of a bare "25000")."""
        try:
            symbol = CURRENCY_SYMBOLS.get(currency, "") if money else ""
            png_bytes = _render_chart_png(kind, labels, values, colors=colors, figsize=figsize, currency_symbol=symbol)
            img = _load_image_rgb_from_bytes(png_bytes)
        except Exception:
            log.exception("Could not render scorecard chart (%s)", kind)
            return None
        if not img:
            return None
        name = f"Chart{len(images) + 1}"
        images[name] = img
        return name, figsize[1] / figsize[0]

    def draw_chart(chart, width=content_w, extra_bottom=14):
        if not chart:
            return
        name, aspect = chart
        draw_h = width * aspect
        ensure_room(draw_h + extra_bottom)
        y = pdf_y(state["top"], draw_h)
        current["cmds"].append(f"q {width:.2f} 0 0 {draw_h:.2f} {margin:.2f} {y:.2f} cm /{name} Do Q")
        state["top"] += draw_h + extra_bottom

    # --- banner --------------------------------------------------------
    banner_h = 92
    fill_rect(0, 0, PAGE_W, banner_h, SIDEBAR)
    text(margin, 24, "SUPPLIER SCORECARD", size=10, bold=True, color=LIGHT_ACCENT)
    text(margin, 44, company_name, size=21, bold=True, color=WHITE)
    generated = (generated_at or datetime.now().isoformat(timespec="seconds"))[:10]
    text(margin, 70, f"Generated {generated}", size=8, bold=False, color=LIGHT_ACCENT)
    state["top"] = banner_h + 22

    # --- headline stat cards --------------------------------------------
    cards = [
        ("Orders", str(data.get("po_count", 0))),
        ("Total spend", money(data.get("total_spend", 0), currency)),
        ("Avg order value", money(data.get("avg_po_value", 0), currency)),
        ("First order", (data.get("first_order_at") or "-")[:10]),
        ("Last order", (data.get("last_order_at") or "-")[:10]),
    ]
    gap = 10
    card_w = (content_w - gap * (len(cards) - 1)) / len(cards)
    card_h = 56
    ensure_room(card_h + 20)
    x = margin
    for label, value in cards:
        fill_rect(x, state["top"], card_w, card_h, CARD_BG)
        fill_rect(x, state["top"], card_w, 3, ACCENT)
        text(x + 8, state["top"] + 24, value[:16], size=12, bold=True, color=NAVY)
        text(x + 8, state["top"] + 42, label, size=7.5, bold=False, color=MUTED)
        x += card_w + gap
    state["top"] += card_h + 20

    # --- monthly spend ----------------------------------------------------
    heading("Monthly spend")
    monthly = data.get("monthly", [])
    if monthly:
        chart = add_chart(
            "bar",
            [m["period"] for m in monthly],
            [m["total"] for m in monthly],
            colors=["#6498be"],
            money=True,
        )
        draw_chart(chart)
        row(["Month", "Orders", "Total"], [120, 100, 150], bold=True, color=MUTED)
        for m in monthly:
            row([m["period"], str(m["po_count"]), money(m["total"], currency)], [120, 100, 150])
    else:
        row(["No orders yet."], [400], color=MUTED)
    state["top"] += 10

    # --- top products -------------------------------------------------
    heading("Top products")
    top_products = data.get("top_products", [])
    if top_products:
        chart_products = top_products[:8]
        chart = add_chart(
            "bar",
            [p["product"][:22] for p in chart_products],
            [p["total_spend"] for p in chart_products],
            colors=["#b079ad"],
            figsize=(7.4, 3.3),
            money=True,
        )
        draw_chart(chart)
        row(["Product", "Qty", "Times ordered", "Total spend"], [210, 60, 100, 110], bold=True, color=MUTED)
        for p in top_products[:20]:
            row(
                [p["product"], str(p["total_qty"]), str(p["times_ordered"]), money(p["total_spend"], currency)],
                [210, 60, 100, 110],
            )
    else:
        row(["No products yet."], [400], color=MUTED)
    state["top"] += 10

    # --- price competitiveness -----------------------------------------
    heading("Price competitiveness")
    price = data.get("price_competitiveness") or {}
    compared_count = price.get("compared_count", 0)
    if compared_count:
        cheapest_count = price.get("cheapest_count", 0)
        pct = price.get("pct_cheapest")
        chart = add_chart(
            "pie",
            ["Cheapest", "Not cheapest"],
            [cheapest_count, max(compared_count - cheapest_count, 0)],
            colors=["#1e9e6b", "#d64545"],
            figsize=(3.6, 3.0),
        )
        draw_chart(chart, width=200)
        window_months = price.get('window_months', 12)
        row([f"Cheapest on {cheapest_count} of {pluralize(compared_count, 'shared product')} "
             f"({pct}%), last {pluralize(window_months, 'month')}."], [480], color=MUTED)
        state["top"] += 4
        row(["Product", "Their price", "Best other price", "Difference"], [180, 100, 110, 100],
            bold=True, color=MUTED)
        for r in price.get("rows", [])[:20]:
            row(
                [r["product"], money(r["my_price"], currency), money(r["best_other_price"], currency),
                 money(r["difference"], currency)],
                [180, 100, 110, 100],
            )
    else:
        row(["Not enough overlapping order history with other suppliers yet."], [480], color=MUTED)
    state["top"] += 10

    # --- issues logged ---------------------------------------------------
    heading("Issues logged")
    issues = data.get("issues", [])
    if issues:
        issue_counts = data.get("issue_counts") or []
        if issue_counts:
            chart = add_chart(
                "pie",
                [cat for cat, _count in issue_counts],
                [count for _cat, count in issue_counts],
                figsize=(3.6, 3.0),
            )
            draw_chart(chart, width=200)
        for issue in issues:
            # the short subject line, not the full note, is what goes in this
            # fixed-width table -- row() hard-truncates at 60 characters, and
            # a subject is guaranteed to already fit (see _fallback_issue_subject
            # for issues logged before this field existed).
            row(
                [(issue.get("at") or "")[:10], issue.get("category") or "-",
                 issue.get("subject") or issue.get("note", "")],
                [70, 100, 300],
            )
    else:
        row(["No issues logged."], [400], color=MUTED)
    state["top"] += 10

    # --- savings recorded --------------------------------------------
    heading("Savings recorded")
    savings_card_w = 200
    ensure_room(card_h)
    fill_rect(margin, state["top"], savings_card_w, card_h, CARD_BG)
    fill_rect(margin, state["top"], savings_card_w, 3, hex_to_rgb01("#1e9e6b"))
    text(margin + 8, state["top"] + 24, money(data.get("savings_total", 0), currency), size=14, bold=True, color=NAVY)
    text(margin + 8, state["top"] + 42, "Total savings recorded", size=7.5, bold=False, color=MUTED)
    state["top"] += card_h

    return [pg["cmds"] for pg in pages], images


def export_supplier_scorecard_to_pdf(path, company_name, data, currency="GBP"):
    pages, images = build_supplier_scorecard_pdf_pages(company_name, data, currency=currency)
    _write_pdf_from_command_pages(path, pages, page_size=(595, 842), images=images)


def build_product_scorecard_pdf_pages(product_name, data, currency="GBP", generated_at=None):
    """A branded, chart-led product scorecard PDF -- the same layout as
    build_supplier_scorecard_pdf_pages, but for one product: headline stat
    cards, monthly spend, spend by supplier (so it's clear who this has
    actually been bought from), and price history over time. Returns
    (pages, images) the same way -- see build_supplier_scorecard_pdf_pages
    for the shared PDF-writing details."""
    PAGE_W, PAGE_H = 595, 842
    margin = 36
    content_w = PAGE_W - margin * 2

    # Batch 93: sourced from the same three module-level brand constants
    # every other branded PDF/email uses now -- see the note in
    # build_supplier_scorecard_pdf_pages just above.
    NAVY = hex_to_rgb01(RCG_INK)
    MUTED = hex_to_rgb01(RCG_MUTED)
    ACCENT = hex_to_rgb01(RCG_ACCENT_LIGHT)
    ACCENT_DARK = hex_to_rgb01(RCG_ACCENT)
    SIDEBAR = hex_to_rgb01(RCG_INK)
    CARD_BG = hex_to_rgb01("#f4f6fa")
    WHITE = (1, 1, 1)
    LIGHT_ACCENT = hex_to_rgb01(RCG_ACCENT_LIGHT)

    pages = [{"cmds": []}]
    current = pages[-1]
    images = {}
    state = {"top": margin}

    def new_page():
        nonlocal current
        current = {"cmds": []}
        pages.append(current)
        state["top"] = margin

    def ensure_room(h):
        if state["top"] + h > PAGE_H - margin:
            new_page()

    def pdf_y(top, height=0):
        return PAGE_H - top - height

    def text(x, top, s, size=10, bold=False, color=NAVY):
        r, g, b = color
        font = "/F2" if bold else "/F1"
        y = PAGE_H - top - size
        current["cmds"].append(
            f"BT {r:.3f} {g:.3f} {b:.3f} rg {font} {size} Tf 1 0 0 1 {x} {y} Tm ({pdf_escape(s)}) Tj ET"
        )

    def fill_rect(x, top, w, h, color):
        r, g, b = color
        y = pdf_y(top, h)
        current["cmds"].append(f"q {r:.3f} {g:.3f} {b:.3f} rg {x:.2f} {y:.2f} {w:.2f} {h:.2f} re f Q")

    def hline(color=hex_to_rgb01("#dde2ea"), w=0.75):
        r, g, b = color
        y = PAGE_H - state["top"]
        current["cmds"].append(f"q {w} w {r:.3f} {g:.3f} {b:.3f} RG {margin} {y} m {PAGE_W - margin} {y} l S Q")

    def heading(s):
        ensure_room(32)
        text(margin, state["top"], s, size=13, bold=True, color=ACCENT_DARK)
        state["top"] += 16
        hline(color=ACCENT, w=1.5)
        state["top"] += 12

    def row(cells, widths, size=9, bold=False, color=NAVY):
        ensure_room(16)
        x = margin
        for cell, w in zip(cells, widths):
            text(x, state["top"], str(cell)[:60], size=size, bold=bold, color=color)
            x += w
        state["top"] += 15

    def add_chart(kind, labels, values, colors=None, figsize=(7.4, 2.7), money=False):
        try:
            symbol = CURRENCY_SYMBOLS.get(currency, "") if money else ""
            png_bytes = _render_chart_png(kind, labels, values, colors=colors, figsize=figsize, currency_symbol=symbol)
            img = _load_image_rgb_from_bytes(png_bytes)
        except Exception:
            log.exception("Could not render product scorecard chart (%s)", kind)
            return None
        if not img:
            return None
        name = f"Chart{len(images) + 1}"
        images[name] = img
        return name, figsize[1] / figsize[0]

    def draw_chart(chart, width=content_w, extra_bottom=14):
        if not chart:
            return
        name, aspect = chart
        draw_h = width * aspect
        ensure_room(draw_h + extra_bottom)
        y = pdf_y(state["top"], draw_h)
        current["cmds"].append(f"q {width:.2f} 0 0 {draw_h:.2f} {margin:.2f} {y:.2f} cm /{name} Do Q")
        state["top"] += draw_h + extra_bottom

    # --- banner --------------------------------------------------------
    banner_h = 92
    fill_rect(0, 0, PAGE_W, banner_h, SIDEBAR)
    text(margin, 24, "PRODUCT SCORECARD", size=10, bold=True, color=LIGHT_ACCENT)
    text(margin, 44, product_name, size=18, bold=True, color=WHITE)
    generated = (generated_at or datetime.now().isoformat(timespec="seconds"))[:10]
    text(margin, 70, f"Generated {generated}", size=8, bold=False, color=LIGHT_ACCENT)
    state["top"] = banner_h + 22

    # --- headline stat cards --------------------------------------------
    lowest_entry = data.get("lowest_price_entry")
    lowest_str = money(data["lowest_price"], currency) if data.get("lowest_price") is not None else "-"
    cards = [
        ("Times ordered", str(data.get("po_count", 0))),
        ("Total spend", money(data.get("total_spend", 0), currency)),
        ("Avg price paid", money(data.get("avg_price", 0), currency)),
        ("Lowest price", lowest_str),
        ("Last ordered", (data.get("last_ordered_at") or "-")[:10]),
    ]
    gap = 10
    card_w = (content_w - gap * (len(cards) - 1)) / len(cards)
    card_h = 56
    ensure_room(card_h + 20)
    x = margin
    for label, value in cards:
        fill_rect(x, state["top"], card_w, card_h, CARD_BG)
        fill_rect(x, state["top"], card_w, 3, ACCENT)
        text(x + 8, state["top"] + 24, value[:16], size=12, bold=True, color=NAVY)
        text(x + 8, state["top"] + 42, label, size=7.5, bold=False, color=MUTED)
        x += card_w + gap
    state["top"] += card_h + 20
    # Tidy-up-pass fix (post-Batch-151, static-analysis finding): the
    # on-screen scorecard (po_generator_qt.py's ProductScorecardContent)
    # shows a "Lowest price paid: ... from <supplier> on <date>" caption
    # under these same stat cards using this exact field -- the PDF export
    # was computing lowest_entry above but never actually drawing it,
    # silently missing that one piece of context a reader of the PDF alone
    # (as opposed to someone looking at the live page) never got to see.
    if data.get("lowest_price") is not None and lowest_entry:
        ensure_room(16)
        caption = (
            f"Lowest price paid: {lowest_str} from "
            f"{lowest_entry.get('supplier_company_name') or '(unknown supplier)'} "
            f"on {(lowest_entry.get('date') or '')[:10]}."
        )
        text(margin, state["top"] + 10, caption, size=8.5, bold=False, color=MUTED)
        state["top"] += 20

    # --- monthly spend ----------------------------------------------------
    heading("Monthly spend")
    monthly = data.get("monthly", [])
    if monthly:
        chart = add_chart(
            "bar",
            [m["period"] for m in monthly],
            [m["total"] for m in monthly],
            colors=["#6498be"],
            money=True,
        )
        draw_chart(chart)
        row(["Month", "Orders", "Total"], [120, 100, 150], bold=True, color=MUTED)
        for m in monthly:
            row([m["period"], str(m["po_count"]), money(m["total"], currency)], [120, 100, 150])
    else:
        row(["No orders yet."], [400], color=MUTED)
    state["top"] += 10

    # --- spend by supplier -------------------------------------------------
    heading("Spend by supplier")
    by_supplier = data.get("by_supplier", [])
    if by_supplier:
        chart_suppliers = by_supplier[:8]
        chart = add_chart(
            "bar",
            [s["supplier"][:22] for s in chart_suppliers],
            [s["total_spend"] for s in chart_suppliers],
            colors=["#b079ad"],
            figsize=(7.4, 3.3),
            money=True,
        )
        draw_chart(chart)
        row(["Supplier", "Qty", "Times ordered", "Last price", "Total spend"], [150, 50, 90, 90, 100],
            bold=True, color=MUTED)
        for s in by_supplier[:20]:
            row(
                [s["supplier"], str(int(s["total_qty"])), str(s["times_ordered"]),
                 money(s["last_price"], currency) if s["last_price"] is not None else "-",
                 money(s["total_spend"], currency)],
                [150, 50, 90, 90, 100],
            )
    else:
        row(["No purchase history yet."], [400], color=MUTED)
    state["top"] += 10

    # --- price history over time -------------------------------------------
    heading("Price history")
    history = list(reversed(data.get("price_history", [])))  # oldest first for a left-to-right trend
    priced_history = [h for h in history if h.get("price") is not None]
    if priced_history:
        chart = add_chart(
            "line",
            [(h["date"] or "")[:10] for h in priced_history],
            [h["price"] for h in priced_history],
            colors=["#1e9e6b"],
            money=True,
        )
        draw_chart(chart)
        row(["Date", "Supplier", "Price", "Source"], [90, 180, 90, 90], bold=True, color=MUTED)
        for h in data.get("price_history", [])[:25]:
            row(
                [(h["date"] or "")[:10], h["supplier_company_name"] or "-",
                 money(h["price"], currency), h["source"]],
                [90, 180, 90, 90],
            )
    else:
        row(["No price history yet."], [400], color=MUTED)
    state["top"] += 10

    return [pg["cmds"] for pg in pages], images


def export_product_scorecard_to_pdf(path, product_name, data, currency="GBP"):
    pages, images = build_product_scorecard_pdf_pages(product_name, data, currency=currency)
    _write_pdf_from_command_pages(path, pages, page_size=(595, 842), images=images)


# ---- PDF attachments for the "content" emails (Phase 11) -------------------
#
# Yitzi: "How about attaching a PDF aswell to the email with the contence
# each time" -- confirmed scope (via an explicit choice, not a guess): the
# weekly/monthly/quarterly/yearly report emails, the stock recap, price
# request summaries, and the PO-to-supplier email itself each get a matching
# branded PDF built from the exact same data as the email; the plain
# outgoing price-ask emails (nothing to summarise yet) and the Zoho export
# reminder (already attaches its own Excel workbook) don't.
#
# The PO PDF originally reused the old save_simple_multipage_pdf/
# build_pdf_pages writer unchanged -- Yitzi later flagged that as "the PDF
# that gets attached to the PO is still the old layout" once every other
# email's PDF had this same banner/table treatment, so build_po_pdf_pages
# further below now covers it too, on the same canvas as the other three.
# save_simple_multipage_pdf/build_pdf_pages themselves are left in place
# (nothing currently calls them) rather than deleted, in case Yitzi ever
# wants the old plain layout back for some other purpose.

def _new_branded_pdf_canvas(kicker, title, subtitle="", conn=None):
    """A small closures-based PDF canvas, factoring out the banner + stat
    card + heading/table primitives build_supplier_scorecard_pdf_pages
    established, so the report-style PDF builders below don't each
    re-implement the same ~80 lines. Returns a dict of callables plus
    finish() -> (pages, images), the same shape _write_pdf_from_command_pages
    expects.

    conn: optional db connection, used to resolve the real RCG logo (or the
    user's own custom logo, or none at all) for the navy banner() the same
    way _effective_pdf_logo_path already does for the old PO PDF writer --
    "the PDF is missing the logo" was a real, confirmed gap in every one of
    these branded-canvas PDFs (they had text-only banners with no raster
    image support at all). Pass conn whenever it's available so Settings >
    PDF & Branding's logo choice/disable toggle is honoured; without one
    (conn=None) the bundled RCG logo is still used so a PDF built without a
    connection handy is still branded out of the box."""
    PAGE_W, PAGE_H = 595, 842
    margin = 36
    content_w = PAGE_W - margin * 2

    NAVY = hex_to_rgb01(RCG_INK)
    MUTED = hex_to_rgb01(RCG_MUTED)
    ACCENT = hex_to_rgb01(RCG_ACCENT_LIGHT)
    ACCENT_DARK = hex_to_rgb01(RCG_ACCENT)
    SIDEBAR = hex_to_rgb01(RCG_INK)
    BORDER = hex_to_rgb01(RCG_LINE)
    CARD_BG = hex_to_rgb01("#f4f6fa")
    WHITE = (1, 1, 1)
    LIGHT_ACCENT = hex_to_rgb01(RCG_ACCENT_LIGHT)
    MAUVE = hex_to_rgb01(RCG_MAUVE)
    GREEN = hex_to_rgb01(RCG_GREEN)
    RED_C = hex_to_rgb01(RCG_RED)

    pages = [{"cmds": []}]
    current = pages[-1]
    images = {}
    state = {"top": margin}

    # Resolve and load the banner logo, same source of truth as the old PO
    # PDF writer (_effective_pdf_logo_path): whatever's chosen in Settings >
    # PDF & Branding, the bundled RCG logo if nothing's chosen and it hasn't
    # been switched off, or nothing at all if it has. Uses _load_image_rgba
    # (not _load_image_rgb) so the logo's real transparency survives into
    # the PDF via a genuine /SMask -- Yitzi: "the logo should not have white
    # background" (batch 92). Before this fix _load_image_rgb flattened
    # transparency onto solid white, which is why every banner showed the
    # logo sitting on a white card instead of directly on the banner fill.
    #
    # This banner's fill is always one of the brand's own colors (never
    # white/light) -- see SIDEBAR above -- so the logo is whitened via
    # _whiten_rgba_image before being drawn, same as the app's own navy
    # sidebar already does with its pre-made white rcg_icon_only.png.
    # Without this, the RCG wordmark's own dark ink text nearly vanishes
    # into a similarly dark banner (batch 93).
    logo_img = None
    try:
        if conn is not None:
            logo_path = _effective_pdf_logo_path(conn)
        else:
            bundled = _rcg_asset_path("assets", "rcg_logo_full_transparent.png")
            logo_path = bundled if os.path.exists(bundled) else None
        if logo_path:
            logo_img = _load_image_rgba(logo_path, max_dim=900)
            if logo_img:
                logo_img = _whiten_rgba_image(logo_img)
    except Exception:
        log.exception("Could not load the PDF banner logo")
        logo_img = None
    if logo_img:
        images["BrandLogo"] = logo_img

    def new_page():
        nonlocal current
        current = {"cmds": []}
        pages.append(current)
        state["top"] = margin

    def ensure_room(h):
        if state["top"] + h > PAGE_H - margin:
            new_page()

    def pdf_y(top, height=0):
        return PAGE_H - top - height

    def text(x, top, s, size=10, bold=False, color=NAVY):
        r, g, b = color
        font = "/F2" if bold else "/F1"
        y = PAGE_H - top - size
        current["cmds"].append(
            f"BT {r:.3f} {g:.3f} {b:.3f} rg {font} {size} Tf 1 0 0 1 {x} {y} Tm ({pdf_escape(s)}) Tj ET"
        )

    def fill_rect(x, top, w, h, color):
        r, g, b = color
        y = pdf_y(top, h)
        current["cmds"].append(f"q {r:.3f} {g:.3f} {b:.3f} rg {x:.2f} {y:.2f} {w:.2f} {h:.2f} re f Q")

    def hline(color=BORDER, w=0.75):
        r, g, b = color
        y = PAGE_H - state["top"]
        current["cmds"].append(f"q {w} w {r:.3f} {g:.3f} {b:.3f} RG {margin} {y} m {PAGE_W - margin} {y} l S Q")

    def heading(s):
        ensure_room(32)
        text(margin, state["top"], s, size=13, bold=True, color=ACCENT_DARK)
        state["top"] += 16
        hline(color=ACCENT, w=1.5)
        state["top"] += 12

    def row(cells, widths, size=9, bold=False, color=NAVY):
        ensure_room(16)
        x = margin
        for cell, w in zip(cells, widths):
            text(x, state["top"], str(cell)[:70], size=size, bold=bold, color=color)
            x += w
        state["top"] += 15

    def wrapped_row(cells, widths, size=9, bold=False, color=NAVY, wrap_col=None, leading=12):
        """Like row(), but one column (wrap_col, by index) may run long and
        wraps onto extra lines rather than being hard-truncated -- used for
        the price request summary's supplier notes, which can be a full
        sentence."""
        if wrap_col is None:
            row(cells, widths, size=size, bold=bold, color=color)
            return
        limit = max(8, int(widths[wrap_col] / (size * 0.52)))
        wrap_text = str(cells[wrap_col])
        words = wrap_text.split()
        lines = []
        cur = ""
        for w in words:
            test = (cur + " " + w).strip()
            if len(test) <= limit:
                cur = test
            else:
                if cur:
                    lines.append(cur)
                cur = w
        if cur:
            lines.append(cur)
        lines = lines or [""]
        ensure_room(leading * len(lines) + 2)
        x = margin
        top0 = state["top"]
        for i, (cell, w) in enumerate(zip(cells, widths)):
            if i == wrap_col:
                for li, ln in enumerate(lines):
                    text(x, top0 + li * leading, ln, size=size, bold=bold, color=color)
            else:
                text(x, top0, str(cell)[:70], size=size, bold=bold, color=color)
            x += w
        state["top"] = top0 + leading * len(lines)

    def table_header(headers, widths, height=22):
        """A shaded table header band -- solid navy fill with white
        uppercase labels -- matching email_table_html's header row, so a
        PDF table reads as the same design as its emailed HTML counterpart
        rather than the plain unshaded label row row()/heading() alone
        produce."""
        ensure_room(height + 4)
        fill_rect(margin, state["top"], content_w, height, NAVY)
        x = margin
        # text()'s "top" argument is the glyph's TOP edge -- it derives the
        # baseline internally as top + size (see text() above). This band's
        # baseline is meant to sit 7pt above the row's own bottom edge, so
        # the font size has to be subtracted back out here, or the baseline
        # (and the descenders hanging off it) lands 'size' points too low,
        # spilling into whatever gets painted below (Batch 92 fix -- see
        # FEATURE_LOG.md for the "PDF tables render as illegible overlapping
        # blocks" bug this caused).
        for h, w in zip(headers, widths):
            text(x + 8, state["top"] + height - 7 - 8, str(h).upper(), size=8, bold=True, color=WHITE)
            x += w
        state["top"] += height

    def table_row(cells, widths, index=0, color=None, bold=False, bg=None, height=20, badge=None):
        """One shaded, alternating-background table body row -- the PDF
        equivalent of email_table_html's white/pale-blue striping. bg
        overrides the alternating shade entirely (e.g. a green highlight
        for a "best price" row); badge, if given, is (text, bg_color,
        text_color) drawn as a small filled chip after the last cell --
        the PDF equivalent of the email's green "BEST PRICE" tag."""
        ensure_room(height)
        row_bg = bg if bg is not None else (WHITE if index % 2 == 0 else CARD_BG)
        fill_rect(margin, state["top"], content_w, height, row_bg)
        x = margin
        last_x = margin
        # Same fix as table_header just above: text()'s "top" is the glyph's
        # TOP edge, not its baseline, so the font size (9 here) has to be
        # subtracted back out to keep the baseline 7pt above the row's own
        # bottom edge instead of 'size' points below it.
        for cell, w in zip(cells, widths):
            text(x + 8, state["top"] + height - 7 - 9, str(cell)[:70], size=9, bold=bold, color=color or NAVY)
            last_x = x
            x += w
        if badge:
            badge_text, badge_bg, badge_color = badge
            badge_w = 12 + len(badge_text) * 4.6
            badge_x = min(last_x + 8, PAGE_W - margin - badge_w)
            fill_rect(badge_x, state["top"] + 3, badge_w, height - 6, badge_bg)
            text(badge_x + 6, state["top"] + height - 7 - 7, badge_text, size=7, bold=True, color=badge_color)
        state["top"] += height

    def add_chart(kind, labels, values, colors=None, figsize=(7.4, 2.7), money_fmt=False, currency="GBP", title=""):
        try:
            symbol = CURRENCY_SYMBOLS.get(currency, "") if money_fmt else ""
            png_bytes = _render_chart_png(kind, labels, values, colors=colors, figsize=figsize,
                                           currency_symbol=symbol, title=title)
            img = _load_image_rgb_from_bytes(png_bytes)
        except Exception:
            log.exception("Could not render PDF chart (%s)", kind)
            return None
        if not img:
            return None
        name = f"Chart{len(images) + 1}"
        images[name] = img
        return name, figsize[1] / figsize[0]

    def draw_chart(chart, width=content_w, extra_bottom=14):
        if not chart:
            return
        name, aspect = chart
        draw_h = width * aspect
        ensure_room(draw_h + extra_bottom)
        y = pdf_y(state["top"], draw_h)
        current["cmds"].append(f"q {width:.2f} 0 0 {draw_h:.2f} {margin:.2f} {y:.2f} cm /{name} Do Q")
        state["top"] += draw_h + extra_bottom

    def stat_cards(cards):
        """cards: list of (label, value) pairs, laid out as an even-width
        row of tinted stat tiles -- the PDF equivalent of email_stat_strip_html.
        The top accent strip alternates light-blue/mauve across the tiles,
        matching email_stat_strip_html's own alternating tint exactly (that
        function is the one place the brand's mauve was already used
        outside of charts) -- batch 94, Yitzi: "do you think we should add
        some purple into the pdf" -> yes, in the same spot the emails
        already do it, rather than inventing a new use for it."""
        if not cards:
            return
        gap = 10
        card_w = (content_w - gap * (len(cards) - 1)) / len(cards)
        card_h = 56
        ensure_room(card_h + 20)
        x = margin
        for i, (label, value) in enumerate(cards):
            strip_color = ACCENT if i % 2 == 0 else MAUVE
            fill_rect(x, state["top"], card_w, card_h, CARD_BG)
            fill_rect(x, state["top"], card_w, 3, strip_color)
            text(x + 8, state["top"] + 24, str(value)[:18], size=12, bold=True, color=NAVY)
            text(x + 8, state["top"] + 42, label, size=7.5, bold=False, color=MUTED)
            x += card_w + gap
        state["top"] += card_h + 20

    def banner(height=92):
        fill_rect(0, 0, PAGE_W, height, SIDEBAR)
        text(margin, 24, (kicker or "").upper(), size=10, bold=True, color=LIGHT_ACCENT)
        text(margin, 44, title[:70], size=19 if len(title) > 46 else 21, bold=True, color=WHITE)
        if subtitle:
            text(margin, 70, subtitle, size=9, bold=False, color=LIGHT_ACCENT)
        if logo_img:
            img_w, img_h = logo_img[0], logo_img[1]
            draw_w = min(108, content_w * 0.3)
            draw_h = draw_w * img_h / img_w
            if draw_h > height - 24:
                draw_h = height - 24
                draw_w = draw_h * img_w / img_h
            x = PAGE_W - margin - draw_w
            logo_top = (height - draw_h) / 2
            y = pdf_y(logo_top, draw_h)
            current["cmds"].append(f"q {draw_w:.2f} 0 0 {draw_h:.2f} {x:.2f} {y:.2f} cm /BrandLogo Do Q")
        state["top"] = height + 22

    def finish():
        return [pg["cmds"] for pg in pages], images

    return {
        "margin": margin, "content_w": content_w, "page_w": PAGE_W, "page_h": PAGE_H, "state": state,
        "text": text, "fill_rect": fill_rect, "hline": hline, "heading": heading,
        "row": row, "wrapped_row": wrapped_row, "ensure_room": ensure_room,
        "table_header": table_header, "table_row": table_row,
        "add_chart": add_chart, "draw_chart": draw_chart, "stat_cards": stat_cards,
        "banner": banner, "finish": finish,
        "colors": {
            "navy": NAVY, "muted": MUTED, "accent": ACCENT, "accent_dark": ACCENT_DARK,
            "border": BORDER, "card_bg": CARD_BG, "white": WHITE, "light_accent": LIGHT_ACCENT,
            "green": GREEN, "red": RED_C, "mauve": MAUVE,
        },
    }


def build_period_report_pdf_pages(conn, period_kind, date_from, date_to, label=None, currency=None):
    """A branded PDF matching build_period_report_email's content exactly --
    recomputed from the same underlying queries (list_pos_in_range,
    _top_items_in_range, compare_periods) rather than threaded through the
    email dict, since the weekly and monthly/quarterly/yearly builders
    return different-shaped dicts (weekly's has the full per-order list;
    the others don't) and re-running these same cheap read-only queries
    keeps this builder simple and independent of that. Weekly gets the full
    per-order breakdown (matching build_po_period_email); monthly/quarterly/
    yearly get the lighter overview (matching build_period_summary_email) --
    same split the email itself makes."""
    currency = currency or get_setting(conn, "currency", "GBP")
    noun = period_kind or "period"
    label = label or f"This {noun}'s"
    comparison = compare_periods(conn, period_kind, date_from, date_to)
    current = comparison["current"]

    c = _new_branded_pdf_canvas(
        "Rose Communications Group",
        f"{label} PO Summary",
        f"{date_from.strftime('%d %b %Y')} to {date_to.strftime('%d %b %Y')}",
        conn=conn,
    )
    c["banner"]()
    c["stat_cards"]([
        (plural_word(current["po_count"], "order") + " placed", str(current["po_count"])),
        ("Total spend", money(current["total_spend"], currency)),
        ("Average PO value", money(current["avg_po_value"], currency)),
        (plural_word(current["total_qty"], "unit") + " ordered", f'{current["total_qty"]:g}'),
        ("Savings recorded", money(current["savings_total"], currency)),
    ])

    # -- comparison to the previous period --
    c["heading"](f"How this compares to the previous {noun}")
    c["table_header"](["Metric", f"This {noun}", f"Last {noun}", "Change"], [180, 120, 120, 100])
    for i, key in enumerate(_COMPARISON_METRICS):
        d = comparison["deltas"][key]
        m_label = _METRIC_LABELS.get(key, key)
        cur_txt = _format_metric_value(key, d["current"], currency)
        prev_txt = _format_metric_value(key, d["previous"], currency)
        pct = d["pct_change"]
        if pct is None:
            change_txt = "n/a"
        else:
            arrow = "Up" if pct > 0.05 else ("Down" if pct < -0.05 else "Flat")
            change_txt = f"{arrow} {abs(pct):.0f}%" if abs(pct) > 0.05 else f"{arrow}"
        c["table_row"]([m_label, cur_txt, prev_txt, change_txt], [180, 120, 120, 100], index=i)
    c["state"]["top"] += 10

    if period_kind == "week":
        pos = list_pos_in_range(conn, date_from, date_to)
        by_supplier = {}
        for p in pos:
            name = p["supplier_company_name"] or "(no supplier)"
            entry = by_supplier.setdefault(name, {"count": 0, "total": 0.0})
            entry["count"] += 1
            entry["total"] += float(p["base_total"] or 0)
        top_suppliers = sorted(by_supplier.items(), key=lambda kv: kv[1]["total"], reverse=True)
        top_items_cost, top_items_qty = _top_items_in_range(conn, pos)
        issues = list_supplier_issues_in_range(conn, date_from, date_to)

        c["heading"]("Main suppliers")
        if top_suppliers:
            top5 = top_suppliers[:5]
            other_total = sum(info["total"] for _n, info in top_suppliers[5:])
            pie_labels = [n for n, _i in top5] + (["Other"] if other_total > 0 else [])
            pie_values = [info["total"] for _n, info in top5] + ([other_total] if other_total > 0 else [])
            chart = c["add_chart"]("pie", pie_labels, pie_values, figsize=(3.6, 3.0))
            c["draw_chart"](chart, width=200)
            c["table_header"](["Supplier", "Orders", "Total"], [220, 140, 140])
            for i, (name, info) in enumerate(top_suppliers):
                c["table_row"]([name, pluralize(info["count"], "order"), money(info["total"], currency)],
                                [220, 140, 140], index=i)
        else:
            c["row"](["No orders placed in this period."], [400], color=c["colors"]["muted"])
        c["state"]["top"] += 10

        c["heading"]("Top items by quantity")
        if top_items_qty:
            bar_labels = [(i["product"][:18] + "...") if len(i["product"]) > 18 else i["product"] for i in top_items_qty]
            bar_values = [i["total_qty"] for i in top_items_qty]
            chart = c["add_chart"]("bar", bar_labels, bar_values, colors=[RCG_MAUVE], figsize=(7.4, 2.7))
            c["draw_chart"](chart)
            c["table_header"](["Product", "Total cost"], [340, 140])
            for i, item in enumerate(top_items_cost):
                c["table_row"]([item["product"], money(item["total_cost"], currency)], [340, 140], index=i)
        else:
            c["row"](["No items ordered in this period."], [400], color=c["colors"]["muted"])
        c["state"]["top"] += 10

        c["heading"](f"Orders placed ({pluralize(len(pos), 'order')})")
        if pos:
            c["table_header"](["PO Ref", "Supplier", "Total"], [180, 220, 100])
            for i, p in enumerate(pos):
                c["table_row"]([p["po_ref"], p["supplier_company_name"] or "", money(p["total"], p["currency"] or currency)],
                                [180, 220, 100], index=i)
        else:
            c["row"](["No orders placed in this period."], [400], color=c["colors"]["muted"])
        c["state"]["top"] += 10

        c["heading"](f"Supplier issues logged ({pluralize(len(issues), 'issue')})")
        if issues:
            c["table_header"](["Date", "Supplier", "Category", "Subject"], [70, 140, 90, 220])
            for i in issues:
                c["wrapped_row"](
                    [(i.get("at") or "")[:10], i["supplier_company_name"], i.get("category") or "-",
                     i.get("subject") or i.get("note", "")],
                    [70, 140, 90, 220], wrap_col=3,
                )
        else:
            c["row"](["No supplier issues logged this period."], [400], color=c["colors"]["muted"])
    else:
        trend_granularity = {"month": "week", "quarter": "month", "year": "month"}.get(period_kind, "month")
        trend = report_spend_over_time(
            conn, date_from=date_from.strftime("%Y-%m-%d"), date_to=date_to.strftime("%Y-%m-%d"),
            granularity=trend_granularity,
        )
        c["heading"](f"Spend trend this {noun}")
        if trend:
            chart = c["add_chart"](
                "line", [t["period"] for t in trend], [t["total"] for t in trend],
                colors=[RCG_ACCENT], figsize=(7.4, 2.7), money_fmt=True, currency=currency,
            )
            c["draw_chart"](chart)
        else:
            c["row"](["No spend recorded in this period."], [400], color=c["colors"]["muted"])
        c["state"]["top"] += 6

        c["heading"]("Top suppliers")
        if current["top_suppliers"]:
            c["table_header"](["Supplier", "Orders", "Total"], [220, 140, 140])
            for i, s in enumerate(current["top_suppliers"]):
                c["table_row"]([s["supplier"] or "(no supplier)", pluralize(s["po_count"], "order"), money(s["total"], currency)],
                                [220, 140, 140], index=i)
        else:
            c["row"](["None"], [400], color=c["colors"]["muted"])
        c["state"]["top"] += 10

        c["heading"]("Top products")
        if current["top_products"]:
            c["table_header"](["Product", "Quantity", "Total spend"], [220, 140, 140])
            for i, p in enumerate(current["top_products"]):
                c["table_row"]([p["product"], f'{p["total_qty"]:g} {plural_word(p["total_qty"], "unit")}',
                                 money(p["total_spend"], currency)], [220, 140, 140], index=i)
        else:
            c["row"](["None"], [400], color=c["colors"]["muted"])

    return c["finish"]()


def export_period_report_to_pdf(path, conn, period_kind, date_from, date_to, label=None, currency=None):
    pages, images = build_period_report_pdf_pages(conn, period_kind, date_from, date_to, label=label, currency=currency)
    _write_pdf_from_command_pages(path, pages, page_size=(595, 842), images=images)


def build_stock_recap_pdf_pages(conn, date_from, date_to, label="Last week's", currency=None, recap_enrichment=None):
    """A branded PDF matching build_stock_recap_email's content exactly --
    every order in the period with its line items, recomputed from the same
    list_pos_in_range + po_items query the email builder uses. recap_enrichment
    is the same optional Batch 98 dict build_stock_recap_email takes (None
    when stock sync is off, leaving this exactly as it always was)."""
    currency = currency or get_setting(conn, "currency", "GBP")
    pos = list_pos_in_range(conn, date_from, date_to)
    items_by_po = {}
    if pos:
        placeholders = ",".join("?" * len(pos))
        for r in conn.execute(
            f"SELECT * FROM po_items WHERE po_id IN ({placeholders}) ORDER BY po_id, position",
            [p["id"] for p in pos],
        ):
            items_by_po.setdefault(r["po_id"], []).append(dict(r))

    c = _new_branded_pdf_canvas(
        "Rose Communications Group",
        f"{label} Stock Recap",
        f"{date_from.strftime('%d %b %Y')} to {date_to.strftime('%d %b %Y')}",
        conn=conn,
    )
    c["banner"]()
    c["stat_cards"]([(plural_word(len(pos), "order") + " placed", str(len(pos)))])

    def render_po_block(p, items):
        po_currency = p.get("currency") or currency
        c["ensure_room"](40)
        heading = f'{p["po_ref"]} - {p.get("supplier_company_name") or "(no supplier)"}'
        status_line = _stock_recap_status_line(p["id"], recap_enrichment)
        c["heading"](heading)
        if status_line:
            c["row"]([f"Status: {status_line[0]}"], [400], color=c["colors"][status_line[1]])
        c["table_header"](["Qty", "Item", "Unit price", "Line total"], [50, 260, 100, 100])
        for i, item in enumerate(items):
            line_total = float(item.get("qty") or 0) * float(item.get("price") or 0)
            c["table_row"](
                [str(item.get("qty", "")), display_product_label(item), money(item.get("price"), po_currency), money(line_total, po_currency)],
                [50, 260, 100, 100], index=i,
            )
        c["state"]["top"] += 4
        c["row"]([f'Order total: {money(p.get("total"), po_currency)}'], [400], bold=True)
        c["state"]["top"] += 8

    if not pos:
        c["row"](["No orders were placed in this period."], [400], color=c["colors"]["muted"])
    for p in pos:
        render_po_block(p, items_by_po.get(p["id"], []))

    outstanding_pos = (recap_enrichment or {}).get("outstanding_previous_pos", [])
    if outstanding_pos:
        c["state"]["top"] += 6
        c["ensure_room"](30)
        c["heading"]("Still outstanding from before this week")
        for p in outstanding_pos:
            render_po_block(p, p.get("items", []))

    return c["finish"]()


def export_stock_recap_to_pdf(path, conn, date_from, date_to, label="Last week's", currency=None, recap_enrichment=None):
    pages, images = build_stock_recap_pdf_pages(
        conn, date_from, date_to, label=label, currency=currency, recap_enrichment=recap_enrichment
    )
    _write_pdf_from_command_pages(path, pages, page_size=(595, 842), images=images)


def build_price_request_summary_pdf_pages(conn, request_id, currency=None):
    """A branded PDF matching build_price_request_summary's content exactly
    -- re-fetches the price request (get_price_request) rather than taking
    the already-built summary dict, since that dict doesn't carry the raw
    per-supplier replies list needed for the table (only the single
    cheapest one)."""
    req = get_price_request(conn, request_id)
    if not req:
        return [], {}
    currency = currency or get_setting(conn, "currency", "GBP")
    replies = req["replies"]
    status_labels = {"pending": "Pending", "received": "Received", "declined": "Unavailable"}
    received = [
        r for r in replies
        if r["status"] == "received" and r.get("price") is not None and float(r["price"]) > 0
    ]
    cheapest_price = min((float(r["price"]) for r in received), default=None)
    price_context = get_last_and_lowest_price_paid(conn, req["product_name"])

    c = _new_branded_pdf_canvas(
        "Rose Communications Group",
        "Pricing Quote",
        f'{req["product_name"]} (qty {req["qty"]})',
        conn=conn,
    )
    c["banner"]()

    # Matches the emailed summary's own table exactly: navy header band,
    # alternating body-row shading, a green highlight + "BEST PRICE" chip on
    # the cheapest received reply, red text for declined suppliers -- "the
    # format is not nice and uniform like email" was a real gap (the old
    # version here was a plain unshaded row() list with no colour on
    # anything but the text itself).
    c["heading"]("Pricing received")
    c["table_header"](["Supplier", "Status", "Price", "Date"], [220, 100, 100, 90])
    for i, r in enumerate(replies):
        label = status_labels.get(r["status"], (r["status"] or "").capitalize())
        price_txt = (
            money(r["price"], currency) if r.get("price") is not None
            else ("Not available" if r["status"] == "declined" else "-")
        )
        is_cheapest = (
            cheapest_price is not None and r["status"] == "received"
            and r.get("price") is not None and float(r["price"]) == cheapest_price
        )
        if is_cheapest:
            bg, color = hex_to_rgb01("#e3f7e8"), c["colors"]["green"]
        elif r["status"] == "declined":
            bg, color = hex_to_rgb01("#fbeaea"), c["colors"]["red"]
        else:
            bg, color = None, c["colors"]["navy"]
        badge = ("BEST PRICE", c["colors"]["green"], c["colors"]["white"]) if is_cheapest else None
        c["table_row"]([r["supplier_company_name"] or "(no supplier)", label, price_txt, (r.get("replied_at") or "")[:10] or "-"],
                        [220, 100, 100, 90], index=i, color=color, bold=is_cheapest, bg=bg, badge=badge)
    c["state"]["top"] += 14

    c["heading"]("Price history")
    history_cards = []
    if price_context["last"]:
        lp = price_context["last"]
        history_cards.append((
            "Last price paid",
            f'{money(lp["price"], currency)} — {lp["supplier_company_name"] or "(no supplier)"} ({(lp["date"] or "")[:10]})',
        ))
    if price_context["lowest"]:
        lo = price_context["lowest"]
        history_cards.append((
            "Lowest price ever paid",
            f'{money(lo["price"], currency)} — {lo["supplier_company_name"] or "(no supplier)"} ({(lo["date"] or "")[:10]})',
        ))
    if history_cards:
        for i, (label, value) in enumerate(history_cards):
            c["table_row"]([label, value], [160, 360], index=i, bold=(i == 1 and len(history_cards) == 2))
    else:
        c["row"](["No price history on file for this product yet."], [400], color=c["colors"]["muted"])

    return c["finish"]()


def export_price_request_summary_to_pdf(path, conn, request_id, currency=None):
    pages, images = build_price_request_summary_pdf_pages(conn, request_id, currency=currency)
    if not pages:
        raise RuntimeError("This price request no longer exists.")
    _write_pdf_from_command_pages(path, pages, page_size=(595, 842), images=images)


def build_combined_price_request_summary_pdf_pages(conn, request_ids, currency=None):
    """Batch 162 -- the PDF half of build_combined_price_request_summary:
    one branded PDF covering several price requests (e.g. an iPhone quote
    and an iPad quote together), rather than a separate PDF per product.
    Reuses the same _new_branded_pdf_canvas + heading()/table_header()/
    table_row() primitives build_price_request_summary_pdf_pages already
    uses for a single request -- the only real difference is that heading()
    is called once per product (it already calls ensure_room() itself, so
    a product's table safely starts a fresh page if it wouldn't otherwise
    fit) instead of the two fixed "Pricing received"/"Price history"
    headings a single-request PDF always has.

    Duplicate ids collapse to one copy (first-occurrence order kept); a
    request that's since been deleted is silently skipped. Returns
    ([], {}) if none of the given ids resolve to a request still on file,
    matching build_price_request_summary_pdf_pages's own "request no
    longer exists" signal."""
    seen_ids = []
    for rid in request_ids:
        if rid not in seen_ids:
            seen_ids.append(rid)
    reqs = [get_price_request(conn, rid) for rid in seen_ids]
    reqs = [r for r in reqs if r]
    if not reqs:
        return [], {}

    currency = currency or get_setting(conn, "currency", "GBP")
    status_labels = {"pending": "Pending", "received": "Received", "declined": "Unavailable"}
    names_for_subtitle = ", ".join(r["product_name"] for r in reqs)
    if len(names_for_subtitle) > 90:
        names_for_subtitle = f"{pluralize(len(reqs), 'product')}"

    c = _new_branded_pdf_canvas(
        "Rose Communications Group", "Pricing Quotes", names_for_subtitle, conn=conn,
    )
    c["banner"]()

    for req in reqs:
        replies = req["replies"]
        received = [
            r for r in replies
            if r["status"] == "received" and r.get("price") is not None and float(r["price"]) > 0
        ]
        cheapest_price = min((float(r["price"]) for r in received), default=None)
        price_context = get_last_and_lowest_price_paid(conn, req["product_name"])

        c["heading"](f'{req["product_name"]} (qty {req["qty"]})')
        c["table_header"](["Supplier", "Status", "Price", "Date"], [220, 100, 100, 90])
        for i, r in enumerate(replies):
            label = status_labels.get(r["status"], (r["status"] or "").capitalize())
            price_txt = (
                money(r["price"], currency) if r.get("price") is not None
                else ("Not available" if r["status"] == "declined" else "-")
            )
            is_cheapest = (
                cheapest_price is not None and r["status"] == "received"
                and r.get("price") is not None and float(r["price"]) == cheapest_price
            )
            if is_cheapest:
                bg, color = hex_to_rgb01("#e3f7e8"), c["colors"]["green"]
            elif r["status"] == "declined":
                bg, color = hex_to_rgb01("#fbeaea"), c["colors"]["red"]
            else:
                bg, color = None, c["colors"]["navy"]
            badge = ("BEST PRICE", c["colors"]["green"], c["colors"]["white"]) if is_cheapest else None
            c["table_row"]([r["supplier_company_name"] or "(no supplier)", label, price_txt, (r.get("replied_at") or "")[:10] or "-"],
                            [220, 100, 100, 90], index=i, color=color, bold=is_cheapest, bg=bg, badge=badge)
        c["state"]["top"] += 6

        history_cards = []
        if price_context["last"]:
            lp = price_context["last"]
            history_cards.append((
                "Last price paid",
                f'{money(lp["price"], currency)} — {lp["supplier_company_name"] or "(no supplier)"} ({(lp["date"] or "")[:10]})',
            ))
        if price_context["lowest"]:
            lo = price_context["lowest"]
            history_cards.append((
                "Lowest price ever paid",
                f'{money(lo["price"], currency)} — {lo["supplier_company_name"] or "(no supplier)"} ({(lo["date"] or "")[:10]})',
            ))
        for i, (label, value) in enumerate(history_cards):
            c["table_row"]([label, value], [160, 360], index=i, bold=(i == 1 and len(history_cards) == 2))
        c["state"]["top"] += 14

    return c["finish"]()


def export_combined_price_request_summary_to_pdf(path, conn, request_ids, currency=None):
    pages, images = build_combined_price_request_summary_pdf_pages(conn, request_ids, currency=currency)
    if not pages:
        raise RuntimeError("None of the selected price requests still exist.")
    _write_pdf_from_command_pages(path, pages, page_size=(595, 842), images=images)


def build_po_pdf_pages(po, conn=None, terms_text=None):
    """A branded PDF matching build_html's own PO email content and visual
    language exactly -- goods table, PO ref/currency/business details, total
    order value, delivery/invoice addresses -- built on the same
    _new_branded_pdf_canvas + table_header/table_row primitives as the
    period report / stock recap / price request summary PDFs. Replaces the
    old plain bordered-box build_pdf_pages layout: Yitzi flagged "the PDF
    that gets attached to the PO is still the old layout" once every other
    email's matching PDF had already been redesigned, so this brings the PO
    PDF's own look in line with the rest.

    Known limitation shared with the other three canvas-based PDFs: a goods
    table that runs long enough to spill onto a second page does not repeat
    its header row there (the canvas's ensure_room()/new_page() has no
    "redraw this on every page" hook) -- acceptable for now since it matches
    the existing behaviour everywhere else this canvas is used, but worth
    revisiting if a PO with a very long line-item list turns out to need
    it."""
    c = _new_branded_pdf_canvas(
        "Rose Communications Group", "Purchase Order", po.get("po_ref", ""), conn=conn,
    )
    c["banner"]()

    c["stat_cards"]([
        ("PO Ref", po.get("po_ref", "")),
        ("Currency", po.get("currency", "")),
        ("Total Order Value", money(po.get("total"), po.get("currency"))),
    ])

    items = po.get("items") or []
    c["heading"](f'Goods Ordered ({pluralize(len(items), "item")})')
    if items:
        c["table_header"](["Qty", "Product", "Price"], [55, 345, 100])
        for i, item in enumerate(items):
            c["table_row"](
                [f'{item.get("qty", "")}X', display_product_label(item), money(item.get("price"), po.get("currency"))],
                [55, 345, 100], index=i,
            )
    else:
        c["row"](["No items on this order."], [400], color=c["colors"]["muted"])
    c["state"]["top"] += 6
    c["table_row"]([f'Total Order Value: {money(po.get("total"), po.get("currency"))}'], [500],
                    bold=True, bg=c["colors"]["card_bg"])
    c["state"]["top"] += 14

    c["heading"]("Business details")
    c["table_row"](["Business name", po.get("business_name", "")], [180, 320], index=0)
    c["table_row"](["Company registration number", po.get("company_number", "")], [180, 320], index=1)
    c["state"]["top"] += 10

    def address_block(title, name, address):
        c["ensure_room"](60)
        c["heading"](title)
        lines = [str(name or "").strip()] + [ln.strip() for ln in str(address or "").splitlines() if ln.strip()]
        for i, ln in enumerate(lines):
            c["text"](c["margin"], c["state"]["top"], ln, size=10, bold=(i == 0), color=c["colors"]["navy"])
            c["state"]["top"] += 14
        c["state"]["top"] += 6

    address_block("Delivery Address", po.get("delivery_name"), po.get("delivery_address"))
    address_block("Invoice Address", po.get("invoice_name"), po.get("invoice_address"))

    if terms_text and str(terms_text).strip():
        c["heading"]("Terms")
        c["wrapped_row"]([str(terms_text).strip()], [500], wrap_col=0, size=9, color=c["colors"]["muted"], leading=13)

    return c["finish"]()


def export_po_pdf(path, po, conn=None, terms_text=None):
    if terms_text is None and conn is not None:
        terms_text = get_setting(conn, "pdf_terms_text", "") or None
    pages, images = build_po_pdf_pages(po, conn=conn, terms_text=terms_text)
    if not pages:
        raise RuntimeError("No PDF pages were generated")
    _write_pdf_from_command_pages(path, pages, page_size=(595, 842), images=images)


def export_report_to_xlsx(path, title, columns, rows):
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = (title or "Report")[:31] or "Report"
    headers = [c[0] for c in columns]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in rows:
        values = []
        for header, key, fmt in columns:
            raw = row.get(key, "")
            values.append(fmt(raw, row) if fmt else raw)
        ws.append(values)
    for i, header in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(i)].width = max(12, len(str(header)) + 4)
    wb.save(path)


# ============================================================
# 7. Backup / restore / export / import
#
# v1 stored everything in two loose JSON files with no backup at all — a
# corrupted write could silently lose every PO ever raised. v2 adds two
# independent safety nets:
#   * backup_now(): a full-restore-capable .zip bundle (the SQLite database
#     plus any external file a restore would otherwise depend on, e.g. a
#     custom PDF logo), taken automatically on every launch, on the
#     twice-daily weekday schedule (see ensure_backup_tasks_registered),
#     and before risky operations like a Zoho import or a catalog reset.
#     Pruned by age, never by count, so nothing inside the retention window
#     is ever at risk of being crowded out.
#   * export_all_json() / import_all_json(): a portable, human-readable
#     full export you can move to another machine or archive off-site.
#
# Everything else Yitzi's data covers -- settings, addresses, suppliers,
# products, price requests, savings, reports config -- already lives in the
# one SQLite database, so the .db file itself is the actual full backup of
# all of that. The only other thing a restore could otherwise be missing is
# an external file the app merely references by path (currently just a
# custom PDF logo override), so that's what gets bundled alongside it.
# Generated export artifacts (Zoho CSV/XLSX files already written out to
# disk) are deliberately not bundled -- they're regenerable from the data
# that IS backed up, not a source of truth themselves.
# ============================================================

BACKUP_MANIFEST_NAME = "manifest.json"
BACKUP_DB_NAME = "data.db"


def _backup_extra_files(conn):
    """Files outside the database that a full restore would need. Currently
    just a custom PDF logo override, if one is set and the file still
    exists. Returns a list of (setting_key, source_path) pairs."""
    extras = []
    logo_path = get_setting(conn, "pdf_logo_path", "").strip()
    if logo_path and Path(logo_path).is_file():
        extras.append(("pdf_logo_path", Path(logo_path)))
    return extras


def backup_now(reason="launch"):
    """Creates one full-restore-capable backup: a .zip bundle containing a
    point-in-time copy of the database, a small manifest recording when/why
    it was taken and which extra files (if any) it carries, and copies of
    any such extra files (e.g. a custom logo) so a restore never depends on
    something outside the app still being where it used to be. Prunes both
    the local backups folder and (if enabled) the secondary folder by age
    afterwards, and never raises -- callers treat a None return as "backup
    failed, see the log"."""
    conn = get_connection()
    conn.commit()
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    dest = BACKUP_DIR / f"data-{stamp}-{reason}.zip"
    if dest.exists():
        # Two backups landing in the same wall-clock second (e.g. the
        # launch backup and a migration's safety backup, both taken within
        # the same second on a fast machine) would otherwise collide on an
        # identical filename and silently overwrite one another -- append
        # a short disambiguator instead of ever losing one.
        n = 2
        while (BACKUP_DIR / f"data-{stamp}-{reason}-{n}.zip").exists():
            n += 1
        dest = BACKUP_DIR / f"data-{stamp}-{reason}-{n}.zip"
    try:
        extras = _backup_extra_files(conn)
        manifest = {
            "app_version": APP_VERSION,
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "reason": reason,
            "extra_files": [
                {"setting_key": key, "archive_name": f"extra/{src.name}", "original_name": src.name}
                for key, src in extras
            ],
        }
        with zipfile.ZipFile(str(dest), "w", zipfile.ZIP_DEFLATED) as zf:
            zf.write(str(DB_PATH), BACKUP_DB_NAME)
            for _key, src in extras:
                zf.write(str(src), f"extra/{src.name}")
            zf.writestr(BACKUP_MANIFEST_NAME, json.dumps(manifest, indent=2))
        log.info("Backup created: %s", dest)
    except Exception:
        log.exception("Backup failed")
        return None
    _prune_old_backups(BACKUP_DIR, conn)
    _copy_backup_to_secondary(conn, dest)
    return dest


def _prune_old_backups(folder, conn):
    """Deletes backups (new-style .zip bundles and any older bare .db files
    already on disk from before this format existed) older than
    auto_backup_retention_days. Age-based only, on purpose -- a busy stretch
    of twice-daily scheduled backups should never be able to push something
    still inside the retention window out early, the way a count-based cap
    could."""
    try:
        retention_days = int(get_setting(conn, "auto_backup_retention_days", "60") or 60)
    except ValueError:
        retention_days = 60
    cutoff = datetime.now() - timedelta(days=retention_days)
    folder = Path(folder)
    if not folder.is_dir():
        return
    for pattern in ("data-*.zip", "data-*.db"):
        for path in folder.glob(pattern):
            try:
                mtime = datetime.fromtimestamp(path.stat().st_mtime)
                if mtime < cutoff:
                    path.unlink()
                    log.info("Pruned backup older than %s days: %s", retention_days, path)
            except Exception:
                log.exception("Could not check/prune backup %s", path)


def _copy_backup_to_secondary(conn, dest):
    """Best-effort copy of a just-taken backup into a second, user-chosen
    folder (e.g. a OneDrive/SharePoint-synced folder or a network share).
    Deliberately never raises -- the primary backup above is the one restore
    actually relies on, so a network drive being offline or a folder being
    temporarily unreachable should never be allowed to look like the whole
    backup failed. The outcome is recorded in settings instead, so Settings
    > Data & Backup can show it plainly. Also prunes the secondary folder by
    the same age-based retention as the local one -- previously it was never
    pruned at all."""
    if get_setting(conn, "secondary_backup_enabled", "0") != "1":
        return
    folder = get_setting(conn, "secondary_backup_folder", "").strip()
    if not folder:
        return
    try:
        target_dir = Path(folder)
        if not target_dir.is_dir():
            raise FileNotFoundError(f"Folder not found or not reachable: {folder}")
        shutil.copyfile(str(dest), str(target_dir / dest.name))
        set_settings(conn, {
            "secondary_backup_last_at": datetime.now().isoformat(timespec="seconds"),
            "secondary_backup_last_error": "",
        })
        log.info("Backup also copied to secondary folder: %s", target_dir / dest.name)
        _prune_old_backups(target_dir, conn)
    except Exception as e:
        log.exception("Secondary backup copy failed")
        set_settings(conn, {"secondary_backup_last_error": str(e)})


def backup_schedule_is_due(conn, now=None):
    """Optional EXTRA weekly scheduled backup, on top of (not instead of)
    the twice-daily weekday backups from ensure_backup_tasks_registered --
    off by default, for anyone who wants an additional one on a specific
    day. Returns True if it's turned on, the scheduled weekday has arrived
    (any time on that day, so this only needs to be checked once at app
    launch), and one hasn't already been taken this week."""
    now = now or datetime.now()
    if get_setting(conn, "auto_backup_schedule_enabled", "0") != "1":
        return False
    try:
        target_dow = int(get_setting(conn, "auto_backup_schedule_day", "0") or 0)
    except ValueError:
        target_dow = 0
    days_since = (now.weekday() - target_dow) % 7
    scheduled_date = (now - timedelta(days=days_since)).date()
    last_at = get_setting(conn, "auto_backup_last_scheduled_at", "")
    if last_at:
        try:
            if datetime.fromisoformat(last_at).date() >= scheduled_date:
                return False
        except ValueError:
            pass
    return now.weekday() == target_dow


def run_scheduled_backup_if_due(conn, now=None):
    """Takes the extra opt-in weekly scheduled backup if it's due, and
    records that it ran. Safe to call on every launch -- it's a no-op
    unless the scheduled day has arrived and this week's backup hasn't
    happened yet. Independent of the twice-daily weekday backups."""
    if not backup_schedule_is_due(conn, now):
        return None
    dest = backup_now(reason="scheduled-weekly")
    set_settings(conn, {"auto_backup_last_scheduled_at": (now or datetime.now()).isoformat(timespec="seconds")})
    return dest


def list_backups():
    """All local backups (new .zip bundles and any legacy bare .db files
    still on disk from before the bundle format existed), most recent
    first."""
    return sorted(
        list(BACKUP_DIR.glob("data-*.zip")) + list(BACKUP_DIR.glob("data-*.db")),
        key=lambda p: p.stat().st_mtime, reverse=True,
    )


_BACKUP_NAME_RE = re.compile(r"^data-(\d{8})-(\d{6})-(.+)\.(zip|db)$")


def describe_backup(path):
    """Parses a backup's timestamp/reason for display, preferring the
    manifest inside a .zip bundle (exact, includes what extra files it
    carries) and falling back to the filename (covers legacy .db backups,
    or a zip that for some reason has no manifest)."""
    path = Path(path)
    info = {
        "path": path, "name": path.name, "when": None, "reason": "",
        "has_extras": False, "is_bundle": path.suffix == ".zip",
    }
    if path.suffix == ".zip":
        try:
            with zipfile.ZipFile(str(path), "r") as zf:
                manifest = json.loads(zf.read(BACKUP_MANIFEST_NAME))
            info["reason"] = manifest.get("reason", "")
            info["has_extras"] = bool(manifest.get("extra_files"))
            try:
                info["when"] = datetime.fromisoformat(manifest.get("created_at", ""))
            except ValueError:
                pass
        except Exception:
            pass
    if info["when"] is None:
        m = _BACKUP_NAME_RE.match(path.name)
        if m:
            try:
                info["when"] = datetime.strptime(m.group(1) + m.group(2), "%Y%m%d%H%M%S")
            except ValueError:
                pass
            info["reason"] = info["reason"] or m.group(3)
    if info["when"] is None:
        try:
            info["when"] = datetime.fromtimestamp(path.stat().st_mtime)
        except Exception:
            pass
    return info


def list_local_backups_detailed():
    """Local backups only, described for display -- deliberately doesn't
    need a working database connection (unlike list_backups_detailed()
    below, which looks up the secondary-folder setting), which is exactly
    what makes it safe to use from the corrupt-database recovery screen:
    if the database itself is what's broken, there's no way to read the
    secondary_backup_folder setting out of it, so that screen can only
    ever offer local backups."""
    rows = []
    for path in list_backups():
        d = describe_backup(path)
        d["location"] = "Local"
        rows.append(d)
    return rows


def list_backups_detailed(conn=None):
    """Local + secondary backups together, each described with its
    date/time, reason, and where it lives, most recent first -- what the
    Settings > Data & Backup restore list is built from. Secondary-folder
    backups are included here even though they weren't previously browsable
    at all; restoring one just copies it in first."""
    conn = conn or get_connection()
    rows = list_local_backups_detailed()
    if get_setting(conn, "secondary_backup_enabled", "0") == "1":
        folder = get_setting(conn, "secondary_backup_folder", "").strip()
        if folder and Path(folder).is_dir():
            secondary_paths = sorted(
                list(Path(folder).glob("data-*.zip")) + list(Path(folder).glob("data-*.db")),
                key=lambda p: p.stat().st_mtime, reverse=True,
            )
            for path in secondary_paths:
                d = describe_backup(path)
                d["location"] = "Secondary (cloud/network folder)"
                rows.append(d)
    rows.sort(key=lambda d: d["when"] or datetime.min, reverse=True)
    return rows


def restore_backup(backup_path, allow_while_shared=False):
    """Restores a backup, handling both the current .zip bundle format
    (extracts the database and any bundled extra files, e.g. a custom logo,
    saving the logo into APP_DIR and pointing pdf_logo_path at the restored
    copy rather than whatever path it originally lived at, which might not
    exist any more) and legacy bare .db backups (just the database file, as
    every backup used to be).

    Batch 113: blocks entirely (raises SharedDatabaseRestoreBlockedError,
    touches nothing on disk) when this computer is using the shared
    database and the caller hasn't explicitly opted in via
    allow_while_shared=True. Reasoning: restoring an old backup just
    overwrites the local file, and the very next get_connection() call --
    which this function makes itself, right at the end -- would open that
    restored file as a turso embedded replica and immediately pull() the
    CURRENT shared database on top of it, silently undoing most or all of
    what the restore was meant to do (or worse, syncing an old file's
    replication bookkeeping against a much newer remote state is untested
    territory this project has no way to verify safely). The one place
    that deliberately passes allow_while_shared=True is the guided
    corrupt-database recovery flow (_DatabaseCorruptRecoveryDialog in
    po_generator_qt.py) -- there the local file is already unusable before
    the restore, so seeding it with ANY valid backup and then letting
    pull() catch it straight back up to the live shared state is exactly
    the right outcome, not a risk."""
    global _connection
    if not allow_while_shared and load_turso_config() is not None:
        raise SharedDatabaseRestoreBlockedError(
            "This computer is using the shared database, so restoring an old backup here isn't "
            "supported yet -- doing so would just get overwritten by the current shared data the "
            "next time this computer syncs, which could leave things in a confusing state. If you "
            "need to undo something in the shared database, hold off and get in touch first so we "
            "can work out the safest way to do it without affecting everyone else's data."
        )
    backup_path = Path(backup_path)
    if _connection is not None:
        try:
            _connection.commit()
            _connection.close()
        except Exception:
            pass
        _connection = None

    if backup_path.suffix == ".zip":
        with zipfile.ZipFile(str(backup_path), "r") as zf:
            names = zf.namelist()
            if BACKUP_DB_NAME not in names:
                raise ValueError(f"{backup_path.name} doesn't look like a valid backup (no database inside).")
            with zf.open(BACKUP_DB_NAME) as src, open(str(DB_PATH), "wb") as out:
                shutil.copyfileobj(src, out)
            manifest = {}
            if BACKUP_MANIFEST_NAME in names:
                try:
                    manifest = json.loads(zf.read(BACKUP_MANIFEST_NAME))
                except Exception:
                    manifest = {}
            restored_extra_settings = {}
            for entry in manifest.get("extra_files", []):
                archive_name = entry.get("archive_name", "")
                if archive_name not in names:
                    continue
                restored_dir = APP_DIR / "restored_from_backup"
                restored_dir.mkdir(parents=True, exist_ok=True)
                out_path = restored_dir / entry.get("original_name", Path(archive_name).name)
                with zf.open(archive_name) as src, open(str(out_path), "wb") as out:
                    shutil.copyfileobj(src, out)
                restored_extra_settings[entry["setting_key"]] = str(out_path)
    else:
        shutil.copyfile(str(backup_path), str(DB_PATH))
        restored_extra_settings = {}

    log.info("Restored backup from %s", backup_path)
    conn = get_connection()
    if restored_extra_settings:
        set_settings(conn, restored_extra_settings)
    return conn


class DatabaseCorruptError(Exception):
    """Raised by get_connection() when the SQLite database file exists but
    fails an integrity check (or isn't a valid SQLite file at all) --
    caught at startup so the app can offer a guided restore-from-backup
    instead of crashing or silently starting from an empty database."""
    pass


class SchemaTooNewError(Exception):
    """Raised by get_connection() when the database's stored schema version
    (PRAGMA user_version) is higher than this build of the app knows about --
    i.e. a newer copy of the app has already written structure or data this
    older copy doesn't understand. Caught at startup so the app can refuse
    to touch the database at all and tell the person to update, instead of
    quietly reading it wrong or writing something the newer copies choke on.
    Carries the two version numbers so the caller can build a clear message."""

    def __init__(self, db_schema_version, app_schema_version):
        self.db_schema_version = db_schema_version
        self.app_schema_version = app_schema_version
        super().__init__(
            f"Database schema version {db_schema_version} is newer than this "
            f"copy of the app understands (version {app_schema_version}). "
            "Update the app before opening this database."
        )


class SharedDatabaseUnreachableError(Exception):
    """Raised by get_connection() when a shared (turso) database is
    configured (Settings > Data & Backup > Shared database) but couldn't
    actually be opened -- bad/expired auth token, the host is unreachable,
    or this particular build doesn't include the 'turso' package at all.

    Batch 110: added specifically because the original Batch 109 design
    silently fell back to a plain local connection whenever this happened,
    reasoning that a broken shared-database setup should never lock
    someone out of their own data. Yitzi pointed out that's the wrong call
    for a database multiple people are meant to be sharing: quietly
    working from a stale local copy means whatever gets saved is invisible
    to everyone else until it happens to sync again, which can silently
    diverge from what the rest of the team sees. So once shared mode is
    configured, opening the database now genuinely requires reaching it --
    this is caught at startup and turned into a dialog with exactly two
    honest options: try again, or explicitly switch this computer back to
    local-only (never a silent default). Carries `detail`, a short
    human-readable reason to show the person."""

    def __init__(self, detail):
        self.detail = detail
        super().__init__(detail)


class BackendAlreadySetUpError(SharedDatabaseUnreachableError):
    """Raised by bootstrap_migration_token() specifically when the backend
    refused with HTTP 403 -- real user accounts already exist on it. This
    is the ONE expected, ordinary reason to fall back to a real login
    instead of bootstrap access (see po_generator_qt.py's
    SettingsPage._migrate_data_to_backend()). Every OTHER failure --
    unreachable, timed out, or (the real bug this was added for) the
    database's tables not existing yet because schema_postgres.sql was
    never run in Neon -- raises the plain base
    SharedDatabaseUnreachableError instead, and must be shown to the
    person directly rather than silently swallowed into a login prompt
    that could never work and would just say "invalid username" for a
    completely different, hidden reason."""
    pass


class ReadOnlyModeError(Exception):
    """Batch 144: raised by _ReadOnlyConnection.commit() (see below) when
    something tries to save while this computer is running in "View
    only" mode. po_generator_qt.py's global exception hook special-cases
    this to show a clear, specific message instead of the generic
    "something went wrong" one every other unexpected error gets, since
    this one is expected and has a genuinely helpful thing to tell the
    person: reconnect to the shared database to make changes again."""
    pass


class _ReadOnlyConnection:
    """Batch 144: wraps a connection to DB_PATH that's genuinely
    read-only at the SQLite level (opened via the 'mode=ro' URI, so
    SQLite itself refuses any write to the file -- this isn't just
    something this wrapper tries to catch in Python) for the new "View
    only" choice on the "Can't reach the shared database" startup
    dialog.

    Yitzi was firm, and correctly so, that the app should never quietly
    work off a stale local copy as if it were the real thing -- "no only
    need the online version in that case as it will cause data errors"
    is why SharedDatabaseUnreachableError exists at all, and why the old
    Batch 109/110 silent local-only fallback was removed in Batch 131.
    This is deliberately NOT a reintroduction of that: it's not a way to
    keep WORKING while disconnected, it's a way to LOOK at the last
    genuinely-synced copy of the data without any risk of local changes
    silently diverging from the shared database, because nothing can be
    saved through it at all, by design, enforced at the file level.

    execute()/executescript() are intercepted, NOT commit(): testing this
    directly against a real 'mode=ro' connection showed SQLite refuses a
    write the moment the DML statement itself is executed (a real
    sqlite3.OperationalError, "attempt to write a readonly database"),
    not deferred until a later commit() -- so that's the actual, correct
    place to turn it into a clear, specific ReadOnlyModeError instead of
    that less obvious raw message. Leaving commit() itself to pass
    through untouched (via __getattr__ below) matters too: a commit()
    that follows nothing but ordinary reads is completely harmless and
    must stay that way, since plenty of existing code calls commit() as
    a matter of routine regardless of whether anything was actually
    changed -- intercepting commit() unconditionally would have wrongly
    flagged those as errors as well. Everything else (row_factory,
    close, ...) passes straight through to the real connection via
    __getattr__/__setattr__, exactly like _PushOnCommitConnection above
    -- so this is invisible to the rest of the app's ordinary read
    logic."""

    def __init__(self, inner):
        object.__setattr__(self, "_inner", inner)

    def _reraise_as_read_only(self, e):
        if "readonly" in str(e).lower():
            raise ReadOnlyModeError(
                "This is a read-only view of the last data this computer synced -- "
                "changes can't be saved right now. Reconnect to the shared database "
                "once it's reachable again to make changes."
            ) from None
        raise e

    def execute(self, *args, **kwargs):
        try:
            return self._inner.execute(*args, **kwargs)
        except sqlite3.OperationalError as e:
            self._reraise_as_read_only(e)

    def executescript(self, *args, **kwargs):
        try:
            return self._inner.executescript(*args, **kwargs)
        except sqlite3.OperationalError as e:
            self._reraise_as_read_only(e)

    def __getattr__(self, name):
        return getattr(self._inner, name)

    def __setattr__(self, name, value):
        setattr(self._inner, name, value)


def is_view_only_mode():
    """True while this process's connection is the read-only fallback
    from open_view_only_connection() below, rather than a normal
    (local-only or shared) working connection. po_generator_qt.py's
    MainWindow uses this both to show a persistent, hard-to-miss banner
    and, in can(), to deny every permission except a plain ".view" one
    regardless of what the logged-in account would normally have -- a
    UI-level politeness layer on top of the real, DB-level guarantee
    _ReadOnlyConnection itself already provides."""
    return isinstance(_connection, _ReadOnlyConnection)


def view_only_synced_at():
    """A human-readable timestamp for when this computer's local replica
    file was last actually written to -- i.e. roughly when it last
    successfully synced with the shared database, since that file is
    only ever touched by a pull or a push (see get_connection()'s own
    docstring on why that file is normally a live mirror, not a stale
    snapshot). Used by MainWindow's view-only banner so "showing the
    last data this computer synced" isn't vague about how recent that
    actually was. Returns "" if DB_PATH doesn't exist at all (nothing to
    report a time for)."""
    if not DB_PATH.exists():
        return ""
    try:
        return datetime.fromtimestamp(DB_PATH.stat().st_mtime).strftime("%d %b %Y, %H:%M")
    except Exception:
        return ""


def open_view_only_connection():
    """Batch 144: the "View only" choice on the "Can't reach the shared
    database" dialog -- see _handle_shared_database_unreachable() in
    po_generator_qt.py. Opens DB_PATH as a genuinely read-only connection
    (see _ReadOnlyConnection above), sets it as the module's real global
    connection (so every existing core.get_connection() caller across
    the app, including MainWindow's own self.conn = core.get_connection(),
    naturally picks up this same read-only connection without needing to
    be told about it specially), and returns it.

    Deliberately makes NO shared-database connection attempt at all --
    this is a plain, local, read-only open of whatever this computer's
    embedded replica file already has on it. Because the shared-database
    setup keeps that same file continuously synced during ordinary,
    healthy use (pulling on every launch, pushing on every save -- see
    get_connection()'s own Batch 109/110 docstring), that file is
    normally very recent, not some old forgotten snapshot, even though
    it can't be refreshed any further while the shared database itself
    can't be reached.

    Raises DatabaseCorruptError via the same check_db_integrity() every
    other path already uses, and RuntimeError if there's no local file
    at all yet to open (this computer has never successfully synced
    even once, so there's genuinely nothing to view)."""
    global _connection
    ok, detail = check_db_integrity(DB_PATH)
    if not ok:
        raise DatabaseCorruptError(detail)
    if not DB_PATH.exists():
        raise RuntimeError(
            "There's no local copy of the database on this computer to view yet -- "
            "this computer has never successfully connected to the shared database."
        )
    inner = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True, check_same_thread=False)
    inner.row_factory = sqlite3.Row
    _connection = _ReadOnlyConnection(inner)
    return _connection


class SharedDatabaseRestoreBlockedError(Exception):
    """Raised by restore_backup() when this computer is using the shared
    database and the caller didn't explicitly opt in with
    allow_while_shared=True. See restore_backup()'s own docstring for the
    full reasoning -- in short, a plain local-backup restore would just be
    silently undone by the next sync from the shared database. Carries
    `detail`, a short human-readable reason to show the person."""

    def __init__(self, detail):
        self.detail = detail
        super().__init__(detail)


def _parse_version(v):
    """Turns "3.4.0" into (3, 4, 0) for a real numeric comparison -- comparing
    the strings directly would sort "3.10.0" before "3.9.0". Anything that
    isn't a clean run of dot-separated integers (missing, blank, malformed)
    comes back as (0,), which sorts below every real version rather than
    raising -- a bad manifest should never crash the update check."""
    try:
        parts = tuple(int(p) for p in str(v).strip().split("."))
        return parts if parts else (0,)
    except (ValueError, AttributeError):
        return (0,)


def check_for_update(manifest_url=None, timeout=5):
    """Looks up the latest published version from manifest_url (falls back
    to UPDATE_MANIFEST_URL) and compares it against APP_VERSION. Returns
    None when: checking is switched off (no URL configured), the check
    fails for any reason at all -- no internet, the host is unreachable, a
    malformed response -- or this copy is already current. A failed check
    must never stop the program opening; plenty of people using this will
    be somewhere with no signal. Otherwise returns a dict:
        {"latest_version": str, "minimum_required_version": str,
         "download_url": str, "notes": str, "required": bool}
    "required" is True once the manifest's minimum_required_version is
    higher than APP_VERSION -- the caller is expected to block on that, not
    just mention it (see _show_forced_update_block in po_generator_qt.py)."""
    url = manifest_url or UPDATE_MANIFEST_URL
    if not url:
        return None
    try:
        with urllib.request.urlopen(url, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except Exception:
        log.info("Update check failed or skipped", exc_info=True)
        return None

    latest = str(data.get("latest_version") or "").strip()
    minimum = str(data.get("minimum_required_version") or "").strip()
    if not latest:
        return None

    current_v = _parse_version(APP_VERSION)
    required = bool(minimum) and _parse_version(minimum) > current_v
    if _parse_version(latest) <= current_v and not required:
        return None

    return {
        "latest_version": latest,
        "minimum_required_version": minimum,
        "download_url": str(data.get("download_url") or "").strip(),
        # Batch 113: optional -- an older manifest (or one Yitzi forgets to
        # fill in) simply won't have it, in which case download_update()
        # below skips the checksum check rather than failing. Present
        # whenever possible though, since it's the only thing standing
        # between "click Update now" and a computer running whatever file
        # actually sat at download_url at the time, unverified.
        "sha256": str(data.get("sha256") or "").strip(),
        "notes": str(data.get("notes") or "").strip(),
        "required": required,
    }


class UpdateDownloadError(Exception):
    """Raised by download_update() on any failure: can't reach the URL, a
    response with nothing in it, or (when the manifest provided one) a
    sha256 checksum mismatch. Carries `detail`, a short human-readable
    reason to show the person. Deliberately the only exception type
    download_update() ever raises -- callers don't need to catch anything
    else to handle every way this can go wrong."""

    def __init__(self, detail):
        self.detail = detail
        super().__init__(detail)


def download_update(download_url, dest_path, expected_sha256="", timeout=30, progress_callback=None, should_cancel=None):
    """Batch 113: downloads a new build of the app from download_url (the
    manifest's own "download_url", see check_for_update() above) and saves
    it to dest_path -- the file the one-click "Update now" flow in
    po_generator_qt.py then hands off to a small relauncher script (see
    build_update_relauncher_script() below) to actually swap into place.

    Written to a temporary "<dest_path>.part" file first and only renamed
    into place at the very end, on full success -- a failed or cancelled
    download (or a checksum mismatch) never leaves anything at dest_path
    that could be mistaken for a complete, verified download.

    If expected_sha256 is given (the manifest's own optional "sha256"
    field) the downloaded bytes' own sha256 is checked against it before
    the file is accepted -- this is the one thing standing between
    clicking "Update now" and running whatever happened to be sitting at
    that URL, so a mismatch is treated as seriously as a network failure:
    the download is thrown away and UpdateDownloadError is raised, nothing
    gets installed. If the manifest didn't provide a checksum at all (an
    older manifest, or Yitzi forgot to fill it in), the download still
    proceeds -- unverified is still how manual downloads via the plain
    link have always worked, so this isn't a new risk, just an unverified
    one, same as before this feature existed.

    progress_callback(bytes_read, total_bytes), if given, is called after
    each chunk (total_bytes is 0 if the server didn't send a
    Content-Length header, which callers should treat as "unknown size").
    should_cancel(), if given, is checked between chunks too -- returning
    True stops the download and raises UpdateDownloadError with a plain
    "cancelled" message, same cleanup as any other failure.

    Only ever raises UpdateDownloadError, regardless of what actually went
    wrong underneath, so callers only need to catch one exception type."""
    dest_path = Path(dest_path)
    part_path = dest_path.with_suffix(dest_path.suffix + ".part")
    hasher = hashlib.sha256()
    bytes_read = 0
    try:
        try:
            resp = urllib.request.urlopen(download_url, timeout=timeout)
        except Exception as e:
            raise UpdateDownloadError(f"Couldn't reach the download link: {e}")
        with resp:
            try:
                total_bytes = int(resp.headers.get("Content-Length") or 0)
            except (TypeError, ValueError):
                total_bytes = 0
            part_path.parent.mkdir(parents=True, exist_ok=True)
            with open(part_path, "wb") as f:
                while True:
                    if should_cancel and should_cancel():
                        raise UpdateDownloadError("Update cancelled -- nothing on this computer was changed.")
                    chunk = resp.read(262144)
                    if not chunk:
                        break
                    f.write(chunk)
                    hasher.update(chunk)
                    bytes_read += len(chunk)
                    if progress_callback:
                        progress_callback(bytes_read, total_bytes)
        if bytes_read == 0:
            raise UpdateDownloadError("The download came back empty, so nothing was installed.")
        if expected_sha256:
            actual = hasher.hexdigest()
            if actual.lower() != expected_sha256.strip().lower():
                raise UpdateDownloadError(
                    "The downloaded file didn't match the expected checksum, so it wasn't installed "
                    "(it may have been corrupted in transit, or the published checksum doesn't match "
                    "the file at the download link). Nothing on this computer was changed."
                )
        if dest_path.exists():
            dest_path.unlink()
        part_path.replace(dest_path)
        return dest_path
    except UpdateDownloadError:
        try:
            if part_path.exists():
                part_path.unlink()
        except Exception:
            pass
        raise
    except Exception as e:
        try:
            if part_path.exists():
                part_path.unlink()
        except Exception:
            pass
        raise UpdateDownloadError(str(e))


def build_update_relauncher_script(current_exe_path, installer_path, pid, log_path=None):
    """Batch 113, reworked in Batch 130, and again in Batch 139: returns
    the text of a small Windows batch script that waits for this process
    (pid) to fully exit, then runs installer_path silently (installer_path
    is now the downloaded Inno Setup installer, not a raw replacement .exe
    -- see below for why) and relaunches current_exe_path once that's done
    -- the actual mechanism behind the "Update now" button swapping the
    running program out from under itself.

    Batch 130 background: the app used to ship as a single PyInstaller
    --onefile .exe, which this script updated by literally moving a new
    .exe over the old one. --onefile re-extracts its entire bundle into a
    temp folder on EVERY launch though, not just the first -- which
    turned out to be the real cause of Yitzi's "2-4 min[ute]... on all
    PCs" startup delay, not anything fixable in this app's own Python
    code. The fix was switching to --onedir (unpacks once, at build time)
    wrapped in a proper Inno Setup installer (installer.iss) instead of a
    portable .exe -- which means there's no longer a single file for this
    script to "move" into place; an entire installed folder needs
    replacing instead, which is exactly what running the installer itself
    (silently, unattended) already knows how to do correctly.

    This still exists because Windows keeps a running program's files
    locked, so nothing can overwrite them while the app is still running
    -- same reason as always, just now solved by re-running the installer
    instead of moving one file. The app writes this script to a temp
    file, launches it as a separate process, then exits itself -- the
    script's wait loop only proceeds once this process's PID has
    genuinely gone, so the install can never race the still-running app
    for a file lock. installer.iss's own [Run] entry that would normally
    launch the app after install is deliberately skipped during a silent
    install (its "skipifsilent" flag) specifically so this script -- not
    the installer -- is what's responsible for reopening the app,
    avoiding any chance of it opening twice. On a clean run, this script
    deletes itself (the "del %~f0" line) so it doesn't linger in the temp
    folder; on a failure, it leaves itself and the log behind instead (see
    below) so there's something to look at afterward.

    Batch 139: Yitzi reported clicking "Update now," seeing the "will now
    close and reopen" message, the app closing -- and then nothing: "the
    program doesn't open again... There's no indication if it's installed
    the updates or not. I don't know how long to wait." Reading how this
    script used to get launched (po_generator_qt.py's _run_one_click_update,
    with the Windows-only DETACHED_PROCESS creation flag) explains that
    part precisely: DETACHED_PROCESS gives the relauncher no console at
    all, by design -- so from the moment the app closes, this entire
    wait-install-reopen sequence has always run completely invisibly,
    whether it takes five seconds or genuinely gets stuck. That's a real
    gap regardless of what (if anything) is actually going wrong
    underneath, so this script now assumes it will be launched into a
    real, visible console window instead (see _run_one_click_update's own
    CREATE_NEW_CONSOLE change) and narrates each step -- waiting for the
    old copy to close, installing, reopening -- so there's always
    something on screen saying what's currently happening.

    The second half of the fix is for the "doesn't reopen at all" case
    specifically, which this sandbox has no way to reproduce or confirm
    the exact cause of (no real Windows machine to run this script or a
    real Inno Setup install on). Rather than guess at a specific root
    cause blind, this makes every way the old version of this script
    could have failed *silently* fail *loudly* instead: the installer is
    now run with Inno Setup's own /LOG switch (writing a real install log
    to log_path, when one is given, instead of nothing) and this script
    checks the installer's exit code -- on a nonzero one, it prints the
    error code and the log's location and pauses instead of quietly
    running `del %~f0` and vanishing. It also checks current_exe_path
    genuinely exists before trying to relaunch it (covering the unlikely
    case where the install "succeeded" but somehow didn't leave a working
    .exe where expected) rather than silently doing nothing if that start
    command fails. Either failure path leaves the script and its log
    sitting in the temp folder instead of deleting them, specifically so
    a log file exists to send back next time this happens, rather than
    another report of "it just didn't come back" with nothing to go on.

    Pure string-building, no filesystem or process access of its own --
    that split is deliberate, so the script's own logic can be checked
    from this Linux sandbox (which can write and read the text, but can
    never actually run a .bat file, run a real Inno Setup installer, or
    observe a real Windows process exiting) without needing a real
    Windows machine."""
    current_exe_path = str(current_exe_path)
    installer_path = str(installer_path)
    log_arg = f' /LOG="{log_path}"' if log_path else ""
    log_line = f"    echo A log with the details was saved to:\r\n    echo   {log_path}\r\n" if log_path else ""
    return (
        "@echo off\r\n"
        "setlocal\r\n"
        'title YJ PO Generator - Installing update\r\n'
        "echo Waiting for YJ PO Generator to close...\r\n"
        ":wait\r\n"
        f'tasklist /FI "PID eq {pid}" 2>NUL | find /I "{pid}" >NUL\r\n'
        "if not errorlevel 1 (\r\n"
        "    timeout /t 1 /nobreak >nul\r\n"
        "    goto wait\r\n"
        ")\r\n"
        "echo Installing the update, this can take a minute or two -- please don't close this window...\r\n"
        f'start /wait "" "{installer_path}" /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /CLOSEAPPLICATIONS{log_arg}\r\n'
        "if errorlevel 1 (\r\n"
        "    echo.\r\n"
        "    echo Something went wrong installing the update (error code %errorlevel%^).\r\n"
        f"{log_line}"
        "    echo.\r\n"
        "    echo Please reopen YJ PO Generator yourself for now, and send this window (or the log^) back so it can be looked into.\r\n"
        "    echo Press any key to close this window.\r\n"
        "    pause >nul\r\n"
        "    exit /b 1\r\n"
        ")\r\n"
        f'if not exist "{current_exe_path}" (\r\n'
        "    echo.\r\n"
        "    echo The update installed, but YJ PO Generator couldn't be found to reopen automatically.\r\n"
        "    echo Please open it yourself from the Start Menu or desktop shortcut.\r\n"
        "    echo Press any key to close this window.\r\n"
        "    pause >nul\r\n"
        "    exit /b 1\r\n"
        ")\r\n"
        "echo Update installed. Reopening YJ PO Generator...\r\n"
        f'start "" "{current_exe_path}"\r\n'
        "timeout /t 2 /nobreak >nul\r\n"
        'del "%~f0"\r\n'
    )


def check_db_integrity(path=None):
    """Runs SQLite's own integrity check against a database file without
    going through the shared global connection (so it's safe to call before
    deciding whether that connection is even safe to open). Returns
    (True, "ok") if the file is missing (nothing to check -- a brand new
    install) or genuinely fine, otherwise (False, <detail>).

    Batch 109: deliberately unchanged for the shared-database (turso) path.
    A turso embedded replica is still a genuine local SQLite-format file on
    disk at DB_PATH -- that's what makes the embedded-replica model work at
    all -- so a plain, short-lived sqlite3 connection can still open and
    check it here exactly as before, whether or not the shared database is
    switched on."""
    path = Path(path) if path else DB_PATH
    if not path.exists():
        return True, "ok"
    try:
        test_conn = sqlite3.connect(str(path))
        try:
            row = test_conn.execute("PRAGMA integrity_check").fetchone()
            result = row[0] if row else "unknown error"
        finally:
            test_conn.close()
        return (result == "ok"), result
    except sqlite3.Error as e:
        return False, str(e)


def export_all_json(path):
    conn = get_connection()
    data = {
        "exported_at": now_iso(),
        "app_version": APP_VERSION,
        "settings": get_all_settings(conn),
        "addresses": list_addresses(conn),
        "suppliers": list_suppliers(conn),
        "products": [dict(r) for r in conn.execute("SELECT * FROM products")],
        "purchase_orders": [],
    }
    for row in conn.execute("SELECT * FROM purchase_orders"):
        po = dict(row)
        items = conn.execute(
            "SELECT id, qty, code, product, price, margin_vat, original_qty FROM po_items "
            "WHERE po_id=? ORDER BY position", (po["id"],)
        ).fetchall()
        po["items"] = [dict(i) for i in items]
        data["purchase_orders"].append(po)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    return len(data["purchase_orders"])


def import_all_json(path, mode="merge"):
    """mode='merge' upserts on top of existing data (safe default).
    mode='replace' wipes purchase_orders/items/products/suppliers/addresses first."""
    conn = get_connection()
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    if mode == "replace":
        conn.execute("DELETE FROM po_items")
        conn.execute("DELETE FROM po_events")
        conn.execute("DELETE FROM purchase_orders")
        conn.execute("DELETE FROM products")
        conn.commit()

    if data.get("settings"):
        set_settings(conn, data["settings"])
    if data.get("addresses"):
        if mode == "replace":
            replace_addresses(conn, data["addresses"])
        else:
            existing_labels = {a["label"] for a in list_addresses(conn)}
            merged = list_addresses(conn) + [a for a in data["addresses"] if a.get("label") not in existing_labels]
            replace_addresses(conn, merged)
    if data.get("suppliers"):
        if mode == "replace":
            replace_suppliers(conn, data["suppliers"])
        else:
            existing = list_suppliers(conn)
            existing_names = {s["company_name"] for s in existing}
            merged = existing + [s for s in data["suppliers"] if s.get("company_name") not in existing_names]
            replace_suppliers(conn, merged)

    imported_pos = 0
    supplier_ids_by_name = {r["company_name"]: r["id"] for r in conn.execute("SELECT id, company_name FROM suppliers")}
    for po in data.get("purchase_orders", []):
        po_ref = po.get("po_ref")
        if not po_ref:
            continue
        items = po.get("items", [])
        po_row = {k: po.get(k, "") for k in PO_FIELDS}
        po_row["po_ref"] = po_ref
        po_row["supplier_id"] = supplier_ids_by_name.get(po.get("supplier_company_name", ""))
        try:
            save_po(conn, po_row, items, status=po.get("status", "Draft"), event="imported",
                     record_product_memory=True, created_at=po.get("created_at"), updated_at=po.get("updated_at"))
            if po.get("deleted"):
                set_po_deleted(conn, po_ref, True)
            imported_pos += 1
        except Exception:
            log.exception("Failed importing PO %s", po_ref)
    return imported_pos


# ============================================================
# 8. Reporting queries
#
# These back the Reports section of the UI: spend-by-supplier, spend-by-
# product, spend-over-time (for trend charts), price-change detection (the
# "price creep" alerts), and a generic grouped query that powers the custom
# report builder. All pure data-in, data-out — no UI framework involved, so
# both UIs (and tests) can use them identically.
# ============================================================

def report_spend_by_supplier(conn, date_from=None, date_to=None, status=None, deleted=False):
    # Sums base_total (each PO's home-currency equivalent), not the raw
    # per-PO total -- suppliers invoiced in a different currency would
    # otherwise silently get blended into this in the wrong units.
    sql = (
        "SELECT supplier_company_name AS supplier, COUNT(*) AS po_count, "
        "COALESCE(SUM(base_total), 0) AS total, COALESCE(AVG(base_total), 0) AS avg_po_value "
        "FROM purchase_orders WHERE deleted=?"
    )
    params = [1 if deleted else 0]
    if date_from:
        sql += " AND date(created_at) >= date(?)"
        params.append(date_from)
    if date_to:
        sql += " AND date(created_at) <= date(?)"
        params.append(date_to)
    if status:
        sql += " AND status=?"
        params.append(status)
    sql += " GROUP BY supplier_company_name ORDER BY total DESC"
    return [dict(r) for r in conn.execute(sql, params)]


def report_spend_by_product(conn, date_from=None, date_to=None, supplier=None, status=None, deleted=False):
    # i.price * po.fx_rate converts each line into the home currency before
    # summing/averaging -- the same product bought from suppliers billing
    # in different currencies would otherwise mix raw numbers together.
    sql = (
        "SELECT i.product AS product, i.code AS code, po.supplier_company_name AS supplier, "
        "SUM(i.qty) AS total_qty, SUM(i.qty * i.price * po.fx_rate) AS total_spend, "
        "AVG(i.price * po.fx_rate) AS avg_price, COUNT(DISTINCT po.id) AS times_ordered "
        "FROM po_items i JOIN purchase_orders po ON i.po_id = po.id WHERE po.deleted=?"
    )
    params = [1 if deleted else 0]
    if date_from:
        sql += " AND date(po.created_at) >= date(?)"
        params.append(date_from)
    if date_to:
        sql += " AND date(po.created_at) <= date(?)"
        params.append(date_to)
    if supplier:
        sql += " AND po.supplier_company_name=?"
        params.append(supplier)
    if status:
        sql += " AND po.status=?"
        params.append(status)
    sql += " GROUP BY i.product, po.supplier_company_name ORDER BY total_spend DESC"
    return [dict(r) for r in conn.execute(sql, params)]


def report_spend_over_time(conn, date_from=None, date_to=None, granularity="month", deleted=False):
    fmt = {"day": "%Y-%m-%d", "week": "%Y-W%W", "year": "%Y"}.get(granularity, "%Y-%m")
    sql = (
        f"SELECT strftime('{fmt}', created_at) AS period, COALESCE(SUM(base_total), 0) AS total, "
        "COUNT(*) AS po_count FROM purchase_orders WHERE deleted=?"
    )
    params = [1 if deleted else 0]
    if date_from:
        sql += " AND date(created_at) >= date(?)"
        params.append(date_from)
    if date_to:
        sql += " AND date(created_at) <= date(?)"
        params.append(date_to)
    sql += " GROUP BY period ORDER BY period"
    return [dict(r) for r in conn.execute(sql, params)]


def report_spend_by_status(conn, date_from=None, date_to=None, supplier=None, deleted=False):
    sql = "SELECT status, COUNT(*) AS po_count, COALESCE(SUM(base_total), 0) AS total FROM purchase_orders WHERE deleted=?"
    params = [1 if deleted else 0]
    if date_from:
        sql += " AND date(created_at) >= date(?)"
        params.append(date_from)
    if date_to:
        sql += " AND date(created_at) <= date(?)"
        params.append(date_to)
    if supplier:
        sql += " AND supplier_company_name=?"
        params.append(supplier)
    sql += " GROUP BY status ORDER BY total DESC"
    return [dict(r) for r in conn.execute(sql, params)]


def report_custom(conn, group_by="supplier", date_from=None, date_to=None,
                   supplier=None, status=None, deleted=False):
    """Powers the custom report builder: group_by one of
    'supplier' | 'product' | 'month' | 'status', with the same filter set
    (date range / supplier / status / deleted) applied across all of them."""
    if group_by == "product":
        return report_spend_by_product(conn, date_from, date_to, supplier, status, deleted)
    if group_by == "month":
        return report_spend_over_time(conn, date_from, date_to, "month", deleted)
    if group_by == "status":
        return report_spend_by_status(conn, date_from, date_to, supplier, deleted)
    return report_spend_by_supplier(conn, date_from, date_to, status, deleted)


def report_outstanding_orders(conn):
    """POs that have been sent but not yet marked Received/Cancelled — the
    'what am I still waiting on' view."""
    rows = conn.execute(
        "SELECT po_ref, supplier_company_name AS supplier, status, total, currency, updated_at "
        "FROM purchase_orders WHERE deleted=0 AND status NOT IN ('Received', 'Cancelled', 'Draft') "
        "ORDER BY updated_at ASC"
    ).fetchall()
    out = []
    now = datetime.now()
    for r in rows:
        d = dict(r)
        try:
            age_days = (now - datetime.fromisoformat(d["updated_at"])).days
        except Exception:
            age_days = None
        d["age_days"] = age_days
        out.append(d)
    return out


def report_price_changes(conn, deleted=False):
    """For every (supplier, product) pair ordered more than once, compares
    the two most recent prices paid. Returns only the ones where price went
    UP, sorted biggest increase first — the 'price creep' alerts that let you
    catch a supplier's prices drifting up before it costs you real money."""
    rows = conn.execute(
        "SELECT po.supplier_company_name AS supplier, i.code AS code, i.product AS product, "
        "i.price AS price, po.created_at AS at, po.po_ref AS po_ref "
        "FROM po_items i JOIN purchase_orders po ON i.po_id = po.id "
        "WHERE po.deleted=? AND i.product != '' "
        "ORDER BY po.supplier_company_name, i.product, po.created_at",
        (1 if deleted else 0,),
    ).fetchall()
    groups = {}
    for r in rows:
        key = (r["supplier"] or "", r["product"] or "")
        groups.setdefault(key, []).append(dict(r))

    alerts = []
    for (supplier, product), history in groups.items():
        if len(history) < 2:
            continue
        prev = history[-2]
        last = history[-1]
        try:
            prev_price = float(prev["price"])
            last_price = float(last["price"])
        except Exception:
            continue
        if prev_price > 0 and last_price > prev_price:
            pct = (last_price - prev_price) / prev_price * 100
            alerts.append({
                "supplier": supplier,
                "product": product,
                "code": last["code"],
                "previous_price": prev_price,
                "current_price": last_price,
                "change_pct": pct,
                "previous_at": prev["at"],
                "current_at": last["at"],
                "previous_po_ref": prev["po_ref"],
                "current_po_ref": last["po_ref"],
                "times_ordered": len(history),
            })
    alerts.sort(key=lambda a: a["change_pct"], reverse=True)
    return alerts


# ---- Periodic (weekly/monthly/quarterly/yearly) PO summary report email ----
#
# This app has no background server, so it can't truly send email on a
# schedule by itself. What it can do: on launch (or whenever asked), check
# whether a report is "due" for the current period and hasn't been sent yet,
# and if so, build the draft and open it ready-to-send in Outlook -- the
# person still reviews and hits Send themselves, same as every other email
# this app prepares.

REPORT_FREQUENCIES = ("weekly", "monthly", "quarterly", "yearly")
# Frequency string -> the plain-English period noun used throughout this
# section (labels, comparison section headings, subjects).
_PERIOD_NOUN = {"weekly": "week", "monthly": "month", "quarterly": "quarter", "yearly": "year"}


def normalize_report_frequency(freq):
    """A po_report_frequency setting value is trusted to already be one of
    REPORT_FREQUENCIES (the Settings combo only ever writes one of those),
    but this is the one place every "what word do we show the user"
    call site funnels through, so a stale/unrecognised value (e.g. read
    from an old settings file, or a hand-edited one) always falls back to
    "weekly" instead of silently producing an empty/wrong label."""
    return freq if freq in REPORT_FREQUENCIES else "weekly"


def _week_bounds_containing(d):
    """Monday 00:00:00 through Sunday 23:59:59 of the week containing date
    or datetime d."""
    d = d if isinstance(d, datetime) else datetime(d.year, d.month, d.day)
    monday = (d - timedelta(days=d.weekday())).replace(hour=0, minute=0, second=0, microsecond=0)
    sunday_end = (monday + timedelta(days=6)).replace(hour=23, minute=59, second=59, microsecond=0)
    return monday, sunday_end


def month_bounds(year, month):
    """1st 00:00:00 through the last day 23:59:59 of the given calendar
    month."""
    start = datetime(year, month, 1)
    next_start = datetime(year + 1, 1, 1) if month == 12 else datetime(year, month + 1, 1)
    end = (next_start - timedelta(seconds=1)).replace(hour=23, minute=59, second=59, microsecond=0)
    return start, end


def quarter_bounds(year, quarter):
    """quarter: 1-4 (Jan-Mar, Apr-Jun, Jul-Sep, Oct-Dec) -- 1st of the
    quarter's first month 00:00:00 through the last day of its third month
    23:59:59."""
    quarter = max(1, min(int(quarter), 4))
    start_month = (quarter - 1) * 3 + 1
    start, _ = month_bounds(year, start_month)
    _, end = month_bounds(year, start_month + 2)
    return start, end


def year_bounds(year):
    """Plain calendar year -- 1 Jan 00:00:00 through 31 Dec 23:59:59. A
    deliberately different concept from the Savings feature's own
    configurable financial year (savings_fiscal_year_bounds) -- PO period
    reporting has no equivalent "financial year start month" setting, so
    "yearly" here always means the calendar year."""
    return datetime(year, 1, 1), datetime(year, 12, 31, 23, 59, 59)


def period_bounds_for(period_kind, year, sub=None):
    """Bounds for a SPECIFIC, explicitly chosen period instance -- used for
    manually generating a report for a past period (change request:
    "allow reports to be manually generated for previous periods"), as
    opposed to _period_bounds below (the currently *configured schedule's*
    "this"/"last" window relative to now).

    period_kind: "week"/"month"/"quarter"/"year".
    sub: for "week", a date/datetime that falls within the target week
    (year is ignored -- the date alone identifies the week); for "month",
    the month number (1-12); for "quarter", the quarter number (1-4); for
    "year", ignored."""
    if period_kind == "week":
        return _week_bounds_containing(sub)
    if period_kind == "month":
        return month_bounds(year, sub)
    if period_kind == "quarter":
        return quarter_bounds(year, sub)
    if period_kind == "year":
        return year_bounds(year)
    raise ValueError(f"Unknown period_kind: {period_kind!r}")


def previous_period_bounds(period_kind, period_start, period_end):
    """Bounds of the period immediately preceding [period_start, period_end]
    -- same kind, used for "each report should compare its selected period
    with the equivalent previous period" (change request section 5).
    Calendar-aware for month/quarter/year (the previous month/quarter/year
    is always the WHOLE previous calendar one, not a fixed day-count shift
    back -- months and quarters vary in length), a plain 7-day shift back
    for week."""
    if period_kind == "week":
        return period_start - timedelta(days=7), period_end - timedelta(days=7)
    if period_kind == "month":
        prev_end = (period_start - timedelta(seconds=1)).replace(hour=23, minute=59, second=59, microsecond=0)
        prev_start = prev_end.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        return prev_start, prev_end
    if period_kind == "quarter":
        prev_end = (period_start - timedelta(seconds=1)).replace(hour=23, minute=59, second=59, microsecond=0)
        # period_start is always the 1st of a quarter's first month (Jan/
        # Apr/Jul/Oct) when produced by quarter_bounds/_period_bounds, so
        # walking back exactly 3 calendar months always lands on the
        # previous quarter's own first month -- safer than a day-count
        # shift, since quarters span 90-92 days depending which they are.
        month = period_start.month - 3
        year = period_start.year
        if month <= 0:
            month += 12
            year -= 1
        prev_start = datetime(year, month, 1)
        return prev_start, prev_end
    if period_kind == "year":
        return datetime(period_start.year - 1, 1, 1), datetime(period_start.year - 1, 12, 31, 23, 59, 59)
    raise ValueError(f"Unknown period_kind: {period_kind!r}")


def _period_bounds_for_freq(conn, freq, now=None):
    """Returns (period_start, period_end, scheduled_dt) for ONE specific
    report frequency's OWN independent schedule -- the multi-frequency
    replacement for the old single-schedule _period_bounds (Batch 99:
    weekly/monthly/quarterly/yearly can now all be enabled at once, each
    on its own po_report_{freq}_enabled/po_report_{freq}_day, rather than
    only one of them ever being "the" active schedule).

    scheduled_dt is only used to decide *whether* a report is due yet /
    already sent for this period (see reports_due), based on the
    configured send day and time.

    period_start/period_end are the actual content window, controlled by
    the shared po_report_window setting (one preference across all four
    frequencies, not per-frequency):
      - "this" (default): the full current period -- Monday through Sunday
        of this calendar week for weekly (e.g. checked on a Friday, still
        shows Monday to Sunday of that same week, not truncated at today),
        the 1st through the last day of this month for monthly, the
        current calendar quarter for quarterly, or the current calendar
        year for yearly. Always anchored to the period's real start and
        never stretches beyond one period, whatever day it's configured to
        send on or how late it's actually checked.
      - "last": the most recently *finished* period instead.

    Monthly/quarterly/yearly are all scheduled on their own
    po_report_{freq}_day as the day-of-month of the period's first month
    (weekly uses po_report_weekly_day as a day-of-week instead), at the
    shared po_report_time."""
    now = now or datetime.now()
    freq = normalize_report_frequency(freq)
    window = get_setting(conn, "po_report_window", "this")
    try:
        hh, mm = (get_setting(conn, "po_report_time", "09:00") or "09:00").split(":")
        hh, mm = int(hh), int(mm)
    except Exception:
        hh, mm = 9, 0

    if freq in ("monthly", "quarterly", "yearly"):
        try:
            day = int(get_setting(conn, f"po_report_{freq}_day", "1") or 1)
        except ValueError:
            day = 1
        day = max(1, min(day, 28))
        if freq == "quarterly":
            this_start, this_end = quarter_bounds(now.year, (now.month - 1) // 3 + 1)
        elif freq == "yearly":
            this_start, this_end = year_bounds(now.year)
        else:
            this_start, this_end = month_bounds(now.year, now.month)
        scheduled_dt = this_start.replace(day=day, hour=hh, minute=mm, second=0, microsecond=0)
        if window == "last":
            period_start, period_end = previous_period_bounds(_PERIOD_NOUN[freq], this_start, this_end)
        else:
            period_start, period_end = this_start, this_end
    else:
        try:
            target_dow = int(get_setting(conn, "po_report_weekly_day", "0") or 0)
        except ValueError:
            target_dow = 0
        days_since = (now.weekday() - target_dow) % 7
        scheduled_dt = (now - timedelta(days=days_since)).replace(hour=hh, minute=mm, second=0, microsecond=0)
        # Monday of the week containing "now" (datetime.weekday(): Monday=0
        # ... Sunday=6).
        this_monday = (now - timedelta(days=now.weekday())).replace(
            hour=0, minute=0, second=0, microsecond=0
        )
        this_sunday_end = (this_monday + timedelta(days=6)).replace(
            hour=23, minute=59, second=59, microsecond=0
        )
        if window == "last":
            period_start = this_monday - timedelta(days=7)
            period_end = (this_monday - timedelta(seconds=1)).replace(
                hour=23, minute=59, second=59, microsecond=0
            )
        else:
            period_start = this_monday
            period_end = this_sunday_end
    return period_start, period_end, scheduled_dt


def dashboard_period_bounds(period, which, now=None):
    """Simple Monday-Sunday week or calendar month bounds for the
    Dashboard's "Spend this period"/"Spend last period" KPI cards.

    Deliberately independent of _period_bounds_for_freq()/the po_report_*
    settings above -- those describe when the periodic PO summary EMAIL gets sent
    (which can be any weekday), while the Dashboard's "this week" should
    always mean the literal current Monday-Sunday, regardless of what day
    reports go out on.

    period: "week" or "month". which: "this" or "last".
    """
    now = now or datetime.now()
    if period == "month":
        this_start = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        next_start = (this_start + timedelta(days=32)).replace(day=1)
        this_end = (next_start - timedelta(seconds=1)).replace(hour=23, minute=59, second=59, microsecond=0)
        if which == "last":
            end = (this_start - timedelta(seconds=1)).replace(hour=23, minute=59, second=59, microsecond=0)
            start = end.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
            return start, end
        return this_start, this_end
    else:
        this_monday = (now - timedelta(days=now.weekday())).replace(hour=0, minute=0, second=0, microsecond=0)
        this_sunday_end = (this_monday + timedelta(days=6)).replace(hour=23, minute=59, second=59, microsecond=0)
        if which == "last":
            start = this_monday - timedelta(days=7)
            end = (this_monday - timedelta(seconds=1)).replace(hour=23, minute=59, second=59, microsecond=0)
            return start, end
        return this_monday, this_sunday_end


def dashboard_period_spend(conn, period, which, now=None):
    """Total (non-deleted) PO spend for dashboard_period_bounds(period, which),
    in the home currency (base_total), so a mix of GBP/USD/etc. orders
    doesn't get silently blended into one wrong number."""
    start, end = dashboard_period_bounds(period, which, now=now)
    pos = list_pos_in_range(conn, start, end)
    return sum(float(p["base_total"] or 0) for p in pos)


def average_order_value(pos):
    """Batch 117: mean base_total (home-currency-normalized, same reasoning
    as dashboard_period_spend above) across a list of PO dicts -- backs the
    Dashboard's "Average order value" tile. Deliberately a plain pure
    function over whatever list the caller already has (DashboardPage._refresh
    passes its existing `active` list -- core.list_pos(conn), the same one
    the "Total purchase orders"/"Active (not deleted)" tiles already use)
    rather than querying again itself. Returns 0.0 for an empty list rather
    than raising a ZeroDivisionError."""
    if not pos:
        return 0.0
    return sum(float(p["base_total"] or 0) for p in pos) / len(pos)


def reports_due(conn, now=None):
    """Returns a list of {"freq", "period_start", "period_end"} dicts, one
    per periodic PO summary report frequency that's both enabled and
    actually due right now -- the multi-frequency replacement for the old
    single-schedule report_is_due (Batch 99). Weekly/monthly/quarterly/
    yearly are fully independent: any subset can be enabled at once
    (Settings > PO Report), each on its own day and its own
    po_report_last_sent_{freq} tracking, so e.g. a weekly reminder firing
    every Monday has no bearing on whether the monthly one is also due --
    both, either, or neither can come back due at the same time.

    Ordered weekly-first (REPORT_FREQUENCIES' own order), the most
    granular/likely reason a check was triggered.

    Always an empty list when the master "po_report_enabled" switch is
    off -- that one setting still silences every frequency at once,
    exactly like it always has."""
    now = now or datetime.now()
    if get_setting(conn, "po_report_enabled", "0") != "1":
        return []
    due = []
    for freq in REPORT_FREQUENCIES:
        if get_setting(conn, f"po_report_{freq}_enabled", "0") != "1":
            continue
        period_start, period_end, scheduled_dt = _period_bounds_for_freq(conn, freq, now=now)
        if now < scheduled_dt:
            continue
        last_sent = get_setting(conn, f"po_report_last_sent_{freq}", "")
        if last_sent:
            try:
                if datetime.fromisoformat(last_sent) >= scheduled_dt:
                    continue
            except ValueError:
                pass
        due.append({"freq": freq, "period_start": period_start, "period_end": period_end})
    return due


def report_pull_target(conn, now=None):
    """Which single frequency + period the Reports page's "Pull report
    now" quick-action should act on -- that button pulls ONE report at a
    time rather than showing a frequency picker (use "Generate historical
    report" for anything more specific). Prefers whichever enabled
    frequency is actually due right now (weekly-first, from reports_due's
    own ordering); falls back to whichever frequency is merely ENABLED
    but not yet due, so the button still does something sensible when
    clicked off-schedule; falls back to plain weekly bounds if nothing is
    enabled at all, matching this app's original always-weekly default
    from before multi-frequency reporting existed.

    Returns (freq, period_start, period_end)."""
    now = now or datetime.now()
    due = reports_due(conn, now=now)
    if due:
        d = due[0]
        return d["freq"], d["period_start"], d["period_end"]
    for freq in REPORT_FREQUENCIES:
        if get_setting(conn, f"po_report_{freq}_enabled", "0") == "1":
            period_start, period_end, _sched = _period_bounds_for_freq(conn, freq, now=now)
            return freq, period_start, period_end
    period_start, period_end, _sched = _period_bounds_for_freq(conn, "weekly", now=now)
    return "weekly", period_start, period_end


def _stock_recap_period_bounds(conn, now=None):
    """Weekly-only equivalent of _period_bounds(), for the stock team
    recap's own independent stock_recap_* schedule settings. Same
    Monday-anchored week logic as the periodic report's weekly branch, just
    never a monthly option -- a stock recap only ever makes sense weekly."""
    now = now or datetime.now()
    window = get_setting(conn, "stock_recap_window", "last")
    try:
        hh, mm = (get_setting(conn, "stock_recap_time", "09:00") or "09:00").split(":")
        hh, mm = int(hh), int(mm)
    except Exception:
        hh, mm = 9, 0
    try:
        target_dow = int(get_setting(conn, "stock_recap_day", "0") or 0)
    except ValueError:
        target_dow = 0
    days_since = (now.weekday() - target_dow) % 7
    scheduled_dt = (now - timedelta(days=days_since)).replace(hour=hh, minute=mm, second=0, microsecond=0)
    this_monday = (now - timedelta(days=now.weekday())).replace(hour=0, minute=0, second=0, microsecond=0)
    this_sunday_end = (this_monday + timedelta(days=6)).replace(hour=23, minute=59, second=59, microsecond=0)
    if window == "last":
        period_start = this_monday - timedelta(days=7)
        period_end = (this_monday - timedelta(seconds=1)).replace(hour=23, minute=59, second=59, microsecond=0)
    else:
        period_start = this_monday
        period_end = this_sunday_end
    return period_start, period_end, scheduled_dt


def stock_recap_is_due(conn, now=None):
    """Returns (due, period_start, period_end) -- a single-schedule check
    in the same shape/spirit as reports_due's per-item dicts, just for the
    independent stock_recap_* schedule (which only ever makes sense
    weekly, so there's no multi-frequency version of this one)."""
    now = now or datetime.now()
    if get_setting(conn, "stock_recap_enabled", "0") != "1":
        return False, None, None
    period_start, period_end, scheduled_dt = _stock_recap_period_bounds(conn, now)
    if now < scheduled_dt:
        return False, None, None
    last_sent = get_setting(conn, "stock_recap_last_sent_at", "")
    if last_sent:
        try:
            if datetime.fromisoformat(last_sent) >= scheduled_dt:
                return False, None, None
        except ValueError:
            pass
    return True, period_start, period_end


def _stock_request_scheduled_dt(conn, now=None):
    """The weekly stock-requests summary's own schedule (Settings > Email
    & Reports > "Stock team requests", Batch 146's stock_request_day/
    stock_request_time settings) -- same "most recent occurrence of the
    scheduled weekday/time, not in the future" shape as
    _stock_recap_period_bounds()'s own scheduled_dt, just without a
    separate content window to resolve alongside it (see
    stock_request_is_due()'s docstring for why this email never has one)."""
    now = now or datetime.now()
    try:
        hh, mm = (get_setting(conn, "stock_request_time", "09:30") or "09:30").split(":")
        hh, mm = int(hh), int(mm)
    except Exception:
        hh, mm = 9, 30
    try:
        target_dow = int(get_setting(conn, "stock_request_day", "0") or 0)
    except ValueError:
        target_dow = 0
    days_since = (now.weekday() - target_dow) % 7
    return (now - timedelta(days=days_since)).replace(hour=hh, minute=mm, second=0, microsecond=0)


def stock_request_is_due(conn, now=None):
    """Batch 150. Yitzi's original request (Task #95, seeded ahead of time
    by Batch 146's Settings card but never actually built until now): a
    weekly nudge to Supply listing every stock request that's still open,
    a belt-and-braces catch for anything the real-time notifications
    (stock_request_item/stock_request_new_item, both routed to Supply the
    moment Stock raises something) got missed on. Same due-check shape as
    stock_recap_is_due() above -- enabled, past its scheduled day/time,
    and not already sent since that scheduled moment -- but returns a
    plain bool rather than a (due, period_start, period_end) tuple: this
    email is always a live snapshot of whatever's genuinely open right
    now (there's no "requests raised in that particular week" grouping
    that would make a bounded period meaningful the way it is for the
    recap's actual placed orders), so there's no period for a caller to
    resolve before building it."""
    if get_setting(conn, "stock_request_enabled", "0") != "1":
        return False
    now = now or datetime.now()
    scheduled_dt = _stock_request_scheduled_dt(conn, now)
    if now < scheduled_dt:
        return False
    last_sent = get_setting(conn, "stock_request_last_sent_at", "")
    if last_sent:
        try:
            if datetime.fromisoformat(last_sent) >= scheduled_dt:
                return False
        except ValueError:
            pass
    return True


def list_pos_in_range(conn, date_from, date_to, deleted=False):
    rows = conn.execute(
        "SELECT * FROM purchase_orders WHERE deleted=? AND datetime(created_at) >= datetime(?) "
        "AND datetime(created_at) <= datetime(?) ORDER BY created_at",
        (1 if deleted else 0, date_from.isoformat(), date_to.isoformat()),
    ).fetchall()
    return [dict(r) for r in rows]


BRAND_CHART_COLORS = ["#2a206f", "#6498be", "#b079ad", "#c98a1a", "#1e9e6b", "#d64545"]


def _render_chart_png(kind, labels, values, title="", colors=None, figsize=(5, 3.2), currency_symbol=""):
    """Renders a simple bar or pie chart to PNG bytes using matplotlib's
    non-interactive Agg backend directly (no pyplot, so it doesn't touch
    whatever GUI backend the Qt app may already have active) in the app's
    brand colors, for embedding in the periodic report email or the
    supplier scorecard PDF. currency_symbol, if given, formats a bar
    chart's value axis as e.g. "£25,000" instead of a bare "25000"."""
    import io
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_agg import FigureCanvasAgg
    from matplotlib.ticker import FuncFormatter

    colors = colors or BRAND_CHART_COLORS
    fig = Figure(figsize=figsize, dpi=150)
    fig.patch.set_facecolor("#ffffff")
    ax = fig.add_subplot(111)

    if not labels:
        ax.text(0.5, 0.5, "No data for this period", ha="center", va="center", color="#6b7280")
        ax.axis("off")
    elif kind == "pie":
        # Labelling every wedge directly on the pie (the old approach) is
        # what caused "text in graphs should not overlap" -- several small
        # adjacent slices (a thin "Other"/"K6Tech" sliver next to another)
        # each get a name label plus a percentage label at almost the same
        # angle, and those text boxes collide. Fix: put the percentage
        # INSIDE each wedge only when the slice is big enough to hold it
        # legibly (small ones render no on-wedge text at all, so there's
        # nothing left to overlap), and move every name off the pie
        # entirely into a proper legend at the side -- which never
        # overlaps regardless of how many thin slices there are.
        wedge_colors = [colors[i % len(colors)] for i in range(len(labels))]
        total = sum(values) or 1

        def _autopct(pct):
            return f"{pct:.0f}%" if pct >= 6 else ""

        wedges, _texts, autotexts = ax.pie(
            values, colors=wedge_colors, autopct=_autopct,
            startangle=90, pctdistance=0.72, textprops={"fontsize": 8},
        )
        for t in autotexts:
            t.set_color("#ffffff")
            t.set_fontsize(8)
            t.set_fontweight("bold")
        ax.axis("equal")
        legend_labels = [f"{lbl} — {v / total * 100:.0f}%" for lbl, v in zip(labels, values)]
        ax.legend(
            wedges, legend_labels, loc="center left", bbox_to_anchor=(1.02, 0.5),
            fontsize=7.5, frameon=False, handlelength=1.1, labelspacing=0.6,
        )
        fig.subplots_adjust(left=0.02, right=0.60, top=0.9, bottom=0.05)
    elif kind == "line":
        x = range(len(labels))
        line_color = colors[0] if colors else BRAND_CHART_COLORS[0]
        ax.plot(list(x), values, color=line_color, marker="o", markersize=3, linewidth=1.5)
        ax.set_xticks(list(x))
        ax.set_xticklabels(labels, fontsize=8, rotation=25, ha="right")
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, pos: f"{currency_symbol}{v:,.0f}"))
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
    else:  # bar
        bar_colors = [colors[i % len(colors)] for i in range(len(labels))]
        x = range(len(labels))
        ax.bar(list(x), values, color=bar_colors)
        ax.set_xticks(list(x))
        ax.set_xticklabels(labels, fontsize=8, rotation=25, ha="right")
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, pos: f"{currency_symbol}{v:,.0f}"))
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)

    if title:
        ax.set_title(title, fontsize=10, color=RCG_INK, fontweight="bold")
    if kind != "pie":
        # The pie branch already positioned itself explicitly via
        # subplots_adjust (to reserve room for its side legend) --
        # tight_layout() would recompute and override that placement.
        fig.tight_layout()

    canvas = FigureCanvasAgg(fig)
    buf = io.BytesIO()
    canvas.print_png(buf)
    return buf.getvalue()


def _top_items_in_range(conn, pos):
    """Returns (top_5_by_cost, top_5_by_qty) -- each a list of
    {"product", "total_qty", "total_cost"} dicts -- aggregated across every
    line item on the given POs. total_cost is in the home currency (each
    line's price * its PO's fx_rate), so orders in different currencies
    don't get summed together in the wrong units."""
    if not pos:
        return [], []
    placeholders = ",".join("?" * len(pos))
    rows = conn.execute(
        f"SELECT i.product AS product, SUM(i.qty) AS total_qty, "
        f"SUM(i.qty * i.price * po.fx_rate) AS total_cost "
        f"FROM po_items i JOIN purchase_orders po ON po.id = i.po_id "
        f"WHERE i.po_id IN ({placeholders}) AND i.product != '' "
        f"GROUP BY i.product",
        [p["id"] for p in pos],
    ).fetchall()
    items = [dict(r) for r in rows]
    by_cost = sorted(items, key=lambda r: r["total_cost"], reverse=True)[:5]
    by_qty = sorted(items, key=lambda r: r["total_qty"], reverse=True)[:5]
    return by_cost, by_qty


def get_period_summary_stats(conn, date_from, date_to):
    """The core "how much happened" numbers for one period -- the shared
    source data behind every report period's key-stats section and the
    period-over-period comparison (change request section 5: "total spend,
    total savings, number of purchase orders, number of products/items
    ordered, total quantity ordered, supplier activity, product activity...
    comparison with the previous period"). date_from/date_to are
    start-of-day/end-of-day datetimes, same shape _period_bounds and the
    other bounds helpers above return."""
    pos = list_pos_in_range(conn, date_from, date_to)
    total_spend = sum(float(p["base_total"] or 0) for p in pos)
    po_count = len(pos)
    avg_po_value = (total_spend / po_count) if po_count else 0.0
    date_from_s, date_to_s = date_from.strftime("%Y-%m-%d"), date_to.strftime("%Y-%m-%d")
    qty_row = conn.execute(
        "SELECT COALESCE(SUM(i.qty), 0) AS total_qty, COUNT(DISTINCT i.product) AS distinct_products "
        "FROM po_items i JOIN purchase_orders po ON po.id = i.po_id "
        "WHERE po.deleted = 0 AND date(po.created_at) >= date(?) AND date(po.created_at) <= date(?)",
        (date_from_s, date_to_s),
    ).fetchone()
    total_qty = float(qty_row["total_qty"] or 0)
    distinct_products = int(qty_row["distinct_products"] or 0)
    savings_total = get_savings_summary(conn, date_from=date_from_s, date_to=date_to_s)["total"]
    top_suppliers = report_spend_by_supplier(conn, date_from=date_from_s, date_to=date_to_s)[:5]
    top_products = report_spend_by_product(conn, date_from=date_from_s, date_to=date_to_s)[:5]
    return {
        "po_count": po_count, "total_spend": total_spend, "total_qty": total_qty,
        "distinct_products": distinct_products, "avg_po_value": avg_po_value,
        "savings_total": savings_total, "top_suppliers": top_suppliers, "top_products": top_products,
        "pos": pos,
    }


# The five figures every report period's "how this compares to the
# previous X" section shows, in display order.
_COMPARISON_METRICS = ["total_spend", "savings_total", "po_count", "total_qty", "avg_po_value"]
_METRIC_LABELS = {
    "total_spend": "Total spend",
    "savings_total": "Total savings",
    "po_count": "Purchase orders placed",
    "total_qty": "Total quantity ordered",
    "avg_po_value": "Average PO value",
}
_MONEY_METRICS = {"total_spend", "savings_total", "avg_po_value"}
# The one figure the change request itself frames as unambiguously better
# when higher ("savings performance") -- every other metric here is shown
# with a neutral arrow/colour, per Yitzi's explicit instruction not to
# assume "up = good, down = bad" (spend going up might just mean the
# business was busier, not overspending).
_METRIC_HIGHER_IS_BETTER = {"savings_total"}


def _pct_change(new, old):
    """None (rendered as "n/a") when there's no previous-period figure to
    compare against at all, rather than a misleading 0% or a divide error."""
    if not old:
        return None
    return (new - old) / old * 100.0


def compare_periods(conn, period_kind, date_from, date_to):
    """Builds the "this period vs. the equivalent previous period"
    comparison every report period now includes (change request: "each
    report should compare its selected period with the equivalent previous
    period"). period_kind: "week"/"month"/"quarter"/"year". Purely data --
    no colour or good/bad framing baked in here, that's a rendering
    decision (_render_comparison_section) kept deliberately separate so
    the underlying numbers stay reusable wherever else they might be
    wanted."""
    prev_start, prev_end = previous_period_bounds(period_kind, date_from, date_to)
    current = get_period_summary_stats(conn, date_from, date_to)
    previous = get_period_summary_stats(conn, prev_start, prev_end)
    deltas = {}
    for key in _COMPARISON_METRICS:
        cur_v, prev_v = current[key], previous[key]
        deltas[key] = {"current": cur_v, "previous": prev_v, "pct_change": _pct_change(cur_v, prev_v)}
    return {
        "period_kind": period_kind, "current": current, "previous": previous,
        "previous_start": prev_start, "previous_end": prev_end, "deltas": deltas,
    }


def _format_metric_value(key, value, currency):
    return money(value, currency) if key in _MONEY_METRICS else f"{value:g}"


def _render_comparison_section(comparison, currency):
    """Shared HTML+plain "how this compares to the previous period" block --
    used by both the weekly report (layered onto build_po_period_email,
    which otherwise keeps its existing full per-order breakdown untouched)
    and the monthly/quarterly/yearly summary report
    (build_period_summary_email). Returns ("", "") if comparison is falsy,
    so every caller can pass it through unconditionally.

    A flat/negligible change (within +/-0.05%) reads as "steady" rather
    than manufacturing a misleading tiny arrow out of rounding noise."""
    if not comparison:
        return "", ""
    # period_kind is already the singular noun ("week"/"month"/"quarter"/
    # "year") -- _PERIOD_NOUN is keyed the other way round (by the
    # "weekly"/"monthly"/... frequency setting, for _period_bounds' own
    # use above), so looking period_kind up in it would always miss and
    # silently fall back to the generic word "period".
    noun = comparison.get("period_kind") or "period"
    header_html = (
        '<tr>'
        f'<td style="padding:2px 0 6px 0;"></td>'
        f'<td align="right" style="padding:2px 0 6px 0;font-size:10px;letter-spacing:0.04em;'
        f'text-transform:uppercase;color:{RCG_MUTED};border-bottom:1px solid {RCG_LINE};">'
        f'This {html_escape(noun)}</td>'
        f'<td align="right" style="padding:2px 0 6px 0;font-size:10px;letter-spacing:0.04em;'
        f'text-transform:uppercase;color:{RCG_MUTED};border-bottom:1px solid {RCG_LINE};">'
        f'Last {html_escape(noun)}</td>'
        f'<td style="padding:2px 0 6px 0;border-bottom:1px solid {RCG_LINE};"></td>'
        '</tr>'
    )
    rows_html, rows_plain = [header_html], []
    for key in _COMPARISON_METRICS:
        d = comparison["deltas"][key]
        label = _METRIC_LABELS.get(key, key)
        cur_txt = _format_metric_value(key, d["current"], currency)
        prev_txt = _format_metric_value(key, d["previous"], currency)
        pct = d["pct_change"]
        if pct is None:
            change_txt, colour, tint = "n/a (no data for the previous period)", RCG_MUTED, "#f1f0f5"
        else:
            arrow = "▲" if pct > 0.05 else ("▼" if pct < -0.05 else "▶")
            change_txt = f"{arrow} {abs(pct):.0f}%" if abs(pct) > 0.05 else f"{arrow} steady"
            if key in _METRIC_HIGHER_IS_BETTER:
                if pct > 0.05:
                    colour, tint = RCG_GREEN, "#e3f5eb"
                elif pct < -0.05:
                    colour, tint = "#c98a1a", "#fbf1de"
                else:
                    colour, tint = RCG_MUTED, "#f1f0f5"
            else:
                # Neutral on purpose -- a change here isn't automatically
                # good or bad (see the note under the table).
                colour, tint = "#3a3164", "#f1f0f5"
        rows_html.append(
            f'<tr><td style="padding:4px 0;font-size:13px;color:#3a3164;">{html_escape(label)}</td>'
            f'<td align="right" style="padding:4px 0;font-size:13px;color:#3a3164;">{cur_txt}</td>'
            f'<td align="right" style="padding:4px 0;font-size:13px;color:{RCG_MUTED};">{prev_txt}</td>'
            f'<td align="right" style="padding:4px 0;"><span style="background:{tint};color:{colour};'
            f'font-size:12px;font-weight:bold;padding:2px 8px;border-radius:10px;">{change_txt}</span></td></tr>'
        )
        rows_plain.append(f"- {label}: {cur_txt} (previous {noun}: {prev_txt}) {change_txt}")
    html = (
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" '
        'style="margin-top:20px;"><tr><td style="border-left:3px solid '
        f'{RCG_ACCENT_LIGHT};background:#f9f8fc;padding:14px 18px;">'
        f'<div style="font-size:13px;font-weight:bold;color:{RCG_INK};margin-bottom:8px;">'
        f'How this compares to the previous {html_escape(noun)}</div>'
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">'
        + "".join(rows_html) + "</table>"
        f'<div style="font-size:11.5px;color:{RCG_MUTED};margin-top:8px;">An increase or decrease here '
        "isn't automatically good or bad on its own &mdash; for example, spend may simply be higher "
        "because more was ordered.</div></td></tr></table>"
    )
    plain = (
        f"How this compares to the previous {noun}:\n" + "\n".join(rows_plain) +
        "\n(An increase or decrease isn't automatically good or bad on its own -- e.g. spend may simply "
        "be higher because more was ordered.)\n\n"
    )
    return html, plain


def _render_exceptions_section(conn, date_from, date_to, currency, issues):
    """The "Exceptions & notable decisions" block, shared by the weekly and
    monthly/quarterly/yearly reports, per Yitzi's feedback that reporting
    should "highlight exceptions/actions, not only the numbers, i.e.
    anything that went wrong, needs attention, or where you made a notable
    procurement decision." Three parts, each pulled from what's already
    tracked in the app rather than relying on anyone retyping it into the
    report by hand: supplier issues logged in the period (what went
    wrong), price increases recorded -- a savings entry with a negative
    amount (what needs attention), and savings or other deliberate
    procurement calls -- a positive amount (the notable decisions), each
    shown with saving_basis_text so the figure is never just a bare
    number. issues is the already-fetched list_supplier_issues_in_range
    result -- the weekly report needs it separately too, for its own
    return dict, so it's passed in rather than queried twice here. Returns
    (html, plain)."""
    date_from_s = date_from.strftime("%Y-%m-%d") if hasattr(date_from, "strftime") else date_from
    date_to_s = date_to.strftime("%Y-%m-%d") if hasattr(date_to, "strftime") else date_to
    period_savings = list_savings(conn, date_from=date_from_s, date_to=date_to_s)
    increases = [s for s in period_savings if float(s["amount"] or 0) < 0]
    decisions = [s for s in period_savings if float(s["amount"] or 0) > 0]

    issues_rows = [
        [html_escape((i.get('at') or '')[:10]), html_escape(i['supplier_company_name']),
         f'<span style="background:{RCG_ROW_MAUVE};color:{RCG_MAUVE};font-size:11px;font-weight:bold;'
         f'padding:2px 8px;border-radius:10px;">{html_escape(i.get("category") or "-")}</span>',
         html_escape(i.get('subject') or i.get('note', ''))]
        for i in issues
    ]
    increase_rows = [
        [html_escape((s.get('created_at') or '')[:10]),
         html_escape(s.get('product') or '(no product)'),
         html_escape(s.get('supplier_company_name') or ''),
         f'<span style="color:{RCG_RED};font-weight:bold;">{html_escape(money(abs(float(s["amount"])), currency))}</span>',
         html_escape(saving_basis_text(s, currency))]
        for s in increases
    ]
    decision_rows = [
        [html_escape((s.get('created_at') or '')[:10]),
         html_escape(s.get('product') or '(no product)'),
         html_escape(s.get('supplier_company_name') or ''),
         html_escape(s.get('category') or ''),
         f'<span style="color:{RCG_GREEN};font-weight:bold;">{html_escape(money(float(s["amount"]), currency))}</span>',
         html_escape(saving_basis_text(s, currency))]
        for s in decisions
    ]

    html = (
        email_section_heading_html("Exceptions & notable decisions this period")
        + f'<div style="font-size:11.5px;color:{RCG_MUTED};margin:-4px 0 10px 0;">'
          'Anything that went wrong or needs attention, alongside any deliberate procurement calls made '
          'this period.</div>'
        + f'<div style="font-size:12.5px;color:{RCG_INK};font-weight:bold;margin-top:10px;">'
          f'Supplier issues logged ({len(issues)})</div>'
        + email_table_html(["Date", "Supplier", "Category", "Subject"], issues_rows)
        + f'<div style="font-size:12.5px;color:{RCG_INK};font-weight:bold;margin-top:16px;">'
          f'Price increases to note ({len(increases)})</div>'
        + email_table_html(["Date", "Product", "Supplier", "Increase", "Basis"], increase_rows)
        + f'<div style="font-size:12.5px;color:{RCG_INK};font-weight:bold;margin-top:16px;">'
          f'Savings &amp; notable procurement decisions ({len(decisions)})</div>'
        + email_table_html(["Date", "Product", "Supplier", "Category", "Saving", "Basis"], decision_rows)
    )

    issues_lines = "".join(
        f"- {(i.get('at') or '')[:10]}  {i['supplier_company_name']}  [{i.get('category') or '-'}]  "
        f"{i.get('subject') or i.get('note', '')}\n" for i in issues
    ) or "- None\n"
    increase_lines = "".join(
        f"- {(s.get('created_at') or '')[:10]}  {s.get('product') or '(no product)'}  "
        f"{s.get('supplier_company_name') or ''}  +{money(abs(float(s['amount'])), currency)}  "
        f"({saving_basis_text(s, currency)})\n" for s in increases
    ) or "- None\n"
    decision_lines = "".join(
        f"- {(s.get('created_at') or '')[:10]}  {s.get('product') or '(no product)'}  "
        f"{s.get('supplier_company_name') or ''}  [{s.get('category') or ''}]  "
        f"{money(float(s['amount']), currency)}  ({saving_basis_text(s, currency)})\n" for s in decisions
    ) or "- None\n"
    plain = (
        "Exceptions & notable decisions this period:\n"
        "(Anything that went wrong or needs attention, alongside any deliberate procurement calls made "
        "this period.)\n"
        f"Supplier issues logged ({len(issues)}):\n{issues_lines}\n"
        f"Price increases to note ({len(increases)}):\n{increase_lines}\n"
        f"Savings & notable procurement decisions ({len(decisions)}):\n{decision_lines}\n"
    )
    return html, plain


def build_po_period_email(conn, date_from, date_to, label="This period's", comparison=None):
    """Builds the subject/html/plain body for a PO summary covering
    date_from..date_to: order count and total, a breakdown by supplier, the
    top 5 items bought (by cost and by quantity), a spend-by-supplier pie
    chart and a top-items bar chart (both in brand colors), a plain list of
    the orders themselves, and an "Exceptions & notable decisions" section
    (see _render_exceptions_section) pulling together every supplier issue,
    price increase, and saving actually logged in the app during this same
    period so they don't have to be retyped here by hand. A short blank
    prompt is still included for anything that isn't tracked in this app
    at all (RMAs, shortages/returns, etc.)."""
    pos = list_pos_in_range(conn, date_from, date_to)
    # "currency" here is the home currency -- used for every cross-PO total
    # (base_total), which is always in that currency regardless of what any
    # individual PO was raised in. Each PO's own row still shows its own
    # currency below, since that's the amount actually ordered in it.
    currency = get_setting(conn, "currency", "GBP")
    total = sum(float(p["base_total"] or 0) for p in pos)

    by_supplier = {}
    for p in pos:
        name = p["supplier_company_name"] or "(no supplier)"
        entry = by_supplier.setdefault(name, {"count": 0, "total": 0.0})
        entry["count"] += 1
        entry["total"] += float(p["base_total"] or 0)
    top_suppliers = sorted(by_supplier.items(), key=lambda kv: kv[1]["total"], reverse=True)

    top_items_cost, top_items_qty = _top_items_in_range(conn, pos)
    issues = list_supplier_issues_in_range(conn, date_from, date_to)

    rows_html_table = [
        [html_escape(p['po_ref']), html_escape(p['supplier_company_name'] or ''),
         html_escape(money(p['total'], p['currency'] or currency))]
        for p in pos
    ]
    suppliers_html = "".join(
        f'<tr><td style="padding:4px 0;color:#3a3164;">&#9679;&nbsp; {html_escape(name)}</td>'
        f'<td align="right" style="padding:4px 0;color:#3a3164;">'
        f'{pluralize(info["count"], "order")}, {money(info["total"], currency)}</td></tr>'
        for name, info in top_suppliers
    ) or f'<tr><td style="padding:4px 0;color:{RCG_MUTED};">None</td></tr>'

    items_cost_html = "".join(
        f'<div style="padding:3px 0;color:#3a3164;">{html_escape(i["product"])} &mdash; '
        f'{money(i["total_cost"], currency)}</div>' for i in top_items_cost
    ) or f'<div style="color:{RCG_MUTED};">None</div>'
    items_qty_html = "".join(
        f'<div style="padding:3px 0;color:#3a3164;">{html_escape(i["product"])} &mdash; '
        f'{i["total_qty"]:g} {plural_word(i["total_qty"], "unit")}</div>' for i in top_items_qty
    ) or f'<div style="color:{RCG_MUTED};">None</div>'
    exceptions_html, exceptions_plain = _render_exceptions_section(conn, date_from, date_to, currency, issues)

    # --- charts, in brand colors (navy / light blue / mauve / amber / green / red) ---
    pie_cid, bar_cid = "supplier_pie_chart", "items_bar_chart"
    top5_suppliers = top_suppliers[:5]
    other_total = sum(info["total"] for _n, info in top_suppliers[5:])
    pie_labels = [n for n, _i in top5_suppliers] + (["Other"] if other_total > 0 else [])
    pie_values = [info["total"] for _n, info in top5_suppliers] + ([other_total] if other_total > 0 else [])
    supplier_pie_png = _render_chart_png("pie", pie_labels, pie_values, title="Spend by supplier")

    bar_labels = [(i["product"][:18] + "…") if len(i["product"]) > 18 else i["product"] for i in top_items_qty]
    bar_values = [i["total_qty"] for i in top_items_qty]
    items_bar_png = _render_chart_png("bar", bar_labels, bar_values, title="Top items by quantity")

    charts_html = (
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" '
        'style="margin-top:16px;"><tr>'
        f'<td width="50%" align="center"><img src="cid:{pie_cid}" alt="Spend by supplier" '
        'style="max-width:100%; height:auto;"></td>'
        f'<td width="50%" align="center"><img src="cid:{bar_cid}" alt="Top items by quantity" '
        'style="max-width:100%; height:auto;"></td>'
        '</tr></table>'
    )

    recipient_name = get_setting(conn, "po_report_recipient_name", "") if conn is not None else ""
    greeting = f"Hi {recipient_name}," if recipient_name else "Hi,"
    comparison_html, comparison_plain = _render_comparison_section(comparison, currency)

    body_html = (
        f'<p style="margin:20px 0 4px 0;font-size:15px;color:{RCG_INK};">{html_escape(greeting)}</p>'
        f'<h2 style="margin:0 0 6px 0;font-size:22px;line-height:1.3;color:{RCG_INK};font-weight:bold;">'
        f'{html_escape(label)} purchase order summary</h2>'
        f'<p style="margin:0;font-size:14px;color:{RCG_MUTED};">'
        f'{date_from.strftime("%d %b %Y")} to {date_to.strftime("%d %b %Y")}</p>'
        + email_stat_strip_html([
            (str(len(pos)), plural_word(len(pos), "order") + " placed"),
            (money(total, currency), "total spend"),
        ])
        + comparison_html
        + email_section_heading_html("Main suppliers")
        + f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" '
          f'style="font-size:13.5px;">{suppliers_html}</table>'
        + charts_html
        + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" '
          'style="margin-top:12px;"><tr>'
        + f'<td width="50%" valign="top">{email_section_heading_html("Top 5 items by cost")}{items_cost_html}</td>'
        + f'<td width="50%" valign="top">{email_section_heading_html("Top 5 items by quantity")}{items_qty_html}</td>'
        + '</tr></table>'
        + email_section_heading_html("Orders placed")
        + email_table_html(["PO Ref", "Supplier", "Total"], rows_html_table, align=["left", "left", "right"])
        + exceptions_html
        + email_section_heading_html("Anything else to flag")
        + f'<p style="margin:0 0 20px 0;font-size:12.5px;color:{RCG_MUTED};font-style:italic;">'
          'RMAs, shortages/returns, or anything else worth noting that isn\'t already tracked above '
          '&mdash; add your own notes here before sending.</p>'
    )
    html = email_shell_html(body_html)

    supplier_lines = "".join(
        f"- {name}: {pluralize(info['count'], 'order')}, {money(info['total'], currency)}\n" for name, info in top_suppliers
    ) or "- None\n"
    items_cost_lines = "".join(
        f"- {i['product']}: {money(i['total_cost'], currency)}\n" for i in top_items_cost
    ) or "- None\n"
    items_qty_lines = "".join(
        f"- {i['product']}: {i['total_qty']:g}\n" for i in top_items_qty
    ) or "- None\n"
    po_lines = "".join(
        f"- {p['po_ref']}  {p['supplier_company_name'] or ''}  {money(p['total'], p['currency'] or currency)}\n"
        for p in pos
    ) or "- No orders placed in this period.\n"
    plain = (
        f"{greeting}\n\n"
        f"{label} purchase order summary\n{date_from:%d %b %Y} to {date_to:%d %b %Y}\n\n"
        f"{pluralize(len(pos), 'order')} placed, total {money(total, currency)}.\n\n"
        f"{comparison_plain}"
        f"Main suppliers:\n{supplier_lines}\n"
        f"Top 5 items by cost:\n{items_cost_lines}\n"
        f"Top 5 items by quantity:\n{items_qty_lines}\n"
        f"Orders placed:\n{po_lines}\n"
        f"{exceptions_plain}"
        f"Anything else to flag:\n(RMAs, shortages/returns, or anything else worth noting that isn't "
        f"already tracked above -- add your own notes here before sending.)\n"
    )

    subject = f"{label} PO Summary — {date_from:%d %b} to {date_to:%d %b %Y}"
    charts = {pie_cid: supplier_pie_png, bar_cid: items_bar_png}
    logo_png = rcg_logo_png_bytes()
    if logo_png:
        charts[RCG_LOGO_CID] = logo_png
    return {
        "subject": subject, "html": html, "plain": plain, "po_count": len(pos), "total": total, "pos": pos,
        "top_items_cost": top_items_cost, "top_items_qty": top_items_qty, "issues": issues,
        "charts": charts, "comparison": comparison,
    }


def build_period_summary_email(conn, period_kind, date_from, date_to, label=None):
    """The Monthly/Quarterly/Yearly report -- a strong OVERVIEW of the
    period rather than the weekly report's full per-order breakdown, per
    the change request: "the monthly, quarterly and yearly reports do not
    need to repeat every individual order in the same detailed breakdown
    format as the weekly report. Instead, they should provide a strong
    overview of the selected period, including: period totals, key
    statistics, comparisons with the previous period, useful summaries and
    trends." Always includes the same "how this compares to the previous
    {period}" section the weekly report also gets (_render_comparison_section)
    plus a spend-trend chart bucketed at a sensible granularity for the
    period's length (weekly buckets across a month, monthly buckets across
    a quarter or year), the period's top suppliers/products, and the same
    "Exceptions & notable decisions" section the weekly report gets
    (_render_exceptions_section) -- previously only weekly surfaced
    supplier issues at all, so anything logged in a month/quarter/year
    that only gets its own periodic report went unseen until now."""
    currency = get_setting(conn, "currency", "GBP")
    comparison = compare_periods(conn, period_kind, date_from, date_to)
    current = comparison["current"]
    # Same fix as _render_comparison_section above -- period_kind ("week"/
    # "month"/"quarter"/"year") is already the noun, not a _PERIOD_NOUN key.
    noun = period_kind or "period"
    label = label or f"This {noun}'s"

    date_from_s, date_to_s = date_from.strftime("%Y-%m-%d"), date_to.strftime("%Y-%m-%d")
    trend_granularity = {"month": "week", "quarter": "month", "year": "month"}.get(period_kind, "month")
    trend = report_spend_over_time(conn, date_from=date_from_s, date_to=date_to_s, granularity=trend_granularity)
    trend_cid = "period_trend_chart"
    trend_png = _render_chart_png(
        "line", [t["period"] for t in trend], [t["total"] for t in trend],
        title=f"Spend trend this {noun}", currency_symbol=CURRENCY_SYMBOLS.get(currency, ""),
    )
    comparison_html, comparison_plain = _render_comparison_section(comparison, currency)

    # Real shaded tables (matching every other data table in the app) in
    # place of the old free-floating bullet list / plain divs -- those had
    # no gutter between the two 50%-wide columns and no row alignment, so
    # a long supplier total or product name would wrap straight into the
    # other column ("in the monthly report the email needs to be better
    # formatted").
    suppliers_rows = [
        [html_escape(s["supplier"] or "(no supplier)"),
         f'{pluralize(s["po_count"], "order")}, {html_escape(money(s["total"], currency))}']
        for s in current["top_suppliers"]
    ]
    suppliers_html = email_table_html(["Supplier", "Orders"], suppliers_rows, align=["left", "right"])
    products_rows = [
        [html_escape(p["product"]),
         f'{html_escape(money(p["total_spend"], currency))} ({p["total_qty"]:g} {plural_word(p["total_qty"], "unit")})']
        for p in current["top_products"]
    ]
    products_html = email_table_html(["Product", "Spend"], products_rows, align=["left", "right"])

    issues = list_supplier_issues_in_range(conn, date_from, date_to)
    exceptions_html, exceptions_plain = _render_exceptions_section(conn, date_from, date_to, currency, issues)

    recipient_name = get_setting(conn, "po_report_recipient_name", "") if conn is not None else ""
    greeting = f"Hi {recipient_name}," if recipient_name else "Hi,"

    body_html = (
        f'<p style="margin:20px 0 4px 0;font-size:15px;color:{RCG_INK};">{html_escape(greeting)}</p>'
        f'<h2 style="margin:0 0 6px 0;font-size:22px;line-height:1.3;color:{RCG_INK};font-weight:bold;">'
        f'{html_escape(label)} purchase order summary</h2>'
        f'<p style="margin:0;font-size:14px;color:{RCG_MUTED};">'
        f'{date_from.strftime("%d %b %Y")} to {date_to.strftime("%d %b %Y")}</p>'
        + email_stat_strip_html([
            (str(current["po_count"]), plural_word(current["po_count"], "order") + " placed"),
            (money(current["total_spend"], currency), "total spend"),
            (money(current["avg_po_value"], currency), "average PO value"),
            (f'{current["total_qty"]:g}', plural_word(current["total_qty"], "unit") + " ordered"),
            (money(current["savings_total"], currency), "savings recorded"),
        ])
        + comparison_html
        + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" '
          'style="margin-top:16px;"><tr>'
        + f'<td align="center"><img src="cid:{trend_cid}" alt="Spend trend" '
          'style="max-width:100%; height:auto;"></td></tr></table>'
        + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" '
          'style="margin-top:12px;"><tr>'
        + '<td width="49%" valign="top">' + email_section_heading_html("Top suppliers") + suppliers_html + '</td>'
        + '<td width="2%" style="font-size:0;line-height:0;">&nbsp;</td>'
        + '<td width="49%" valign="top">' + email_section_heading_html("Top products") + products_html + '</td>'
        + '</tr></table>'
        + exceptions_html
    )
    supplier_lines = "".join(
        f"- {s['supplier'] or '(no supplier)'}: {pluralize(s['po_count'], 'order')}, {money(s['total'], currency)}\n"
        for s in current["top_suppliers"]
    ) or "- None\n"
    product_lines = "".join(
        f"- {p['product']}: {money(p['total_spend'], currency)} "
        f"({p['total_qty']:g} {plural_word(p['total_qty'], 'unit')})\n"
        for p in current["top_products"]
    ) or "- None\n"
    plain = (
        f"{greeting}\n\n"
        f"{label} purchase order summary\n{date_from:%d %b %Y} to {date_to:%d %b %Y}\n\n"
        f"{pluralize(current['po_count'], 'order')} placed, total {money(current['total_spend'], currency)}, "
        f"average PO value {money(current['avg_po_value'], currency)}, "
        f"{current['total_qty']:g} {plural_word(current['total_qty'], 'unit')} ordered, "
        f"savings recorded {money(current['savings_total'], currency)}.\n\n"
        f"{comparison_plain}"
        f"Top suppliers:\n{supplier_lines}\n"
        f"Top products:\n{product_lines}\n\n"
        f"{exceptions_plain}"
    )
    html = email_shell_html(body_html)
    subject = f"{label} PO Summary — {date_from:%d %b} to {date_to:%d %b %Y}"
    charts = {trend_cid: trend_png}
    logo_png = rcg_logo_png_bytes()
    if logo_png:
        charts[RCG_LOGO_CID] = logo_png
    return {
        "subject": subject, "html": html, "plain": plain,
        "po_count": current["po_count"], "total": current["total_spend"],
        "comparison": comparison, "charts": charts, "issues": issues,
    }


def build_period_report_email(conn, period_kind, date_from, date_to, label=None):
    """Single entry point for all four report periods (change request
    section 5) -- dispatches to the right builder: weekly keeps its full
    per-order breakdown (build_po_period_email, unchanged) with the new
    comparison section layered on top; monthly/quarterly/yearly get the
    lighter overview-style build_period_summary_email instead. Both
    return the same dict shape (subject/html/plain/po_count/total/
    comparison/charts), so callers (open_po_report_email, the manual
    historical-report dialog) don't need to know which one ran."""
    if period_kind == "week":
        comparison = compare_periods(conn, "week", date_from, date_to)
        return build_po_period_email(conn, date_from, date_to, label=label or "This week's", comparison=comparison)
    return build_period_summary_email(conn, period_kind, date_from, date_to, label=label)


def mark_report_sent(conn, freq, when=None):
    """Records that the periodic PO summary report was just sent/pulled
    for one specific frequency -- each of weekly/monthly/quarterly/yearly
    tracks its own po_report_last_sent_{freq} independently (Batch 99), so
    marking (say) monthly as sent has no effect on whether weekly is still
    due, and vice versa."""
    freq = normalize_report_frequency(freq)
    set_settings(conn, {f"po_report_last_sent_{freq}": (when or datetime.now()).isoformat(timespec="seconds")})


def _format_recap_timestamp(iso_str):
    """'2026-08-28T09:53:40' -> '28 Aug 2026, 09:53', tolerating whatever
    partial/odd string a hand-edited settings value might contain -- falls
    back to returning the raw string rather than raising."""
    if not iso_str:
        return ""
    try:
        return datetime.fromisoformat(iso_str).strftime("%d %b %Y, %H:%M")
    except ValueError:
        return iso_str


_STOCK_RECAP_STATUS_COLOR_HEX = {"green": RCG_GREEN, "red": RCG_RED, "mauve": RCG_MAUVE, "muted": RCG_MUTED}


def _stock_recap_status_line(po_id, recap_enrichment):
    """Batch 98: the real receiving status line for one PO in the weekly
    recap, or None if recap_enrichment wasn't supplied (stock sync off --
    see build_stock_recap_email's own docstring for why that has to leave
    the email completely unchanged). Returns (text, color_key), where
    color_key is one of "green"/"red"/"mauve"/"muted" -- a shared vocabulary
    both the HTML email (via _STOCK_RECAP_STATUS_COLOR_HEX) and the PDF
    (via the branded canvas's own c["colors"] dict, same keys) understand,
    so this one function serves both renderers without either needing to
    know how the other represents color."""
    if recap_enrichment is None:
        return None
    status = recap_enrichment.get("receiving_by_po_id", {}).get(po_id)
    summary = (status or {}).get("summary", "")
    when = _format_recap_timestamp((status or {}).get("last_updated_at"))
    if summary == "Received":
        return (f"Received{f'  ({when})' if when else ''}", "green")
    if summary == "Returned":
        return (f"Returned{f'  ({when})' if when else ''}", "red")
    if summary == "Partially received":
        return (f"Partially received{f'  — last update {when}' if when else ''}", "mauve")
    return ("Not received yet", "muted")


def build_stock_recap_email(conn, date_from, date_to, label="Last week's", recap_enrichment=None):
    """Builds the subject/html/plain body for the weekly stock team recap:
    every order placed in date_from..date_to with its line items, followed
    by a standing prompt asking the stock team to flag anything missing,
    not arrived yet, or any supplier issue/return/RMA.

    This app has no tracked delivery/ETA date per order (nothing captures
    when an item is actually expected), so this deliberately doesn't claim
    a specific arrival date for anything -- it just lists what was ordered
    in the period and leans on the normal couple-of-days turnaround plus
    the prompt at the end to surface anything that's actually outstanding.

    Batch 98, recap_enrichment (optional): once the stock team is actually
    using the companion app to mark things received, the recap can show
    real status instead of that generic line. This function has zero
    dependency on stock_sync.py (same as everywhere else in po_core.py --
    the dependency only ever runs the other way), so the caller
    (po_generator_qt.py) checks stock_sync.is_enabled() itself and builds
    this dict via stock_sync.get_stock_recap_enrichment() only when it's
    genuinely on, passing it in here. recap_enrichment=None (the default,
    and always the case when stock sync is off) means: render exactly the
    way this always has, generic line and all -- Yitzi was explicit that
    turning the stock plugin off must leave the recap untouched, not show
    blank/missing status for a stock team that was never using it."""
    pos = list_pos_in_range(conn, date_from, date_to)
    currency = get_setting(conn, "currency", "GBP")
    recipient_name = get_setting(conn, "stock_recap_recipient_name", "")
    greeting = f"Hi {recipient_name}," if recipient_name else "Hi,"
    prompt = get_setting(conn, "stock_recap_prompt", DEFAULT_SETTINGS["stock_recap_prompt"])

    # One query for every line item across all these POs, same batch
    # pattern used by the Zoho export (see build_zoho_export_rows) rather
    # than one extra query per PO.
    items_by_po = {}
    if pos:
        placeholders = ",".join("?" * len(pos))
        for r in conn.execute(
            f"SELECT * FROM po_items WHERE po_id IN ({placeholders}) ORDER BY po_id, position",
            [p["id"] for p in pos],
        ):
            items_by_po.setdefault(r["po_id"], []).append(dict(r))

    def render_po_block(p, items):
        po_currency = p.get("currency") or currency
        item_rows = [
            [html_escape(str(i.get('qty', ''))), html_escape(display_product_label(i)),
             html_escape(money(i.get('price'), po_currency)),
             html_escape(money(float(i.get('qty') or 0) * float(i.get('price') or 0), po_currency))]
            for i in items
        ]
        items_table = email_table_html(
            ["Qty", "Item", "Unit price", "Line total"], item_rows,
            align=["left", "left", "right", "right"],
        )
        status_line = _stock_recap_status_line(p["id"], recap_enrichment)
        status_html = (
            f'<span style="color:{_STOCK_RECAP_STATUS_COLOR_HEX[status_line[1]]};font-weight:600;"> '
            f'&mdash; {html_escape(status_line[0])}</span>'
            if status_line else ""
        )
        status_plain = f" -- {status_line[0]}" if status_line else ""
        html_block = (
            f'<div style="margin-top:18px;padding:12px 14px;background:{RCG_ROW_BLUE};">'
            f'<span style="font-weight:bold;color:{RCG_INK};">{html_escape(p["po_ref"])}</span>'
            f'<span style="color:{RCG_MUTED};"> &mdash; '
            f'{html_escape(p.get("supplier_company_name") or "(no supplier)")} '
            f'({html_escape((p.get("created_at") or "")[:10])})</span>{status_html}</div>'
            + items_table
            + f'<div style="text-align:right;padding:4px 10px 0 0;font-size:13px;color:{RCG_INK};">'
              f'<b>Order total: {html_escape(money(p.get("total"), po_currency))}</b></div>'
        )
        item_lines = "".join(
            f"  - {i.get('qty', '')}x {display_product_label(i)} "
            f"@ {money(i.get('price'), po_currency)} = "
            f"{money(float(i.get('qty') or 0) * float(i.get('price') or 0), po_currency)}\n"
            for i in items
        ) or "  - No items\n"
        plain_block = (
            f"{p['po_ref']} - {p.get('supplier_company_name') or '(no supplier)'} "
            f"({(p.get('created_at') or '')[:10]}){status_plain}\n{item_lines}"
            f"  Order total: {money(p.get('total'), po_currency)}\n"
        )
        return html_block, plain_block

    po_blocks_html, po_lines_plain = [], []
    for p in pos:
        html_block, plain_block = render_po_block(p, items_by_po.get(p["id"], []))
        po_blocks_html.append(html_block)
        po_lines_plain.append(plain_block)

    orders_html = "".join(po_blocks_html) or (
        f'<p style="color:{RCG_MUTED};">No orders were placed in this period.</p>'
    )
    orders_plain = "\n".join(po_lines_plain) or "No orders were placed in this period.\n"

    # Batch 98: orders from BEFORE this period that still aren't fully
    # received -- so the stock team (and Yitzi, reading the same email)
    # never loses track of something just because a new week started.
    # Only ever appears when recap_enrichment is supplied.
    outstanding_html, outstanding_plain, outstanding_count = "", "", 0
    outstanding_pos = (recap_enrichment or {}).get("outstanding_previous_pos", [])
    if outstanding_pos:
        outstanding_count = len(outstanding_pos)
        blocks_html, blocks_plain = [], []
        for p in outstanding_pos:
            html_block, plain_block = render_po_block(p, p.get("items", []))
            blocks_html.append(html_block)
            blocks_plain.append(plain_block)
        verb = "isn't" if outstanding_count == 1 else "aren't"
        outstanding_html = (
            f'<h3 style="margin:26px 0 4px 0;font-size:16px;color:{RCG_INK};">'
            f'Still outstanding from before this week</h3>'
            f'<p style="margin:0;font-size:13px;color:{RCG_MUTED};">'
            f'{pluralize(outstanding_count, "order")} placed earlier that {verb} '
            f'fully received yet.</p>' + "".join(blocks_html)
        )
        outstanding_plain = (
            f"\nStill outstanding from before this week "
            f"({pluralize(outstanding_count, 'order')}):\n\n" + "\n".join(blocks_plain)
        )

    if recap_enrichment is not None:
        summary_line_html = f'{pluralize(len(pos), "order")} placed this week.'
        summary_line_plain = f"{pluralize(len(pos), 'order')} placed this week."
    else:
        summary_line_html = f'{pluralize(len(pos), "order")} placed &mdash; most of this should have arrived by now.'
        summary_line_plain = f"{pluralize(len(pos), 'order')} placed -- most of this should have arrived by now."

    body_html = (
        f'<p style="margin:20px 0 4px 0;font-size:15px;color:{RCG_INK};">{html_escape(greeting)}</p>'
        f'<h2 style="margin:0 0 6px 0;font-size:22px;line-height:1.3;color:{RCG_INK};font-weight:bold;">'
        f'{html_escape(label)} stock recap</h2>'
        f'<p style="margin:0;font-size:14px;color:{RCG_MUTED};">'
        f'{date_from.strftime("%d %b %Y")} to {date_to.strftime("%d %b %Y")}</p>'
        f'<p style="margin:14px 0 0 0;font-size:13.5px;color:#3a3164;">{summary_line_html}</p>'
        f'{orders_html}'
        f'{outstanding_html}'
        f'<div style="margin:20px 0;padding:12px 16px;background:#f9f8fc;border-left:3px solid '
        f'{RCG_ACCENT_LIGHT};font-size:13.5px;color:#3a3164;">{html_escape(prompt)}</div>'
    )
    html = email_shell_html(body_html)
    plain = (
        f"{greeting}\n\n"
        f"{label} stock recap\n{date_from:%d %b %Y} to {date_to:%d %b %Y}\n\n"
        f"{summary_line_plain}\n\n"
        f"{orders_plain}\n"
        f"{outstanding_plain}"
        f"{prompt}\n"
    )

    subject_tmpl = get_setting(conn, "stock_recap_subject", DEFAULT_SETTINGS["stock_recap_subject"])
    try:
        subject = subject_tmpl.format(date=date_to.strftime("%d %b %Y"), count=str(len(pos)))
    except (KeyError, IndexError):
        subject = subject_tmpl

    charts = {}
    logo_png = rcg_logo_png_bytes()
    if logo_png:
        charts[RCG_LOGO_CID] = logo_png
    return {
        "subject": subject, "html": html, "plain": plain, "po_count": len(pos),
        "outstanding_count": outstanding_count, "pos": pos, "charts": charts,
    }


def mark_stock_recap_sent(conn, when=None):
    set_settings(conn, {"stock_recap_last_sent_at": (when or datetime.now()).isoformat(timespec="seconds")})


_OUTSTANDING_REPORT_STATUS_LABELS = {
    "not_received": "Not received",
    "partial": "Partially received",
    "returned": "Returned",
}


def build_stock_outstanding_report_email(conn, report_rows):
    """Batch 108. The reverse direction of build_stock_recap_email above:
    that one is Yitzi telling the stock team what was ordered ("Hi Tommy,
    here's this week's orders, flag anything missing"); this one is the
    stock team telling Yitzi what's still outstanding right now, so he can
    chase it up -- his own words: "they basically have to show me what
    hasn't yet arrived so we can chase it out."

    report_rows comes from stock_sync.get_outstanding_report_data(conn) --
    same one-way dependency rule as build_stock_recap_email's own
    recap_enrichment argument: po_core.py has zero dependency on
    stock_sync.py, so the caller (po_generator_qt.py) builds that list and
    hands it in here rather than this function importing stock_sync
    itself. Each row is a PO dict with an "items" list already filtered
    down to only the lines NOT YET marked received -- a partially received
    order shows just what's still missing, not the lines already ticked
    off, so nobody has to read past what they actually need to chase."""
    currency_default = get_setting(conn, "currency", "GBP")
    recipient_name = get_setting(conn, "stock_outstanding_report_recipient_name", "")
    greeting = f"Hi {recipient_name}," if recipient_name else "Hi,"

    def render_po_block(p):
        po_currency = p.get("currency") or currency_default
        item_rows = [
            [html_escape(_OUTSTANDING_REPORT_STATUS_LABELS.get(i.get("status"), i.get("status") or "")),
             html_escape(str(i.get("qty", ""))), html_escape(display_product_label(i)),
             html_escape(money(i.get("price"), po_currency))]
            for i in p["items"]
        ]
        items_table = email_table_html(
            ["Status", "Qty", "Item", "Unit price"], item_rows,
            align=["left", "left", "left", "right"],
        )
        html_block = (
            f'<div style="margin-top:18px;padding:12px 14px;background:{RCG_ROW_BLUE};">'
            f'<span style="font-weight:bold;color:{RCG_INK};">{html_escape(p["po_ref"])}</span>'
            f'<span style="color:{RCG_MUTED};"> &mdash; '
            f'{html_escape(p.get("supplier_company_name") or "(no supplier)")} '
            f'(placed {html_escape((p.get("created_at") or "")[:10])})</span></div>'
            + items_table
        )
        item_lines = "".join(
            f"  - [{_OUTSTANDING_REPORT_STATUS_LABELS.get(i.get('status'), i.get('status') or '')}] "
            f"{i.get('qty', '')}x {display_product_label(i)} @ {money(i.get('price'), po_currency)}\n"
            for i in p["items"]
        )
        plain_block = (
            f"{p['po_ref']} - {p.get('supplier_company_name') or '(no supplier)'} "
            f"(placed {(p.get('created_at') or '')[:10]})\n{item_lines}"
        )
        return html_block, plain_block

    blocks_html, blocks_plain = [], []
    for p in report_rows:
        h, pl = render_po_block(p)
        blocks_html.append(h)
        blocks_plain.append(pl)

    orders_html = "".join(blocks_html) or (
        f'<p style="color:{RCG_MUTED};">Nothing outstanding -- everything has been received.</p>'
    )
    orders_plain = "\n".join(blocks_plain) or "Nothing outstanding -- everything has been received.\n"
    summary_line = f"{pluralize(len(report_rows), 'order')} still outstanding."

    body_html = (
        f'<p style="margin:20px 0 4px 0;font-size:15px;color:{RCG_INK};">{html_escape(greeting)}</p>'
        f'<h2 style="margin:0 0 6px 0;font-size:22px;line-height:1.3;color:{RCG_INK};font-weight:bold;">'
        f'Outstanding orders</h2>'
        f'<p style="margin:0;font-size:14px;color:{RCG_MUTED};">What hasn&rsquo;t arrived yet, as of today.</p>'
        f'<p style="margin:14px 0 0 0;font-size:13.5px;color:#3a3164;">{html_escape(summary_line)}</p>'
        f'{orders_html}'
    )
    html = email_shell_html(body_html)
    plain = (
        f"{greeting}\n\nOutstanding orders -- what hasn't arrived yet, as of today.\n\n"
        f"{summary_line}\n\n{orders_plain}"
    )

    subject_tmpl = get_setting(conn, "stock_outstanding_report_subject", DEFAULT_SETTINGS["stock_outstanding_report_subject"])
    try:
        subject = subject_tmpl.format(date=datetime.now().strftime("%d %b %Y"), count=str(len(report_rows)))
    except (KeyError, IndexError):
        subject = subject_tmpl

    charts = {}
    logo_png = rcg_logo_png_bytes()
    if logo_png:
        charts[RCG_LOGO_CID] = logo_png
    return {
        "subject": subject, "html": html, "plain": plain,
        "po_count": len(report_rows), "charts": charts,
    }


def mark_stock_outstanding_report_sent(conn, when=None):
    set_settings(conn, {"stock_outstanding_report_last_sent_at": (when or datetime.now()).isoformat(timespec="seconds")})


_STOCK_REQUEST_KIND_LABELS = {"new": "New item", "existing": "Restock"}


def build_stock_request_summary_email(conn):
    """Batch 150 (Task #95, seeded by Batch 146's Settings card). The
    weekly nudge to Supply listing every stock request Stock has raised
    that's still open -- same "belt and braces on top of the real-time
    notification, not instead of it" relationship
    build_stock_outstanding_report_email above has to individual PO
    status changes, just for stock_request_item/stock_request_new_item
    instead. Always a live snapshot of everything genuinely still open
    right now -- see stock_request_is_due()'s own docstring for why
    there's no period window to pick here, unlike the recap's actual
    placed-orders window."""
    rows = list_product_requests(conn, status="open")

    def render_row(r):
        kind_label = _STOCK_REQUEST_KIND_LABELS.get(r["kind"], r["kind"])
        return [
            html_escape(kind_label),
            html_escape(r["product_name"]),
            html_escape(str(r["qty"]) if r["qty"] else ""),
            html_escape(r.get("supplier_name") or "(no preference)"),
            html_escape(r.get("requester_name") or ""),
            html_escape((r.get("created_at") or "")[:10]),
            html_escape(r.get("notes") or ""),
        ]

    if rows:
        table_html = email_table_html(
            ["Type", "Item", "Qty", "Preferred supplier", "Requested by", "Raised", "Notes"],
            [render_row(r) for r in rows],
            align=["left", "left", "right", "left", "left", "left", "left"],
        )
    else:
        table_html = f'<p style="color:{RCG_MUTED};">Nothing outstanding -- every stock request has been actioned.</p>'

    plain_lines = "".join(
        f"  - [{_STOCK_REQUEST_KIND_LABELS.get(r['kind'], r['kind'])}] "
        f"{(str(r['qty']) + 'x ') if r['qty'] else ''}{r['product_name']} "
        f"(supplier: {r.get('supplier_name') or '(no preference)'}, "
        f"requested by {r.get('requester_name') or ''} on {(r.get('created_at') or '')[:10]})"
        f"{(' -- ' + r['notes']) if r.get('notes') else ''}\n"
        for r in rows
    ) or "Nothing outstanding -- every stock request has been actioned.\n"

    summary_line = f"{pluralize(len(rows), 'request')} still open."
    body_html = (
        f'<h2 style="margin:0 0 6px 0;font-size:22px;line-height:1.3;color:{RCG_INK};font-weight:bold;">'
        f'Stock requests</h2>'
        f'<p style="margin:0;font-size:14px;color:{RCG_MUTED};">Everything Stock has flagged that&rsquo;s '
        f'still waiting on Supply, as of today.</p>'
        f'<p style="margin:14px 0 0 0;font-size:13.5px;color:#3a3164;">{html_escape(summary_line)}</p>'
        f'{table_html}'
    )
    html = email_shell_html(body_html)
    plain = (
        f"Stock requests -- everything Stock has flagged that's still waiting on Supply, as of today.\n\n"
        f"{summary_line}\n\n{plain_lines}"
    )

    # "w/c" (week commencing) in the default subject template means the
    # Monday of the week this is being sent in -- not tied to whichever
    # weekday stock_request_day is actually scheduled for, since "w/c"
    # should read the same regardless of which day mid-week the email
    # happens to go out on.
    now = datetime.now()
    this_monday = now - timedelta(days=now.weekday())
    subject_tmpl = get_setting(conn, "stock_request_subject", DEFAULT_SETTINGS["stock_request_subject"])
    try:
        subject = subject_tmpl.format(date=this_monday.strftime("%d %b %Y"), count=str(len(rows)))
    except (KeyError, IndexError):
        subject = subject_tmpl

    charts = {}
    logo_png = rcg_logo_png_bytes()
    if logo_png:
        charts[RCG_LOGO_CID] = logo_png
    return {
        "subject": subject, "html": html, "plain": plain,
        "request_count": len(rows), "charts": charts,
    }


def mark_stock_request_summary_sent(conn, when=None):
    set_settings(conn, {"stock_request_last_sent_at": (when or datetime.now()).isoformat(timespec="seconds")})


# ============================================================
# 9. Duplicate product detection / merge (suggest-and-confirm only)
# ============================================================
#
# This never merges or changes anything on its own. It only ever proposes
# candidate pairs; the caller (the UI) must present each one to the user and
# get an explicit decision before merge_products() or ignore_duplicate_pair()
# is called.

def _normalize_product_name(name):
    return re.sub(r"[^A-Z0-9]+", " ", (name or "").upper()).strip()


# Model/variant marker words (and the "+" symbol, which means the same thing
# as "PLUS") that make two product names genuinely different products even
# when the rest of the text is near-identical -- e.g. "S25" and "S25+" (a
# base model vs. its Plus variant) must NOT be suggested as duplicates, but
# "S25 Plus" and "S25+" (the same variant, just written differently) should
# still be caught. A pair is only compared for similarity if both names
# carry the *same* set of these markers; otherwise it's skipped outright,
# regardless of how similar the rest of the text looks. This is a genuine,
# unambiguous product difference -- there's no supplier-wording variation
# that turns a base model into a Plus/Pro one, so it stays a hard rule.
_VARIANT_MARKER_RE = re.compile(
    r'(\+|\bPLUS\b|\bPRO\b|\bMAX\b|\bULTRA\b|\bMINI\b|\bLITE\b|\bSE\b|\bFE\b|\bXL\b|\bXXL\b|\bNEO\b|\bAIR\b)',
    re.IGNORECASE,
)

# Storage/size markers like "128GB", "1TB", "500ML" -- two products that are
# identical except for capacity (e.g. "iPhone 16 128GB Black" vs "iPhone 16
# 256GB Black") are different, separately-stocked SKUs, not the same product
# written two different ways, so they must never be suggested as duplicates.
# Also a hard rule, and deliberately so (Yitzi: "different storage sizes...
# should never auto-merge and should not be forced together") -- a storage
# figure is just a number, there's no wording ambiguity the way there is
# with a colour name or a model number, so there's no reason a human ever
# needs to review "is 128GB maybe the same as 256GB".
_CAPACITY_MARKER_RE = re.compile(r'\b(\d+)\s*(GB|TB|MB|ML|CL|KG)\b', re.IGNORECASE)

# Accessory/category words -- "Nokia 3210 Black" and "Nokia 3210 Black Case"
# share almost all of their text (the second is even a superset of the
# first), but a case FOR a phone is not the phone -- these must never be
# treated as the same product no matter how similar the rest of the name
# looks. Also a hard rule -- an accessory is never just a wording variant
# of the product it's for.
_ACCESSORY_WORD_RE = re.compile(
    r'\b(CASE|COVER|PROTECTOR|CHARGER|CABLE|ADAPTER|ADAPTOR|STAND|HOLDER|MOUNT'
    r'|SLEEVE|POUCH|STRAP|DOCK)\b',
    re.IGNORECASE,
)

# Brand words/aliases (Batch 61) -- a handset's brand is a hard fact, never
# just a wording variant, per Yitzi's spec: storage, brand, and model must
# never be suggested as a match when they clearly differ. Several aliases
# map to the same canonical brand ("iPhone" implies Apple even when the
# word "Apple" isn't written; "Galaxy" implies Samsung) so two names don't
# get treated as different brands just because one wrote the product-line
# name and the other wrote the maker's name. Only used to EXCLUDE a pair
# when both sides confidently name exactly one, different brand -- a name
# with no recognizable brand word at all (not every catalog entry has one)
# never triggers this on its own, see _brand_markers.
_BRAND_ALIASES = {
    "APPLE": "APPLE", "IPHONE": "APPLE",
    "SAMSUNG": "SAMSUNG", "GALAXY": "SAMSUNG",
    "GOOGLE": "GOOGLE", "PIXEL": "GOOGLE",
    "NOKIA": "NOKIA",
    "HUAWEI": "HUAWEI",
    "XIAOMI": "XIAOMI", "REDMI": "XIAOMI", "POCO": "XIAOMI",
    "ONEPLUS": "ONEPLUS",
    "SONY": "SONY", "XPERIA": "SONY",
    "MOTOROLA": "MOTOROLA", "MOTO": "MOTOROLA",
    "OPPO": "OPPO",
    "VIVO": "VIVO",
    "HONOR": "HONOR",
    "REALME": "REALME",
    "LENOVO": "LENOVO",
    "MICROSOFT": "MICROSOFT", "SURFACE": "MICROSOFT",
    "ASUS": "ASUS",
    "HTC": "HTC",
    "LG": "LG",
    "YEALINK": "YEALINK",
}
_BRAND_MARKER_RE = re.compile(
    r'\b(' + '|'.join(sorted(_BRAND_ALIASES.keys(), key=len, reverse=True)) + r')\b',
    re.IGNORECASE,
)

# Common colour words -- colour is NOT a flat hard-exclusion rule the way
# storage/brand/model are (see _fuzzy_product_match) -- different suppliers
# genuinely do write the same colour differently ("Blue" vs "Navy", "Grey"
# vs "Titanium"), and hard-blocking on colour risks hiding a real duplicate
# before a human ever sees it. But two colour words that aren't even
# plausibly the same real-world colour (eg. "Green" vs "Yellow") are also
# never suggested -- see _COLOUR_FAMILY and _fuzzy_product_match for where
# that line is drawn.
_COLOR_MARKER_RE = re.compile(
    r'\b(BLACK|WHITE|BLUE|RED|GREEN|GREY|GRAY|SILVER|GOLD|PURPLE|PINK|YELLOW|ORANGE'
    r'|TITANIUM|GRAPHITE|MIDNIGHT|STARLIGHT|NATURAL|BEIGE|BRONZE|COPPER|NAVY|TEAL'
    r'|CREAM|CHARCOAL|MINT|CORAL|LAVENDER|MAROON|IVORY|CRIMSON|TURQUOISE)\b',
    re.IGNORECASE,
)

# Which "family" each recognized colour word belongs to (Batch 61) -- two
# colour words in the SAME family (eg. "Blue" and "Navy", or "Cream" and
# "White") are plausibly the same real-world colour written differently by
# different suppliers, so that pair still surfaces with a caution for a
# human to decide. Two colour words in DIFFERENT families (eg. "Green" and
# "Yellow") are never plausibly the same colour, so that pair is excluded
# outright, same as a genuine storage/brand/model difference -- this is a
# judgment call on real phone colour naming, not an exact science, and can
# be adjusted if a specific pair of words turns out to be grouped wrong.
_COLOUR_FAMILY = {
    "BLACK": "BLACK", "MIDNIGHT": "BLACK", "CHARCOAL": "BLACK", "GRAPHITE": "BLACK",
    "WHITE": "WHITE", "CREAM": "WHITE", "IVORY": "WHITE", "STARLIGHT": "WHITE",
    "NATURAL": "WHITE", "BEIGE": "WHITE",
    "BLUE": "BLUE", "NAVY": "BLUE", "TEAL": "BLUE", "TURQUOISE": "BLUE",
    "GREY": "GREY", "SILVER": "GREY", "TITANIUM": "GREY",
    "GOLD": "GOLD", "BRONZE": "GOLD", "COPPER": "GOLD",
    "GREEN": "GREEN", "MINT": "GREEN",
    "RED": "RED", "MAROON": "RED", "CRIMSON": "RED",
    "PINK": "PINK", "CORAL": "PINK",
    "PURPLE": "PURPLE", "LAVENDER": "PURPLE",
    "YELLOW": "YELLOW",
    "ORANGE": "ORANGE",
}

# Any alphanumeric token that carries at least one digit -- model numbers
# ("A56", "A11", "S25", "X135", "17E", "17"), part/SKU codes ("MHRV4QN"),
# generation/RAM shorthand ("5G", the standalone "8" in "8/128GB"), year
# suffixes, etc. Two names whose token sets genuinely conflict (neither is
# a subset of the other -- eg. "17E" on one side, "17" on the other, with
# nothing in common) are excluded outright (Batch 61 -- Yitzi: "a 17E is
# not a match to 17", never even wants that pair suggested). A name
# imported verbatim from a supplier/Zoho often carries an extra
# manufacturer part number ("Apple iPhone 17e 256GB Black MHRV4QN/A") that
# a manually-cleaned catalog entry ("iPhone 17e 256GB Black") doesn't --
# that's a subset relationship (no conflict), so it's still allowed
# through, same as before.
_MODEL_CODE_TOKEN_RE = re.compile(r'\b[A-Z0-9]*\d[A-Z0-9]*\b', re.IGNORECASE)


def _variant_markers(name):
    """Returns the set of HARD-rule markers found in a product name --
    model/Plus-Pro/Max variant, storage capacity, and accessory words only.
    Two products are only ever compared at all if this set matches exactly
    -- these three are unambiguous, genuine product differences with no
    "maybe it's just written differently" case (a Pro model, a different
    storage size, or an accessory FOR a product are never the same product
    as what they're being compared to). See _brand_markers and
    _model_code_tokens for the other two hard exclusions (Batch 61), and
    _colour_markers/_COLOUR_FAMILY for the one dimension that's still only
    a caution, not a hard exclusion."""
    found = set()
    text = name or ""
    for m in _VARIANT_MARKER_RE.finditer(text):
        token = m.group(1).upper()
        if token == "+":
            token = "PLUS"
        found.add(token)
    for m in _CAPACITY_MARKER_RE.finditer(text):
        found.add(f"{m.group(1)}{m.group(2).upper()}")
    for m in _ACCESSORY_WORD_RE.finditer(text):
        found.add(m.group(1).upper())
    return found


def _product_names_match(name_a, name_b):
    """True only when two names refer to the SAME catalogue product --
    every place in this app that decides "is this an existing product or a
    genuinely new one" (record_product_price, upsert_product_manual's
    create-path dedup, find_close_matching_products' own exact-match
    shortcut, ensure_product_in_catalog, and the Zoho import planner) must
    use this, not a bare _normalize_product_name(...) == ... comparison.

    Audit finding (tidy-up pass, post-Batch-151): _normalize_product_name
    strips "+" as ordinary punctuation, so "Galaxy S25+" normalized to the
    exact same text as "Galaxy S25" -- meaning every one of those five call
    sites was silently treating a Plus/Pro/Max/storage-capacity variant as
    "just a differently-typed version of the same product", directly
    contradicting this same file's own explicit, hard rule (_variant_markers,
    used by find_close_matching_products' fuzzy comparison below) that a
    variant marker makes two names genuinely different products, never a
    spelling difference. Concretely, typing/importing "Galaxy S25+" for a
    supplier that already stocks "Galaxy S25" would silently overwrite the
    base model's own catalog row (price, code, everything) instead of
    creating the Plus variant as its own product. Requiring BOTH the
    normalized text AND the variant-marker set to match closes that gap
    while leaving every ordinary case (different capitalization/spacing,
    no markers involved at all) working exactly as before."""
    return (
        _normalize_product_name(name_a) == _normalize_product_name(name_b)
        and _variant_markers(name_a) == _variant_markers(name_b)
    )


def _brand_markers(name):
    """Every canonical brand recognized in a product name (Batch 61) -- eg.
    {"APPLE"} for a name containing "Apple" or "iPhone", {"SAMSUNG"} for
    "Samsung" or "Galaxy". Usually 0 or 1 elements; see _fuzzy_product_match
    for how this is used (only excludes a pair when BOTH sides confidently
    name exactly one, different brand -- a name with no recognized brand at
    all, or an ambiguous one naming more than one, never excludes on its
    own)."""
    found = set()
    for m in _BRAND_MARKER_RE.finditer(name or ""):
        found.add(_BRAND_ALIASES[m.group(1).upper()])
    return found


def _colour_markers(name):
    """Every colour word found in a product name, canonicalized so British/
    American spelling of the same colour reads as one thing (eg. "Grey"
    and "Gray" both become "GREY"). See _fuzzy_product_match and
    _COLOUR_FAMILY for how this is used -- two colours in the same family
    only raise a caution, two colours in different families exclude the
    pair outright."""
    found = set()
    for m in _COLOR_MARKER_RE.finditer(name or ""):
        color = m.group(1).upper()
        if color == "GRAY":
            color = "GREY"
        found.add(color)
    return found


def _model_code_tokens(name):
    """Every digit-bearing alphanumeric token in a product name (see
    _MODEL_CODE_TOKEN_RE above). See _fuzzy_product_match for how this is
    used -- a genuine conflict (neither token set is a subset of the
    other) excludes the pair outright as of Batch 61."""
    return {m.group(0).upper() for m in _MODEL_CODE_TOKEN_RE.finditer(name or "")}


def _fuzzy_match_entry(name):
    """Builds the (normalized_name, hard_markers, model_code_tokens,
    colour_markers, brand_markers) tuple _fuzzy_product_match compares --
    computed once per product name and reused across every pair it's
    checked against, rather than recomputed per pair."""
    return (
        _normalize_product_name(name),
        _variant_markers(name),
        _model_code_tokens(name),
        _colour_markers(name),
        _brand_markers(name),
    )


def _fuzzy_product_match(entry_a, entry_b, min_ratio, matcher=None):
    """Shared "are these two product names probably the same thing" check,
    used both by find_duplicate_product_candidates (catalog cleanup) and
    find_orphaned_purchase_history (history repair) so a fix to this logic
    always applies to both at once. entry_a/entry_b are each
    (normalized_name, hard_markers_set, model_code_tokens_set,
    colour_markers_set, brand_markers_set) tuples, as produced by
    _fuzzy_match_entry. Returns None if a hard guard excludes the pair
    outright, or if the text isn't similar enough even to be worth
    showing. Otherwise returns (ratio, caution): ratio is the similarity
    score (float, 0..1); caution is None for an ordinary match, or a short
    human-readable string when the colour words name plausibly-the-same-
    family colours written differently (eg. "different colour written
    (BLUE vs NAVY)") -- that one case is deliberately never excluded, since
    a supplier writing the same colour inconsistently is common and must
    never silently hide a real duplicate, but it's flagged so a genuinely
    different product doesn't get waved through on a deceptively high
    text-similarity score alone.

    As of Batch 61 (direct spec from Yitzi: "I don't need any product to
    be suggested as a match if they have clearly different storage, brand,
    or model"), a pair is EXCLUDED outright -- never even shown, not just
    never auto-merged -- when: variant markers differ (Plus/Pro/Max/
    capacity/accessory-word, unchanged from before); both sides confidently
    name exactly one, different brand; the model/part-number tokens
    genuinely conflict (eg. "17E" vs "17", "A56" vs "A35"); or the colour
    words are in clearly different families (eg. "Green" vs "Yellow").
    Colour words in the SAME family (eg. "Blue" vs "Navy", "Cream" vs
    "White") are the one remaining case that still just raises a caution
    rather than excluding -- real-world colour naming is genuinely
    ambiguous in a way brand/model/storage never are, so that call is
    still left to a human.
    """
    na, va, ta, ca, ba = entry_a
    nb, vb, tb, cb, bb = entry_b
    if not na or not nb or va != vb:
        return None
    # Brand: only exclude when BOTH sides confidently name exactly one
    # brand and they differ -- a name with no recognized brand word at all
    # (plenty of catalog entries don't have one) or an ambiguous one naming
    # more than one never excludes on its own.
    if len(ba) == 1 and len(bb) == 1 and ba != bb:
        return None
    # Model/part-number tokens: exclude on a genuine conflict (neither set
    # is a subset of the other). A one-sided extra manufacturer part number
    # (a true subset relationship) is still fine and falls through.
    if ta and tb:
        smaller, larger = (ta, tb) if len(ta) <= len(tb) else (tb, ta)
        if not smaller <= larger:
            return None
    # Colour: only a caution (never excluded) when both sides name a
    # colour and share at least one family; excluded outright when they
    # name colours from entirely different families.
    caution = None
    if ca and cb and ca != cb:
        fam_a = {_COLOUR_FAMILY.get(c, c) for c in ca}
        fam_b = {_COLOUR_FAMILY.get(c, c) for c in cb}
        if not (fam_a & fam_b):
            return None
        caution = f"different colour written ({', '.join(sorted(ca))} vs {', '.join(sorted(cb))})"

    if matcher is None:
        matcher = difflib.SequenceMatcher(None, na)
    matcher.set_seq2(nb)
    if na in nb or nb in na:
        # One name fully contains the other (eg. a manufacturer part number
        # tacked on the end) -- a strong signal even when the raw ratio
        # would fall a bit short, so it's boosted to at least 0.9. But that
        # boost must never let a pair through BELOW what the caller
        # actually asked for -- the auto-merge pass deliberately calls this
        # with min_ratio=0.999 specifically because it merges with no human
        # review at all, and a containment pair sitting at, say, 0.93 raw
        # ratio has no business clearing that bar just because 0.9 < 0.93.
        ratio = max(matcher.ratio(), 0.9)
    else:
        if matcher.real_quick_ratio() < min_ratio or matcher.quick_ratio() < min_ratio:
            return None
        ratio = matcher.ratio()
    if ratio < min_ratio:
        return None

    return (ratio, caution)


def reset_product_catalog(conn):
    """Wipes the product catalog and every merge/dismiss decision made
    against it, then rebuilds the catalog completely fresh straight from PO
    history -- as if every order had been entered for the first time with
    no linking ever done. Requested directly by Yitzi after the Find
    Possible Duplicates / Repair Purchase History / Review Dismissed Pairs
    trio (batches 54, 56-59) grew into more buttons and options than he
    wanted to deal with: "restore all the products back to how they were,
    no more anything linked, and then start linking from the beginning
    again."

    What this touches:
    - products: deleted entirely, then rebuilt one row per (supplier, PO
      line-item name) by replaying every non-deleted PO's line items
      through record_product_price in chronological order -- the exact
      same function normal PO saving uses, so the rebuilt catalog is
      indistinguishable from one that had simply never been merged or
      edited. last_price ends up as the most recent price paid, and
      times_ordered as a true count of how many line items used that exact
      name -- both recomputed from scratch, not carried over.
    - product_merge_ignored: cleared entirely. Every "Not a duplicate"
      dismissal is forgotten, since the products those decisions were
      about no longer exist as the same rows once the catalog is rebuilt.

    What this does NOT touch: po_items.product/code text itself. Past
    merges (and manual renames made with rewrite_history=True) already
    rewrote that text permanently -- e.g. if "S26 Ultra", "Galaxy S26
    Ultra" and "Samsung Galaxy S26 Ultra" were previously merged into one,
    every one of those historical line items now says "S26 Ultra" and
    there is no record left anywhere of the original wording, so the
    rebuilt catalog will show one "S26 Ultra" product rather than three
    separate ones. That's expected and unavoidable -- only a full backup
    restore could undo it, which rolls back everything else too, not just
    the catalog. purchase_orders and po_items themselves are completely
    unaffected either way.

    One consolidation IS applied automatically after the raw rebuild,
    though: any product bought under the exact same name text from more
    than one real supplier comes out of the raw replay as one row per
    supplier (products is keyed by (supplier_id, name), same as normal PO
    saving), even though it's plainly one product. That's not "linking"
    in the fuzzy-match sense Yitzi wanted undone -- it's the same
    zero-judgment case MergeDuplicatesDialog already auto-merges on open
    (identical name, nothing to decide) -- so this rebuild folds those
    exact-text groups back into one row per name before returning,
    otherwise a reset would visibly fragment products like "iPhone 17 Pro
    Max 256GB Silver" (ordered from three different suppliers under that
    exact text) into three duplicate-looking catalog rows all showing the
    same combined purchase history when opened.

    Returns the number of distinct products the rebuild produced. Callers
    should take a backup (see backup_now) before calling this -- it is
    destructive to the current catalog state and has no separate undo of
    its own.
    """
    conn.execute("DELETE FROM product_merge_ignored")
    conn.execute("DELETE FROM products")
    rows = conn.execute(
        "SELECT pi.product AS product, pi.code AS code, pi.price AS price, "
        "       po.supplier_id AS supplier_id, po.created_at AS created_at "
        "FROM po_items pi JOIN purchase_orders po ON po.id = pi.po_id "
        "WHERE po.deleted = 0 AND pi.product != '' "
        "ORDER BY po.created_at ASC, po.id ASC, pi.id ASC"
    ).fetchall()
    for r in rows:
        record_product_price(conn, r["supplier_id"], r["code"], r["product"], r["price"], at=r["created_at"])
    # merge_products opens its own transaction ("BEGIN") below, which SQLite
    # refuses to do while one is already open -- the DELETEs/INSERTs above
    # implicitly started one, so it has to be closed out first.
    conn.commit()

    # Fold exact-text, cross-supplier duplicates back into one row each --
    # same "auto-merge every 100%, non-cautioned match" pass
    # MergeDuplicatesDialog already runs on open, reused here so a reset
    # doesn't visibly fragment a product that was simply ordered from more
    # than one supplier under identical wording.
    exact_matches = find_duplicate_product_candidates(conn, min_ratio=0.999, limit=None)
    for c in exact_matches:
        if c.get("caution"):
            continue
        still_a = conn.execute("SELECT 1 FROM products WHERE id=?", (c["product_id_a"],)).fetchone()
        still_b = conn.execute("SELECT 1 FROM products WHERE id=?", (c["product_id_b"],)).fetchone()
        if not still_a or not still_b:
            continue
        try:
            merge_products(conn, c["product_id_a"], c["product_id_b"])
        except Exception:
            continue

    conn.commit()
    return conn.execute("SELECT COUNT(*) AS n FROM products").fetchone()["n"]


def find_duplicate_product_candidates(conn, min_ratio=0.55, limit=200):
    """Return a list of candidate duplicate product pairs, sorted by
    similarity (best matches first). Compares every product against every
    other product REGARDLESS of which supplier it's filed under -- the
    same physical product is very often genuinely bought from more than
    one supplier, and sometimes recorded under a slightly different name
    each time (eg. "Tab Active 5 5G 128GB" from one supplier, "Tab Active 5
    5G 128GB (X306)" from another) -- that's exactly the case this is
    meant to catch and unify, so a Purchase history / Scorecard / price
    comparison for "the same product" actually reflects every time it's
    been bought, from anyone, under any name it's been entered as. A
    supplier being different is not on its own a reason two products can't
    be the same thing -- merge_products already knows how to combine two
    different suppliers' history into one product correctly. Excludes any
    pair already dismissed via ignore_duplicate_pair.

    Pairs whose names carry different Plus/Pro/Max/Ultra/Mini/... markers,
    a different storage capacity, an accessory word like "Case", a
    different (confidently recognized) brand, or a genuinely conflicting
    model/part number (eg. "17E" vs "17", "A56" vs "A35") are never
    suggested -- those are unambiguous, different products, not the same
    product written two different ways (Batch 61, direct spec from Yitzi).
    A colour difference is the one dimension still NOT excluded when the
    two colours are at least plausibly the same real-world colour (eg.
    Blue vs Navy, Cream vs White) -- suppliers write those inconsistently
    all the time, so that pair still surfaces here with a "caution" note on
    the candidate dict, and it's a human's call in the review dialog
    whether to merge or reject. Two colours that are NOT plausibly the same
    (eg. Green vs Yellow) are excluded outright, same as brand/model/
    storage. See _fuzzy_product_match for the full matching logic (also
    used by find_orphaned_purchase_history, so a fix to it applies to both
    tools at once).
    """
    products = [dict(r) for r in conn.execute("SELECT * FROM products")]
    ignored = {
        (r["product_id_a"], r["product_id_b"])
        for r in conn.execute("SELECT * FROM product_merge_ignored")
    }

    # Precompute the normalized name/variant markers/model-code tokens once
    # per product rather than once per pair, and bucket products by their
    # exact variant-marker set (colour/capacity/Plus-Pro/accessory-word)
    # rather than by supplier -- _fuzzy_product_match already requires an
    # exact marker-set match before it'll even look at text similarity, so
    # two products that can never match are guaranteed to already land in
    # different buckets. This keeps the pairwise comparison below cheap
    # across the whole catalog (a typical catalog splits into many small
    # marker buckets rather than one everything-together group) without
    # needing the old supplier-based grouping at all.
    info_by_id = {p["id"]: _fuzzy_match_entry(p["name"]) for p in products}
    by_markers = {}
    for p in products:
        entry = info_by_id[p["id"]]
        if not entry[0]:
            continue
        by_markers.setdefault(frozenset(entry[1]), []).append(p)

    candidates = []
    for group in by_markers.values():
        if len(group) < 2:
            continue
        for i in range(len(group)):
            a = group[i]
            entry_a = info_by_id[a["id"]]
            matcher = difflib.SequenceMatcher(None, entry_a[0])
            for j in range(i + 1, len(group)):
                b = group[j]
                pair = tuple(sorted((a["id"], b["id"])))
                if pair in ignored:
                    continue
                match = _fuzzy_product_match(entry_a, info_by_id[b["id"]], min_ratio, matcher=matcher)
                if match is not None:
                    ratio, caution = match
                    candidates.append({
                        "product_id_a": a["id"],
                        "product_id_b": b["id"],
                        "product_a": a,
                        "product_b": b,
                        "ratio": ratio,
                        "caution": caution,
                    })

    candidates.sort(key=lambda c: c["ratio"], reverse=True)
    return candidates[:limit]


def find_close_matching_products(conn, name, min_ratio=0.72, exclude_product_id=None):
    """Used right before a brand-new catalog entry is about to be created by
    hand (Products page "Add product") to check whether something close
    already exists first -- Yitzi's "close match detection" requirement:
    "When creating a genuinely new product, check whether a close or
    similar product already exists. If a close match is found: alert the
    user, show the existing similar product(s), allow the user to choose to
    use the existing product instead, [or] continue and deliberately create
    a new product if it is genuinely different. Do not silently merge
    genuinely different products."

    Reuses the exact same _fuzzy_match_entry/_fuzzy_product_match logic
    find_duplicate_product_candidates already uses for after-the-fact
    catalog cleanup, so "close" means the same thing in both places -- a
    pair that would be excluded outright there (different capacity, brand,
    or model/part number, or a different colour family entirely) is never
    suggested here either, so a genuinely different product is never
    wrongly flagged. min_ratio=0.72 matches the same "worth a human's
    attention" threshold MergeDuplicatesDialog's own review pass uses.

    Deliberately does NOT surface an exact case/whitespace-only match
    (ratio 1.0, same normalized name) -- that case is handled silently,
    with no prompt at all, by upsert_product_manual's own normalized-name
    check (Batch 77): it just becomes an update to the existing row. This
    function is only for the "close, but not simply the same text
    differently capitalized" case, which genuinely needs a human's
    judgment call.

    Returns a list of dicts (closest first): {"product_id", "name",
    "supplier_id", "last_price", "times_ordered", "ratio", "caution"} --
    empty if nothing close enough is on file, or if `name` is blank.
    """
    norm = _normalize_product_name(name)
    if not norm:
        return []
    entry_new = _fuzzy_match_entry(name)
    matcher = difflib.SequenceMatcher(None, entry_new[0])
    matches = []
    for r in conn.execute("SELECT id, name, supplier_id, last_price, times_ordered FROM products"):
        if exclude_product_id is not None and r["id"] == exclude_product_id:
            continue
        if _product_names_match(r["name"], name):
            continue  # exact match -- silently handled elsewhere, not a "close match" prompt
        match = _fuzzy_product_match(entry_new, _fuzzy_match_entry(r["name"]), min_ratio, matcher=matcher)
        if match is None:
            continue
        ratio, caution = match
        matches.append({
            "product_id": r["id"], "name": r["name"], "supplier_id": r["supplier_id"],
            "last_price": r["last_price"], "times_ordered": r["times_ordered"],
            "ratio": ratio, "caution": caution,
        })
    matches.sort(key=lambda m: m["ratio"], reverse=True)
    return matches


def find_orphaned_purchase_history(conn, min_ratio=0.72):
    """Every distinct product name that shows up in PO line-item history but
    no longer matches ANY current catalog product exactly.

    Purchase history and the catalog's live Supplier/Last price columns are
    matched to po_items by exact name text (see get_product_purchase_history
    / get_product_catalog_live_info), not by a stable id -- so a product
    renamed via Products > Edit selected before that started rewriting
    history too (see upsert_product_manual's rewrite_history), or a product
    that's since been deleted outright, can leave PO line items pointing at
    a name that doesn't match anything in the catalog any more. Those items
    are still there and still count in totals, they just silently stop
    showing up under any product's own Purchase history / Scorecard tabs.

    For each orphaned name, suggests the closest current catalog product
    it's probably the same thing as, using the same fuzzy-match logic (and
    the same Plus/Pro/capacity/colour variant guards) as
    find_duplicate_product_candidates -- an orphaned "iPhone 16 128GB
    Black" is never suggested to relink onto a current "iPhone 16 256GB
    Black".

    Returns a list of dicts, most-affected first:
    {"old_name": str, "item_count": int, "po_count": int,
     "suggested_id": int or None, "suggested_name": str or None, "ratio": float,
     "caution": str or None}
    suggested_id is None when nothing on file looks like a confident match
    (ratio below min_ratio) -- that name needs a human to pick manually.
    caution carries the same "different model/part number..." / "different
    colour written..." note as find_duplicate_product_candidates when the
    best match's model/colour tokens genuinely conflict with the orphaned
    name's -- it's never a reason to withhold the suggestion, only a flag
    for the human reviewing it.
    """
    products = [dict(r) for r in conn.execute("SELECT * FROM products")]
    current_names = {p["name"] for p in products}
    info_by_name = {p["name"]: _fuzzy_match_entry(p["name"]) for p in products}

    rows = conn.execute(
        "SELECT pi.product AS product, pi.po_id AS po_id FROM po_items pi "
        "JOIN purchase_orders po ON po.id = pi.po_id WHERE po.deleted = 0 AND pi.product != ''"
    ).fetchall()
    counts = {}
    for r in rows:
        name = r["product"]
        if name in current_names:
            continue
        entry = counts.setdefault(name, {"item_count": 0, "po_ids": set()})
        entry["item_count"] += 1
        entry["po_ids"].add(r["po_id"])

    results = []
    for old_name, agg in counts.items():
        entry_a = _fuzzy_match_entry(old_name)
        best_id, best_name, best_ratio, best_caution = None, None, 0.0, None
        if entry_a[0]:
            matcher = difflib.SequenceMatcher(None, entry_a[0])
            for p in products:
                match = _fuzzy_product_match(entry_a, info_by_name[p["name"]], min_ratio, matcher=matcher)
                if match is not None and match[0] > best_ratio:
                    best_ratio, best_caution = match
                    best_id, best_name = p["id"], p["name"]
        results.append({
            "old_name": old_name,
            "item_count": agg["item_count"],
            "po_count": len(agg["po_ids"]),
            "suggested_id": best_id if best_ratio >= min_ratio else None,
            "suggested_name": best_name if best_ratio >= min_ratio else None,
            "ratio": best_ratio,
            "caution": best_caution if best_ratio >= min_ratio else None,
        })
    results.sort(key=lambda r: r["item_count"], reverse=True)
    return results


def relink_orphaned_history(conn, old_name, new_product_id):
    """Rewrites every po_items.product cell that still says old_name to the
    current name of new_product_id -- the manual-repair counterpart to
    upsert_product_manual's automatic rewrite_history, for history that was
    already orphaned before that existed. Only touches the product label
    text; each item's own code/price/qty is left exactly as it was.
    Returns how many PO line items were relinked."""
    row = conn.execute("SELECT name FROM products WHERE id=?", (new_product_id,)).fetchone()
    if not row:
        raise ValueError("That product no longer exists.")
    new_name = row["name"]
    cur = conn.execute("UPDATE po_items SET product=? WHERE product=?", (new_name, old_name))
    conn.commit()
    return cur.rowcount


def ignore_duplicate_pair(conn, id_a, id_b):
    """Mark a suggested pair as 'not a duplicate' so it's never suggested
    again."""
    lo, hi = sorted((id_a, id_b))
    # Batch 157: ON CONFLICT DO NOTHING instead of "INSERT OR IGNORE" -- same
    # portability swap as create_user()/set_user_permissions() above, needed
    # for the move off Turso onto real Postgres. product_merge_ignored's own
    # PRIMARY KEY is (product_id_a, product_id_b).
    conn.execute(
        "INSERT INTO product_merge_ignored(product_id_a, product_id_b) VALUES (?, ?) "
        "ON CONFLICT (product_id_a, product_id_b) DO NOTHING",
        (lo, hi),
    )
    conn.commit()


def unignore_duplicate_pair(conn, id_a, id_b):
    """Reverses ignore_duplicate_pair -- used by the "Back" button in the
    duplicate-review dialog, so clicking 'Not a duplicate' by mistake can
    actually be undone rather than the pair being silently lost forever."""
    lo, hi = sorted((id_a, id_b))
    conn.execute(
        "DELETE FROM product_merge_ignored WHERE product_id_a=? AND product_id_b=?",
        (lo, hi),
    )
    conn.commit()


def find_ignored_duplicate_candidates(conn):
    """Every pair still sitting in product_merge_ignored -- ie. every pair a
    human has clicked "Not a duplicate" on in Find Possible Duplicates,
    which is exactly why find_duplicate_product_candidates never suggests
    it again.

    This exists because that dismissal is permanent and the matching logic
    underneath it has changed twice since (batch 57: cross-supplier pairs
    can match at all; batch 58: a colour or model/part-number difference no
    longer hard-excludes a pair, it surfaces with a caution instead) --
    someone using an earlier, more confusing version of this dialog (no
    supplier names shown, no caution note, cross-supplier pairs invisible)
    could easily have dismissed a pair that was never actually a false
    positive, just poorly explained at the time. Those pairs are otherwise
    gone for good: nothing else in the app ever surfaces something once
    it's in this table. Powers the Product Catalog's "Review dismissed
    pairs" button, which reopens each one with today's clearer UI so it can
    get a second, better-informed look -- reviewing here never changes
    anything by itself, same as the main dialog.

    Returns the same shape as find_duplicate_product_candidates (including
    "ratio"/"caution"), sorted the same way, but sourced from the ignored
    table instead of a fresh similarity scan, and with no min_ratio floor
    -- a pair that was explicitly marked as a duplicate-candidate once
    stays reviewable however its score reads today. Skips any pair where
    either product has since been deleted or merged away (nothing left to
    review).
    """
    rows = conn.execute("SELECT product_id_a, product_id_b FROM product_merge_ignored").fetchall()
    candidates = []
    for r in rows:
        a = conn.execute("SELECT * FROM products WHERE id=?", (r["product_id_a"],)).fetchone()
        b = conn.execute("SELECT * FROM products WHERE id=?", (r["product_id_b"],)).fetchone()
        if not a or not b:
            continue
        a, b = dict(a), dict(b)
        entry_a, entry_b = _fuzzy_match_entry(a["name"]), _fuzzy_match_entry(b["name"])
        # Bypass the hard-marker gate here (unlike find_duplicate_product_candidates)
        # -- these were already explicitly flagged as possible duplicates once by a
        # human, so it's still worth a second look even if today's guards would
        # never have surfaced them fresh (eg. the marker-detection regexes have
        # changed since some of these were first dismissed). Fall back to a plain
        # text ratio (and a generic caution) for a pair the hard-marker gate would
        # still refuse to score at all, so it's never silently dropped from this
        # list -- it's always worth a look, even if the tool itself still thinks
        # it's an unambiguous non-match.
        match = _fuzzy_product_match(entry_a, entry_b, 0.0)
        if match is not None:
            ratio, caution = match
        else:
            matcher = difflib.SequenceMatcher(None, entry_a[0], entry_b[0])
            ratio = matcher.ratio() if entry_a[0] and entry_b[0] else 0.0
            caution = "different capacity, brand, model, variant, accessory word, or colour family -- usually a genuine non-match"
        candidates.append({
            "product_id_a": a["id"],
            "product_id_b": b["id"],
            "product_a": a,
            "product_b": b,
            "ratio": ratio,
            "caution": caution,
        })
    candidates.sort(key=lambda c: c["ratio"], reverse=True)
    return candidates


def merge_products(conn, keep_id, remove_id, merged_name=None, merged_code=None, rewrite_history=True, capture_undo=False):
    """Merge remove_id into keep_id. Only ever called after an explicit user
    confirmation in the UI — never automatically.

    - The surviving row (keep_id) gets the chosen name/code (or, if not
      given, whichever of the two was more recently updated), the summed
      times_ordered, and the more recent last_price.
    - If rewrite_history is True, past PO line items that used the removed
      product's exact name/code get rewritten to the kept name/code, so
      historical reports read as one consistent product going forward. Past
      PO documents already sent/saved are untouched (this only affects how
      the catalog and future reports label them; it does not edit PDFs).
    - Safe against the UNIQUE(supplier_id, name) constraint: if the chosen
      final name collides with a third, unrelated product for that
      supplier, the merge is aborted and an error is raised rather than
      corrupting data.
    - If capture_undo is True, returns (result, undo_snapshot) instead of
      just result -- undo_snapshot can be passed straight to
      undo_product_merge() to reverse exactly this merge (the removed
      product's row, the kept product's pre-merge name/code/price/times
      ordered, every po_items row that got rewritten, and any dismissed-
      duplicate pairs cleared as a side effect). Callers that don't need
      undo (eg. the bulk manual-merge dialog) can ignore this and keep
      getting just result back.
    """
    a = conn.execute("SELECT * FROM products WHERE id=?", (keep_id,)).fetchone()
    b = conn.execute("SELECT * FROM products WHERE id=?", (remove_id,)).fetchone()
    if not a or not b:
        raise ValueError("One or both products no longer exist")
    a, b = dict(a), dict(b)

    newer, older = (a, b) if (a.get("updated_at") or "") >= (b.get("updated_at") or "") else (b, a)
    final_name = merged_name or newer["name"]
    final_code = merged_code if merged_code is not None else newer["code"]
    final_price = newer["last_price"]
    final_times = (a.get("times_ordered") or 0) + (b.get("times_ordered") or 0)
    final_updated = now_iso()

    clash = conn.execute(
        "SELECT id FROM products WHERE supplier_id=? AND name=? AND id NOT IN (?, ?)",
        (a["supplier_id"], final_name, keep_id, remove_id),
    ).fetchone()
    if clash:
        raise ValueError(
            f"Can't merge: another product already named '{final_name}' exists for this supplier."
        )

    undo_snapshot = None
    try:
        # See save_po's own comment on its equivalent guard (tidy-up pass,
        # post-Batch-151) -- self-heals against a dangling transaction left
        # by some unrelated, unprotected write elsewhere on this connection.
        if conn.in_transaction:
            conn.rollback()
        conn.execute("BEGIN")
        rewritten_items = []
        if rewrite_history:
            # Rewrite BOTH sides' own original name/code, not just the
            # removed product's -- final_name is usually whichever side was
            # more recently updated, which is very often the REMOVED side
            # (b), not the kept one (a). When that happens, a's own past
            # po_items still carry a's original (now-superseded) text, and
            # need rewriting to final_name just as much as b's do. The
            # original version of this only ever rewrote b's old text,
            # silently leaving the kept row's own history orphaned any time
            # final_name != a's original name -- found while chasing a
            # report that case-only duplicates ("Samsung" vs "SAMSUNG")
            # weren't merging cleanly. No-op skipped for whichever side
            # already matches final_name (usually saves one pointless
            # UPDATE, but also avoids capturing a bogus "rewritten" entry in
            # the undo snapshot for text that never actually changed).
            sides = [(a["name"], a["code"]), (b["name"], b["code"])]
            # See upsert_product_manual's matching comment (Batch 62 fix) --
            # a po_items row is rewritten when its product text matches AND
            # either its own code matches that side's old code, that side's
            # old code was blank, or THAT ROW'S OWN code is blank. Without
            # that last clause, a product ordered once with no code and once
            # with a code filled in would only get its coded order
            # rewritten by a merge, silently leaving the blank-code order
            # behind under the old (pre-merge) name -- it kept counting
            # toward the merged times_ordered total but stopped showing up
            # in Purchase History, which looks the product up by its
            # current, now-mismatched name.
            code_guard = "(code=? OR ?='' OR code='')"
            cross_supplier = a["supplier_id"] == 0 or b["supplier_id"] == 0 or a["supplier_id"] != b["supplier_id"]
            for old_name, old_code in sides:
                if old_name == final_name and old_code == final_code:
                    continue
                if cross_supplier:
                    # Either side is the "any supplier" catalog (not tied to
                    # one supplier's POs), or this is a cross-supplier merge
                    # (e.g. an exact-duplicate "any supplier" placeholder
                    # merged into a real supplier's product) -- old_name/
                    # old_code items could be under any supplier's PO in
                    # that case, so rewrite matching items across all POs
                    # rather than restricting to just one supplier's.
                    if capture_undo:
                        rewritten_items += [
                            {"id": r["id"], "product": r["product"], "code": r["code"]}
                            for r in conn.execute(
                                f"SELECT id, product, code FROM po_items WHERE product=? AND {code_guard}",
                                (old_name, old_code, old_code),
                            )
                        ]
                    conn.execute(
                        f"UPDATE po_items SET product=?, code=? WHERE product=? AND {code_guard}",
                        (final_name, final_code, old_name, old_code, old_code),
                    )
                else:
                    if capture_undo:
                        rewritten_items += [
                            {"id": r["id"], "product": r["product"], "code": r["code"]}
                            for r in conn.execute(
                                f"SELECT id, product, code FROM po_items WHERE product=? AND {code_guard} "
                                "AND po_id IN (SELECT id FROM purchase_orders WHERE supplier_id=?)",
                                (old_name, old_code, old_code, a["supplier_id"]),
                            )
                        ]
                    conn.execute(
                        f"UPDATE po_items SET product=?, code=? WHERE product=? AND {code_guard} "
                        "AND po_id IN (SELECT id FROM purchase_orders WHERE supplier_id=?)",
                        (final_name, final_code, old_name, old_code, old_code, a["supplier_id"]),
                    )
        if capture_undo:
            ignored_pairs = [
                (r["product_id_a"], r["product_id_b"])
                for r in conn.execute(
                    "SELECT product_id_a, product_id_b FROM product_merge_ignored "
                    "WHERE product_id_a IN (?, ?) OR product_id_b IN (?, ?)",
                    (keep_id, remove_id, keep_id, remove_id),
                )
            ]
            undo_snapshot = {
                "keep_before": dict(a),
                "removed_before": dict(b),
                "rewritten_items": rewritten_items,
                "ignored_pairs": ignored_pairs,
            }
        # remove_id's row has to go FIRST, before keep_id is renamed to
        # final_name -- if final_name is remove_id's own current name (it
        # wins as "newer" more often than not), updating keep_id to that
        # name while remove_id's row still physically holds it collides
        # with the UNIQUE(supplier_id, name) constraint even though
        # remove_id is about to disappear. Found while chasing a report
        # that case-only duplicates ("Samsung" vs "SAMSUNG") weren't
        # actually auto-merging: the auto-merge pass's blanket `except
        # Exception: continue` was silently swallowing exactly this crash
        # whenever the more-recently-updated side of the pair happened to
        # be the one being removed, leaving the pair sitting there
        # unmerged with no visible error at all.
        conn.execute("DELETE FROM products WHERE id=?", (remove_id,))
        conn.execute(
            "UPDATE products SET name=?, code=?, last_price=?, times_ordered=?, updated_at=? WHERE id=?",
            (final_name, final_code, final_price, final_times, final_updated, keep_id),
        )
        conn.execute(
            "DELETE FROM product_merge_ignored WHERE product_id_a IN (?, ?) OR product_id_b IN (?, ?)",
            (keep_id, remove_id, keep_id, remove_id),
        )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    result = dict(conn.execute("SELECT * FROM products WHERE id=?", (keep_id,)).fetchone())
    if capture_undo:
        return result, undo_snapshot
    return result


def undo_product_merge(conn, snapshot):
    """Reverses exactly one merge_products() call using the undo_snapshot
    it returned (see merge_products' capture_undo) -- restores the removed
    product's row exactly as it was, restores the kept product's row to its
    pre-merge name/code/price/times-ordered, reverts every po_items row
    that got rewritten back to its original product/code, and restores any
    dismissed-duplicate pairs that were cleared as a side effect of the
    merge. Only meant to be called right after the matching merge (eg. a
    "Back" click in the same review session) -- if the kept or removed
    product has since been edited/merged again, those later changes are
    simply overwritten by this snapshot, same as any other undo.
    """
    keep_before = snapshot["keep_before"]
    removed_before = snapshot["removed_before"]
    cols = [c for c in removed_before.keys() if c != "id"]
    try:
        # See save_po's own comment on its equivalent guard (tidy-up pass,
        # post-Batch-151).
        if conn.in_transaction:
            conn.rollback()
        conn.execute("BEGIN")
        # Put the kept row back to its pre-merge name/code first -- if the
        # merge kept the removed product's exact name, inserting the
        # removed row back while the kept row still holds that name would
        # trip the UNIQUE(supplier_id, name) constraint.
        set_clause = ",".join(f"{c}=?" for c in cols)
        conn.execute(
            f"UPDATE products SET {set_clause} WHERE id=?",
            [keep_before[c] for c in cols] + [keep_before["id"]],
        )
        col_list = ",".join(["id"] + cols)
        placeholders = ",".join("?" * (len(cols) + 1))
        conn.execute(
            f"INSERT INTO products ({col_list}) VALUES ({placeholders})",
            [removed_before["id"]] + [removed_before[c] for c in cols],
        )
        for item in snapshot.get("rewritten_items", []):
            conn.execute(
                "UPDATE po_items SET product=?, code=? WHERE id=?",
                (item["product"], item["code"], item["id"]),
            )
        for pair in snapshot.get("ignored_pairs", []):
            # Batch 157: same ON CONFLICT portability swap as
            # ignore_duplicate_pair() above.
            conn.execute(
                "INSERT INTO product_merge_ignored(product_id_a, product_id_b) VALUES (?, ?) "
                "ON CONFLICT (product_id_a, product_id_b) DO NOTHING",
                pair,
            )
        conn.commit()
    except Exception:
        conn.rollback()
        raise


# ---- Duplicate supplier detection / merge (suggest-and-confirm only) ----
#
# Same approach as products: only ever suggests, and only ever acts on an
# explicit per-pair decision.

def find_duplicate_supplier_candidates(conn, min_ratio=0.6, limit=100):
    """Return candidate duplicate supplier pairs, best match first. Compares
    every supplier against every other (the list is normally small enough
    that this is cheap), excluding pairs already dismissed."""
    suppliers = [dict(r) for r in conn.execute("SELECT * FROM suppliers")]
    ignored = {
        (r["supplier_id_a"], r["supplier_id_b"])
        for r in conn.execute("SELECT * FROM supplier_merge_ignored")
    }
    candidates = []
    for i in range(len(suppliers)):
        for j in range(i + 1, len(suppliers)):
            a, b = suppliers[i], suppliers[j]
            pair = tuple(sorted((a["id"], b["id"])))
            if pair in ignored:
                continue
            na, nb = _normalize_product_name(a["company_name"]), _normalize_product_name(b["company_name"])
            if not na or not nb:
                continue
            ratio = difflib.SequenceMatcher(None, na, nb).ratio()
            if na in nb or nb in na:
                ratio = max(ratio, 0.85)
            if ratio >= min_ratio:
                candidates.append({
                    "supplier_id_a": a["id"], "supplier_id_b": b["id"],
                    "supplier_a": a, "supplier_b": b, "ratio": ratio,
                })
    candidates.sort(key=lambda c: c["ratio"], reverse=True)
    return candidates[:limit]


def find_best_matching_supplier_by_name(conn, name, exclude_id=None, min_ratio=0.6):
    """Fuzzy-matches one candidate company name against every existing
    supplier (same normalized-name + SequenceMatcher approach as
    find_duplicate_supplier_candidates), for a live "a similar supplier
    already exists" warning at the point a supplier is added or renamed --
    rather than only ever catching a near-duplicate afterwards via the
    batch Find Possible Duplicates scan. exclude_id skips a supplier's own
    existing row when checking an edit (so saving a supplier with its name
    unchanged never warns against itself). Returns the best-matching
    supplier dict (with a "match_ratio" key added), or None if nothing
    clears min_ratio."""
    name = (name or "").strip()
    if not name:
        return None
    na = _normalize_product_name(name)
    if not na:
        return None
    best_ratio, best_supplier = 0.0, None
    for s in list_suppliers(conn):
        if exclude_id is not None and s["id"] == exclude_id:
            continue
        ns = _normalize_product_name(s["company_name"])
        if not ns:
            continue
        if na == ns:
            ratio = 1.0
        else:
            ratio = difflib.SequenceMatcher(None, na, ns).ratio()
            if na in ns or ns in na:
                ratio = max(ratio, 0.85)
        if ratio > best_ratio:
            best_ratio, best_supplier = ratio, s
    if best_supplier is not None and best_ratio >= min_ratio:
        result = dict(best_supplier)
        result["match_ratio"] = best_ratio
        return result
    return None


def ignore_duplicate_supplier_pair(conn, id_a, id_b):
    lo, hi = sorted((id_a, id_b))
    # Batch 157: same ON CONFLICT portability swap as ignore_duplicate_pair()
    # above -- supplier_merge_ignored's own PRIMARY KEY is
    # (supplier_id_a, supplier_id_b).
    conn.execute(
        "INSERT INTO supplier_merge_ignored(supplier_id_a, supplier_id_b) VALUES (?, ?) "
        "ON CONFLICT (supplier_id_a, supplier_id_b) DO NOTHING",
        (lo, hi),
    )
    conn.commit()


# ---- Supplier naming consistency ("Clean up naming", suggest-and-confirm
# only, same review-before-acting approach as the duplicate finder above)
#
# Prompted directly by Yitzi's own feedback: "please keep supplier naming/
# formatting consistent throughout. For example, 'rvt' should be 'RVT',
# and I noticed one or two others where the capitalisation/naming isn't
# consistent." This never renames anything on its own -- it only ever
# suggests, via the Suppliers page's "Clean up naming" tool, and a rename
# is only ever applied once someone's ticked it and clicked Apply.

_SUPPLIER_NAME_MINOR_WORDS = {"and", "of", "the", "for", "&"}


def _normalized_supplier_word(word, i, acronym_set):
    """One word's casing decision for normalize_supplier_name -- kept
    deliberately conservative: a word only ever gets rewritten when it's
    unambiguously plain lowercase (or a bare initial), a recognised
    acronym, or already all caps. Anything with its own internal
    capitalisation already ("KoTech", "McDonald's", "iPhone") is left
    exactly as-is -- there's no reliable way to tell a deliberate brand
    style from a typo, so guessing would risk "fixing" something that was
    never broken."""
    letters = "".join(ch for ch in word if ch.isalpha())
    if not letters:
        return word
    if word.upper() in acronym_set:
        return word.upper()
    if len(letters) >= 2 and letters.isupper():
        return word.upper()  # already all caps -- treat as an intentional acronym
    if len(letters) <= 1:
        return word.upper()  # a bare initial reads the same either way
    if letters.islower():
        if i > 0 and word.lower() in _SUPPLIER_NAME_MINOR_WORDS:
            return word.lower()
        return word[:1].upper() + word[1:].lower()
    if word[:1].isupper() and letters[1:].islower():
        return word  # already properly cased (one leading capital) -- no-op
    return word  # internal/mixed caps -- ambiguous, leave untouched


def normalize_supplier_name(name, acronyms=None):
    """A smart, conservative title-case for supplier/company names:
    capitalizes each plain-lowercase word normally, keeps a small set of
    minor connector words lowercase (except as the first word), and
    renders a word fully uppercase when it's a recognised acronym --
    either already written in all caps in the source (so an already-
    correct name like "EGE" is never second-guessed) or listed in
    acronyms (so "rvt" becomes "RVT" once that's been told to the app --
    there's no way to tell a lowercase word is actually meant to be an
    acronym just by looking at it). A word that already carries its own
    internal capitalisation (see _normalized_supplier_word) is left
    untouched rather than guessed at. Returns name unchanged if it's
    blank. Never called automatically -- see suggest_supplier_name_
    cleanup, which is the suggest-and-confirm tool that actually uses
    this."""
    name = (name or "").strip()
    if not name:
        return name
    acronym_set = {a.strip().upper() for a in (acronyms or []) if a.strip()}
    words = name.split(" ")
    return " ".join(
        _normalized_supplier_word(word, i, acronym_set) if word else word
        for i, word in enumerate(words)
    )


def suggest_supplier_name_cleanup(conn, acronyms=None):
    """Every supplier whose name would change under normalize_supplier_
    name, old -> new -- the suggest-and-confirm list behind the Suppliers
    page's "Clean up naming" tool. acronyms defaults to the saved
    supplier_name_acronyms setting (a comma-separated list, editable from
    the tool itself)."""
    if acronyms is None:
        acronyms = [a.strip() for a in get_setting(conn, "supplier_name_acronyms", "").split(",") if a.strip()]
    suggestions = []
    for s in list_suppliers(conn):
        suggested = normalize_supplier_name(s["company_name"], acronyms)
        if suggested and suggested != s["company_name"]:
            suggestions.append({"id": s["id"], "old_name": s["company_name"], "new_name": suggested})
    suggestions.sort(key=lambda r: r["old_name"].lower())
    return suggestions


def apply_supplier_name_cleanup(conn, accepted):
    """Applies a chosen subset of suggest_supplier_name_cleanup's
    suggestions (a list of {"id", "new_name"} dicts -- only the ones the
    user actually ticked) via the same rename path (and full history
    cascade) as any other supplier edit, replace_suppliers. Returns how
    many were actually applied."""
    if not accepted:
        return 0
    suppliers = {r["id"]: dict(r) for r in conn.execute("SELECT * FROM suppliers")}
    applied = 0
    for entry in accepted:
        row = suppliers.get(entry.get("id"))
        new_name = (entry.get("new_name") or "").strip()
        if not row or not new_name:
            continue
        row["company_name"] = new_name
        applied += 1
    if applied:
        replace_suppliers(conn, list(suppliers.values()))
    return applied


def merge_suppliers(conn, keep_id, remove_id, merged_name=None, rewrite_history=True):
    """Merge remove_id into keep_id: reassigns that supplier's products and
    purchase orders, updates any remembered Zoho vendor mappings so future
    imports keep working, and deletes the removed supplier. If
    rewrite_history is True (the default -- unlike product merges, this
    matters here because Reports groups by the supplier name stored on each
    PO, not just by supplier_id), past POs under the removed supplier's name
    are relabelled with the kept supplier's name so spend reports read as
    one consistent supplier going forward."""
    a = conn.execute("SELECT * FROM suppliers WHERE id=?", (keep_id,)).fetchone()
    b = conn.execute("SELECT * FROM suppliers WHERE id=?", (remove_id,)).fetchone()
    if not a or not b:
        raise ValueError("One or both suppliers no longer exist")
    a, b = dict(a), dict(b)
    final_name = merged_name or a["company_name"]

    try:
        # See save_po's own comment on its equivalent guard (tidy-up pass,
        # post-Batch-151).
        if conn.in_transaction:
            conn.rollback()
        conn.execute("BEGIN")
        if rewrite_history:
            # _cascade_supplier_name_change matches on "supplier_id=keep_id
            # OR supplier_company_name=old_name", so passing the removed
            # supplier's own old name here covers both directions in one
            # go: the removed supplier's history (matched by remove_id,
            # already reassigned above, or its old name text) AND the kept
            # supplier's own prior history (matched by supplier_id=keep_id
            # unconditionally) -- so even if final_name is the removed
            # supplier's name rather than the kept one's, the kept
            # supplier's past orders/replies get relabelled to match too,
            # not just the removed side.
            _cascade_supplier_name_change(conn, keep_id, b["company_name"], final_name)
        else:
            conn.execute("UPDATE purchase_orders SET supplier_id=? WHERE supplier_id=?", (keep_id, remove_id))
            conn.execute("UPDATE price_request_replies SET supplier_id=? WHERE supplier_id=?", (keep_id, remove_id))

        for p in conn.execute("SELECT id, name FROM products WHERE supplier_id=?", (remove_id,)).fetchall():
            clash = conn.execute(
                "SELECT id FROM products WHERE supplier_id=? AND name=?", (keep_id, p["name"])
            ).fetchone()
            if clash:
                # Genuinely the same product already exists under the kept
                # supplier -- drop the redundant copy rather than fail.
                conn.execute("DELETE FROM products WHERE id=?", (p["id"],))
            else:
                conn.execute("UPDATE products SET supplier_id=? WHERE id=?", (keep_id, p["id"]))

        conn.execute("UPDATE zoho_vendor_map SET supplier_id=? WHERE supplier_id=?", (keep_id, remove_id))

        conn.execute(
            "UPDATE suppliers SET company_name=?, contact_name=COALESCE(NULLIF(contact_name,''), ?), "
            "email=COALESCE(NULLIF(email,''), ?), cc_emails=COALESCE(NULLIF(cc_emails,''), ?), "
            "phone=COALESCE(NULLIF(phone,''), ?), address=COALESCE(NULLIF(address,''), ?) WHERE id=?",
            (final_name, b["contact_name"], b["email"], b["cc_emails"], b["phone"], b["address"], keep_id),
        )
        conn.execute("DELETE FROM suppliers WHERE id=?", (remove_id,))
        conn.execute(
            "DELETE FROM supplier_merge_ignored WHERE supplier_id_a IN (?, ?) OR supplier_id_b IN (?, ?)",
            (keep_id, remove_id, keep_id, remove_id),
        )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    return dict(conn.execute("SELECT * FROM suppliers WHERE id=?", (keep_id,)).fetchone())


# ============================================================
# 9b. Supplier scorecard -- order history, spend trend, price trend,
#     and a simple issue log per supplier (matched by company_name, the
#     same denormalized field Reports and PO history already group by).
# ============================================================

def list_supplier_issues(conn, company_name):
    return [dict(r) for r in conn.execute(
        "SELECT * FROM supplier_issues WHERE supplier_company_name=? ORDER BY at DESC",
        (company_name,),
    )]


def list_supplier_issues_in_range(conn, date_from, date_to):
    """Every supplier issue logged within date_from..date_to (by when it
    was logged, not any date mentioned in the note), newest first --
    used to pull recently-logged issues straight into the periodic PO
    summary report instead of relying on someone remembering to type
    them in there by hand."""
    rows = conn.execute(
        "SELECT * FROM supplier_issues WHERE datetime(at) >= datetime(?) AND datetime(at) <= datetime(?) "
        "ORDER BY at DESC",
        (date_from.isoformat(), date_to.isoformat()),
    ).fetchall()
    return [dict(r) for r in rows]


def _fallback_issue_subject(note):
    """Used when an issue is logged with no explicit subject -- a short
    lead-in from the note itself, kept safely under the length the
    scorecard PDF's fixed-width table truncates at (see build_supplier_
    scorecard_pdf_pages), so old issues (logged before the subject field
    existed) still display sensibly instead of getting a mid-word cut."""
    note = note or ""
    return (note[:57] + "...") if len(note) > 60 else note


def add_supplier_issue(conn, company_name, note, category="", subject="", at=None):
    note = (note or "").strip()
    subject = (subject or "").strip() or _fallback_issue_subject(note)
    if not note and not subject:
        return
    conn.execute(
        "INSERT INTO supplier_issues(supplier_company_name, note, category, subject, at) VALUES (?, ?, ?, ?, ?)",
        (company_name, note, (category or "").strip(), subject, at or now_iso()),
    )
    conn.commit()


def update_supplier_issue(conn, issue_id, note=None, category=None, subject=None):
    """Edits a previously-logged issue -- only fields actually passed
    (not None) are changed. A blank subject falls back the same way
    add_supplier_issue does, so clearing it doesn't leave the PDF table
    with nothing to show for that row."""
    row = conn.execute("SELECT * FROM supplier_issues WHERE id=?", (issue_id,)).fetchone()
    if not row:
        return
    new_note = row["note"] if note is None else note.strip()
    new_category = row["category"] if category is None else category.strip()
    new_subject = (row["subject"] if subject is None else subject.strip()) or _fallback_issue_subject(new_note)
    conn.execute(
        "UPDATE supplier_issues SET note=?, category=?, subject=? WHERE id=?",
        (new_note, new_category, new_subject, issue_id),
    )
    conn.commit()


def delete_supplier_issue(conn, issue_id):
    conn.execute("DELETE FROM supplier_issues WHERE id=?", (issue_id,))
    conn.commit()


# ---- supplier issue categories (user-editable, same idea as savings
# categories below) ----

def list_supplier_issue_categories(conn, active_only=True):
    q = "SELECT * FROM supplier_issue_categories"
    if active_only:
        q += " WHERE active=1"
    q += " ORDER BY sort_order, id"
    return [dict(r) for r in conn.execute(q)]


def add_supplier_issue_category(conn, name):
    name = (name or "").strip()
    if not name:
        raise ValueError("Enter a category name.")
    existing = conn.execute(
        "SELECT id, active FROM supplier_issue_categories WHERE name=? COLLATE NOCASE", (name,)
    ).fetchone()
    if existing:
        if not existing["active"]:
            conn.execute("UPDATE supplier_issue_categories SET active=1 WHERE id=?", (existing["id"],))
            conn.commit()
        return
    max_order = conn.execute(
        "SELECT COALESCE(MAX(sort_order), -1) AS m FROM supplier_issue_categories"
    ).fetchone()["m"]
    conn.execute(
        "INSERT INTO supplier_issue_categories(name, sort_order, active) VALUES (?, ?, 1)", (name, max_order + 1)
    )
    conn.commit()


def rename_supplier_issue_category(conn, category_id, new_name):
    new_name = (new_name or "").strip()
    if not new_name:
        raise ValueError("Enter a category name.")
    row = conn.execute("SELECT name FROM supplier_issue_categories WHERE id=?", (category_id,)).fetchone()
    if not row:
        return
    old_name = row["name"]
    conn.execute("UPDATE supplier_issue_categories SET name=? WHERE id=?", (new_name, category_id))
    conn.execute("UPDATE supplier_issues SET category=? WHERE category=?", (new_name, old_name))
    conn.commit()


def set_supplier_issue_category_active(conn, category_id, active):
    conn.execute("UPDATE supplier_issue_categories SET active=? WHERE id=?", (1 if active else 0, category_id))
    conn.commit()


def get_supplier_price_competitiveness(conn, company_name, months=None):
    """For every product this supplier has sold us, compares their most
    recently paid price against the best (lowest) price any other
    supplier has sold the exact same product for, within the configured
    window -- built entirely from order history already in the app, no
    extra data entry needed. Returns the per-product comparison plus how
    often this supplier came out cheapest."""
    if months is None:
        try:
            months = int(get_setting(conn, "scorecard_price_window_months", "12") or 12)
        except ValueError:
            months = 12
    cutoff = (datetime.now() - timedelta(days=max(months, 1) * 30)).isoformat()
    rows = conn.execute(
        # pi.price * po.fx_rate converts each price into the home currency
        # before comparing -- otherwise a supplier billing in a different
        # currency could look artificially cheaper or dearer than one billing
        # in the home currency, just from the raw numbers not being on the
        # same scale.
        "SELECT pi.product AS product, po.supplier_company_name AS supplier, pi.price * po.fx_rate AS price, "
        "po.created_at AS created_at, po.po_ref AS po_ref FROM po_items pi JOIN purchase_orders po ON po.id = pi.po_id "
        "WHERE po.deleted = 0 AND po.created_at >= ? AND po.supplier_company_name != '' "
        "ORDER BY po.created_at DESC",
        (cutoff,),
    ).fetchall()
    latest = {}
    for r in rows:
        key = (r["product"], r["supplier"])
        price = float(r["price"] or 0)
        # rows are newest-first, so the first hit per key is normally the
        # latest price -- but a £0 row (a historical purchase with no
        # price ever recorded) is skipped over so this lands on the most
        # recent row that actually has a real price instead. Without this,
        # a supplier's most recent order for a product happening to be an
        # unrecorded-price import could make them look either impossibly
        # cheap (0 <= any other price) or wildly overpriced by comparison,
        # neither of which reflects a real price at all.
        if key not in latest and price > 0:
            latest[key] = {
                "price": price, "created_at": r["created_at"], "po_ref": r["po_ref"],
            }
    by_product = {}
    for (product, supplier), entry in latest.items():
        by_product.setdefault(product, {})[supplier] = entry

    compared = []
    cheapest_count = 0
    for product, by_supplier in by_product.items():
        if company_name not in by_supplier or len(by_supplier) < 2:
            continue
        my_price = by_supplier[company_name]["price"]
        others = [(s, e) for s, e in by_supplier.items() if s != company_name]
        best_other_supplier, best_other_entry = min(others, key=lambda se: se[1]["price"])
        best_other = best_other_entry["price"]
        is_cheapest = my_price <= best_other
        if is_cheapest:
            cheapest_count += 1
        compared.append({
            "product": product,
            "my_price": my_price,
            "best_other_price": best_other,
            # who that best-other-price actually came from, and which order,
            # so the comparison isn't just a bare number -- Yitzi can jump
            # straight to checking it rather than having to go hunting.
            "best_other_supplier": best_other_supplier,
            "best_other_po_ref": best_other_entry["po_ref"],
            "best_other_at": best_other_entry["created_at"],
            "difference": my_price - best_other,
            "is_cheapest": is_cheapest,
        })
    compared.sort(key=lambda r: r["difference"], reverse=True)  # biggest overpay first -- most actionable
    compared_count = len(compared)
    pct_cheapest = round(cheapest_count / compared_count * 100, 1) if compared_count else None
    return {
        "rows": compared,
        "compared_count": compared_count,
        "cheapest_count": cheapest_count,
        "pct_cheapest": pct_cheapest,
        "window_months": months,
    }


def get_supplier_scorecard(conn, company_name):
    """Everything shown on a supplier's scorecard: order count/spend
    summary, a monthly spend trend, the products bought from them most
    (with quantity/spend so a price trend is visible), the full order
    history, any issues logged against them (with category counts), how
    their prices on shared products compare to other suppliers, and any
    savings recorded against them."""
    # Summary/monthly/top_products all sum base_total (or price * fx_rate)
    # -- the home-currency equivalent -- so a supplier billing in a
    # different currency doesn't distort these totals. The order-by-order
    # list keeps each PO's own raw total + currency, since that's the
    # actual amount that order was raised for.
    summary = conn.execute(
        "SELECT COUNT(*) AS po_count, COALESCE(SUM(base_total), 0) AS total_spend, "
        "COALESCE(AVG(base_total), 0) AS avg_po_value, MIN(created_at) AS first_order_at, "
        "MAX(created_at) AS last_order_at FROM purchase_orders "
        "WHERE deleted=0 AND supplier_company_name=?",
        (company_name,),
    ).fetchone()

    monthly = [dict(r) for r in conn.execute(
        "SELECT strftime('%Y-%m', created_at) AS period, COALESCE(SUM(base_total), 0) AS total, "
        "COUNT(*) AS po_count FROM purchase_orders "
        "WHERE deleted=0 AND supplier_company_name=? GROUP BY period ORDER BY period",
        (company_name,),
    )]

    orders = [dict(r) for r in conn.execute(
        "SELECT po_ref, status, total, currency, created_at, updated_at FROM purchase_orders "
        "WHERE deleted=0 AND supplier_company_name=? ORDER BY created_at DESC",
        (company_name,),
    )]

    top_products = [dict(r) for r in conn.execute(
        "SELECT i.product AS product, SUM(i.qty) AS total_qty, SUM(i.qty * i.price * po.fx_rate) AS total_spend, "
        "COUNT(DISTINCT po.id) AS times_ordered, AVG(i.price * po.fx_rate) AS avg_price "
        "FROM po_items i JOIN purchase_orders po ON i.po_id = po.id "
        "WHERE po.deleted=0 AND po.supplier_company_name=? "
        "GROUP BY i.product ORDER BY total_spend DESC LIMIT 10",
        (company_name,),
    )]

    issues = list_supplier_issues(conn, company_name)
    issue_counts = {}
    for it in issues:
        cat = it.get("category") or "Uncategorised"
        issue_counts[cat] = issue_counts.get(cat, 0) + 1

    savings_total = conn.execute(
        "SELECT COALESCE(SUM(amount), 0) AS t FROM savings WHERE deleted=0 AND supplier_company_name=?",
        (company_name,),
    ).fetchone()["t"]

    return {
        "po_count": summary["po_count"] or 0,
        "total_spend": float(summary["total_spend"] or 0),
        "avg_po_value": float(summary["avg_po_value"] or 0),
        "first_order_at": summary["first_order_at"] or "",
        "last_order_at": summary["last_order_at"] or "",
        "monthly": monthly,
        "orders": orders,
        "top_products": top_products,
        "issues": issues,
        "issue_counts": sorted(issue_counts.items(), key=lambda kv: -kv[1]),
        "price_competitiveness": get_supplier_price_competitiveness(conn, company_name),
        "savings_total": float(savings_total or 0),
    }


def get_all_suppliers_scorecard_summary(conn, active_only=True):
    """One row per supplier for the side-by-side Supplier Scorecards
    comparison page -- spend, order count, issue count, and how often
    they're the cheapest option on shared products, all pulled from the
    same data as the individual scorecards rather than re-derived."""
    out = []
    for supplier in list_suppliers(conn, active_only=active_only):
        name = supplier["company_name"]
        summary = conn.execute(
            "SELECT COUNT(*) AS po_count, COALESCE(SUM(base_total), 0) AS total_spend "
            "FROM purchase_orders WHERE deleted=0 AND supplier_company_name=?",
            (name,),
        ).fetchone()
        issue_count = conn.execute(
            "SELECT COUNT(*) AS c FROM supplier_issues WHERE supplier_company_name=?", (name,)
        ).fetchone()["c"]
        price = get_supplier_price_competitiveness(conn, name)
        savings_total = conn.execute(
            "SELECT COALESCE(SUM(amount), 0) AS t FROM savings WHERE deleted=0 AND supplier_company_name=?",
            (name,),
        ).fetchone()["t"]
        out.append({
            "company_name": name,
            "po_count": summary["po_count"] or 0,
            "total_spend": float(summary["total_spend"] or 0),
            "issue_count": issue_count,
            "pct_cheapest": price["pct_cheapest"],
            "compared_count": price["compared_count"],
            "savings_total": float(savings_total or 0),
        })
    return out


# ============================================================
# 9c. Price requests (RFQ) -- "what's your best price on X?" sent to
#     several suppliers at once, tracked until each one replies, plus a
#     price-comparison view combining those replies with what's actually
#     been paid before (from the product catalog).
# ============================================================

def ensure_product_in_catalog(conn, name):
    """Makes sure `name` shows up in Product Catalog even though it's never
    actually been bought yet -- Yitzi: "any new product i put a request in
    for needs to be saved to product catalog to be able to show up there."
    Called whenever a price request is created or its product is renamed.

    Matches case/spacing-insensitively (_normalize_product_name, the same
    comparison the Batch 65 "Samsung/SAMSUNG" duplicate fix relies on)
    against every product already on file, under any supplier -- so this
    never creates a redundant near-duplicate catalog row for a product
    that's already there under slightly different casing or punctuation;
    it only adds a genuinely new one. A new row is filed under
    supplier_id=0, the same "(any supplier)" convention used elsewhere in
    this app for a product not yet tied to one particular supplier, with
    no price and times_ordered left at 0 -- asking what something costs
    isn't the same as having actually bought it, so it shouldn't look like
    a real purchase happened."""
    name = (name or "").strip()
    if not name:
        return
    for r in conn.execute("SELECT name FROM products"):
        if _product_names_match(r["name"], name):
            return
    upsert_product_manual(conn, None, 0, "", name, None, commit=False)


def create_price_request(conn, product_name, qty, notes, supplier_ids):
    """Creates a price request for product_name/qty and one reply row per
    supplier_id (status 'pending'). Returns (request_id, replies) where
    replies is a list of dicts with the supplier's name/email attached, for
    the caller (the UI) to actually open an email draft per supplier and
    then call mark_price_reply_sent for each one that was opened OK."""
    product_name = (product_name or "").strip()
    if not product_name:
        raise ValueError("Enter a product to ask about.")
    if not supplier_ids:
        raise ValueError("Choose at least one supplier to ask.")
    now = now_iso()
    cur = conn.execute(
        "INSERT INTO price_requests(product_name, qty, notes, created_at) VALUES (?, ?, ?, ?)",
        (product_name, qty or 1, notes or "", now),
    )
    request_id = cur.lastrowid
    ensure_product_in_catalog(conn, product_name)
    replies = []
    for sid in supplier_ids:
        srow = conn.execute("SELECT * FROM suppliers WHERE id=?", (sid,)).fetchone()
        if not srow:
            continue
        supplier = dict(srow)
        rcur = conn.execute(
            "INSERT INTO price_request_replies"
            "(request_id, supplier_id, supplier_company_name, status, notes) "
            "VALUES (?, ?, ?, 'pending', '')",
            (request_id, supplier["id"], supplier["company_name"]),
        )
        replies.append({
            "id": rcur.lastrowid,
            "request_id": request_id,
            "supplier_id": supplier["id"],
            "supplier_company_name": supplier["company_name"],
            "supplier_email": supplier.get("email", ""),
            "supplier_contact_name": supplier.get("contact_name", ""),
            "status": "pending",
        })
    conn.commit()
    return request_id, replies


def add_suppliers_to_price_request(conn, request_id, supplier_ids):
    """Adds one reply row (status 'pending') per new supplier_id to an
    already-existing price request -- for asking a supplier you didn't
    think of, or forgot to tick, the first time round. Suppliers already
    on the request are skipped rather than duplicated. Returns the list of
    newly-added replies, same shape as create_price_request's, so the
    caller can open email drafts for them the same way."""
    req = conn.execute("SELECT * FROM price_requests WHERE id=?", (request_id,)).fetchone()
    if not req:
        raise ValueError("This price request no longer exists.")
    already = {
        r["supplier_id"] for r in conn.execute(
            "SELECT supplier_id FROM price_request_replies WHERE request_id=?", (request_id,)
        )
    }
    replies = []
    for sid in supplier_ids:
        if sid in already:
            continue
        srow = conn.execute("SELECT * FROM suppliers WHERE id=?", (sid,)).fetchone()
        if not srow:
            continue
        supplier = dict(srow)
        rcur = conn.execute(
            "INSERT INTO price_request_replies"
            "(request_id, supplier_id, supplier_company_name, status, notes) "
            "VALUES (?, ?, ?, 'pending', '')",
            (request_id, supplier["id"], supplier["company_name"]),
        )
        replies.append({
            "id": rcur.lastrowid,
            "request_id": request_id,
            "supplier_id": supplier["id"],
            "supplier_company_name": supplier["company_name"],
            "supplier_email": supplier.get("email", ""),
            "supplier_contact_name": supplier.get("contact_name", ""),
            "status": "pending",
        })
        already.add(sid)
    conn.commit()
    return replies


def mark_price_reply_sent(conn, reply_id):
    conn.execute(
        "UPDATE price_request_replies SET sent_at=? WHERE id=?",
        (now_iso(), reply_id),
    )
    conn.commit()


def list_price_requests(conn):
    """All price requests, newest first, each with a received/pending/total
    count of the suppliers asked."""
    requests = [dict(r) for r in conn.execute("SELECT * FROM price_requests ORDER BY created_at DESC")]
    for req in requests:
        counts = conn.execute(
            "SELECT status, COUNT(*) AS c FROM price_request_replies WHERE request_id=? GROUP BY status",
            (req["id"],),
        ).fetchall()
        by_status = {c["status"]: c["c"] for c in counts}
        req["total_count"] = sum(by_status.values())
        req["received_count"] = by_status.get("received", 0)
        req["pending_count"] = by_status.get("pending", 0)
        req["declined_count"] = by_status.get("declined", 0)
    return requests


def get_price_request(conn, request_id):
    row = conn.execute("SELECT * FROM price_requests WHERE id=?", (request_id,)).fetchone()
    if not row:
        return None
    req = dict(row)
    req["replies"] = [
        dict(r) for r in conn.execute(
            "SELECT * FROM price_request_replies WHERE request_id=? "
            "ORDER BY (price IS NULL), price ASC, supplier_company_name COLLATE NOCASE",
            (request_id,),
        )
    ]
    return req


def record_price_reply(conn, reply_id, price, notes=""):
    """Enters the price a supplier came back with -- marks that reply
    'received'."""
    conn.execute(
        "UPDATE price_request_replies SET status='received', price=?, notes=?, replied_at=? WHERE id=?",
        (price, notes or "", now_iso(), reply_id),
    )
    conn.commit()


def mark_price_reply_declined(conn, reply_id, notes=""):
    conn.execute(
        "UPDATE price_request_replies SET status='declined', notes=?, replied_at=? WHERE id=?",
        (notes or "", now_iso(), reply_id),
    )
    conn.commit()


def reopen_price_reply(conn, reply_id):
    """Undoes a received/declined mark -- back to waiting on this supplier."""
    conn.execute(
        "UPDATE price_request_replies SET status='pending', price=NULL, replied_at=NULL WHERE id=?",
        (reply_id,),
    )
    conn.commit()


def list_suppliers_with_pending_price_requests(conn):
    """Batch 166 -- Yitzi: "staff will ask a lot of times for multiple
    items at once ... so I'm sending out multiple pricing to suppliers at
    once ... the next part is when I receive an email back from suppliers
    with all the pricing I want to be able to fill them all in quickly
    ... currently I need to go by item and find the supplier inside the
    item but would be easier if I could do by supplier as well." One
    supplier's reply email often covers several different price requests
    (e.g. iPhone 17 256GB, iPhone 17 512GB, iPhone 16 256GB all asked in
    separate requests) -- this is the supplier picker behind
    EnterPricesBySupplierDialog, listing only suppliers who actually have
    something outstanding to enter, each with how many, so there's
    nothing to hunt through for a supplier who's already fully answered."""
    rows = conn.execute(
        "SELECT prr.supplier_id, prr.supplier_company_name, COUNT(*) AS pending_count "
        "FROM price_request_replies prr "
        "WHERE prr.status='pending' AND prr.supplier_id IS NOT NULL "
        "GROUP BY prr.supplier_id, prr.supplier_company_name "
        "ORDER BY prr.supplier_company_name COLLATE NOCASE"
    )
    return [dict(r) for r in rows]


def list_price_replies_by_supplier(conn, supplier_id):
    """Batch 166 -- the actual "fill them all in by supplier" list: every
    price request reply tied to this one supplier, across every separate
    price request they've ever been asked on (not just one product at a
    time, the way the existing per-product PriceRequestDetailDialog
    works), each carrying its own product_name/qty/request_id so a price
    can be entered against the right one. Pending first (needs action),
    then received, then declined, each group newest-asked first -- a
    still-pending item is always the reason this screen is being opened,
    so it should never be buried under ones already dealt with.

    Returns a list of dicts merging price_request_replies' own columns
    with product_name/qty/request created_at from the price_requests row
    each belongs to (aliased request_created_at to avoid clashing with
    the reply's own created_at-less schema, which uses sent_at/replied_at
    instead)."""
    rows = conn.execute(
        "SELECT prr.*, pr.product_name, pr.qty, pr.created_at AS request_created_at "
        "FROM price_request_replies prr "
        "JOIN price_requests pr ON pr.id = prr.request_id "
        "WHERE prr.supplier_id=? "
        "ORDER BY (prr.status != 'pending'), (prr.status = 'declined'), pr.created_at DESC",
        (supplier_id,),
    )
    return [dict(r) for r in rows]


def delete_price_request(conn, request_id):
    conn.execute("DELETE FROM price_requests WHERE id=?", (request_id,))
    conn.commit()


def update_price_request_notes(conn, request_id, notes):
    """Adds or edits the note on an already-created price request --
    notes could previously only be set once, at the moment of sending a
    New Price Request, with no way back in afterward (Yitzi: "I want to be
    able to add a note to price request it should be visible when
    opening"). This is a plain notes-only update -- product/qty/suppliers
    are unaffected."""
    conn.execute("UPDATE price_requests SET notes=? WHERE id=?", ((notes or "").strip(), request_id))
    conn.commit()


def rename_price_request_product(conn, request_id, new_name):
    """Fixes a price request's product name after the fact -- Yitzi: "after
    doing a price request i want to be able to change the name of the
    product incase it was entered incorrectly." Only this one request's
    product_name is changed; if the same typo was made on a separate
    request too, that one needs fixing on its own the same way, since
    there's no shared "product" record underneath price requests the way
    there is for purchase orders (get_product_price_request_history groups
    purely by product_name text). The corrected name is also made sure to
    exist in the product catalog (ensure_product_in_catalog), same as a
    brand new request already gets -- otherwise a renamed request could
    vanish from what Product Catalog can find it under. Nothing about the
    request's suppliers, replies, qty, or notes is touched."""
    new_name = (new_name or "").strip()
    if not new_name:
        raise ValueError("Enter a product name.")
    req = conn.execute("SELECT * FROM price_requests WHERE id=?", (request_id,)).fetchone()
    if not req:
        raise ValueError("This price request no longer exists.")
    conn.execute("UPDATE price_requests SET product_name=? WHERE id=?", (new_name, request_id))
    ensure_product_in_catalog(conn, new_name)
    conn.commit()


def get_product_price_request_history(conn, product_name):
    """Every price request reply ever logged for this exact product name,
    across every request and every supplier asked, newest request first --
    one row per supplier asked, each carrying its own request_id/request
    date, qty, notes, status, price, and (once replied) the date it
    actually came back. Requested by Yitzi so a product's price-request
    history shows up right on its own Product Catalog entry ("maybe
    another tab after scorecard that track all price requests for that
    product and the date if there are many [quotes] on different dates") --
    a product genuinely can be asked about more than once, sometimes months
    apart or from different suppliers each time, and this is meant to make
    that whole timeline visible in one place rather than needing to hunt
    through the separate Price Requests page one request at a time.
    Matches by exact product_name text, same convention as
    get_product_purchase_history."""
    name = (product_name or "").strip()
    if not name:
        return []
    rows = [dict(r) for r in conn.execute(
        "SELECT pr.id AS request_id, pr.qty AS qty, pr.notes AS notes, pr.created_at AS requested_at, "
        "       prr.id AS reply_id, prr.supplier_company_name AS supplier_company_name, "
        "       prr.status AS status, prr.price AS price, prr.replied_at AS replied_at "
        "FROM price_requests pr "
        "JOIN price_request_replies prr ON prr.request_id = pr.id "
        "WHERE pr.product_name = ? "
        "ORDER BY pr.created_at DESC, prr.supplier_company_name COLLATE NOCASE",
        (name,),
    ).fetchall()]

    # Flag which received quote is each supplier's CURRENT one -- the most
    # recently replied_at -- vs an older one that's since been superseded.
    # Direct spec from Yitzi: "if Corptel give me a price, and then six
    # months later they give me a price again, slightly more expensive...
    # once I receive a new price from a supplier, the older one should
    # expire, I'll still be showing in the history, meaning the cheapest
    # price won't be worked out of an old quote." An expired quote is never
    # deleted or hidden -- every row stays visible here -- but only each
    # supplier's latest received price should ever feed into a "cheapest"
    # figure. Pending/declined rows aren't a price at all, so they're left
    # with is_current_quote=None (not applicable) rather than True/False.
    latest_replied_at = {}
    for r in rows:
        if r["status"] == "received" and r.get("price") is not None:
            name_key = r["supplier_company_name"] or ""
            if name_key not in latest_replied_at or (r["replied_at"] or "") > latest_replied_at[name_key]:
                latest_replied_at[name_key] = r["replied_at"] or ""
    for r in rows:
        if r["status"] == "received" and r.get("price") is not None:
            name_key = r["supplier_company_name"] or ""
            r["is_current_quote"] = (r["replied_at"] or "") == latest_replied_at.get(name_key)
        else:
            r["is_current_quote"] = None
    return rows


def build_price_request_email(conn, supplier, product_name, qty):
    """Builds the subject/plain/html body for one supplier's price-request
    email from the Settings-configurable template. Placeholders:
    {supplier_name}, {product}, {qty}, {your_name}."""
    your_name = get_setting(conn, "your_name", "")
    contact = str(supplier.get("contact_name", "")).strip()
    company = str(supplier.get("company_name", "")).strip()
    supplier_name = first_name_of(contact) or company or "there"
    fields = {
        "supplier_name": supplier_name,
        "product": product_name,
        "qty": str(qty),
        "your_name": your_name,
    }
    subject_tmpl = get_setting(conn, "price_request_email_subject", DEFAULT_SETTINGS["price_request_email_subject"])
    body_tmpl = get_setting(conn, "price_request_email_body", DEFAULT_SETTINGS["price_request_email_body"])
    try:
        subject = subject_tmpl.format(**fields)
    except (KeyError, IndexError):
        subject = subject_tmpl
    try:
        plain = body_tmpl.format(**fields)
    except (KeyError, IndexError):
        plain = body_tmpl
    if your_name.strip() and "kind regards" not in plain.lower():
        plain += f"\n\nKind regards\n{your_name.strip()}"
    body_html = "".join(
        f'<div style="font-family:Arial,\'Segoe UI\',Helvetica,sans-serif;font-size:14px;color:#3a3164;'
        f'margin:14px 0 0 0;">{html_escape(line)}</div>'
        if line.strip() else '<div style="margin:10px 0 0 0;">&nbsp;</div>'
        for line in plain.splitlines()
    )
    html = email_shell_html(body_html, width=600)
    charts = {}
    logo_png = rcg_logo_png_bytes()
    if logo_png:
        charts[RCG_LOGO_CID] = logo_png
    return {"subject": subject, "plain": plain, "html": html, "charts": charts}


def build_price_request_email_multi(conn, supplier, items):
    """Same as build_price_request_email, but for asking one supplier about
    several products in a single email (each one is still tracked as its
    own separate price request underneath -- see NewPriceRequestDialog).
    items is a list of {"product_name", "qty"} dicts. Not driven by the
    Settings-configurable single-product template (its {product}/{qty}
    placeholders don't make sense for a list), but opens with the same
    "Hi {supplier_name}," / sign-off shape so it reads the same way."""
    your_name = get_setting(conn, "your_name", "")
    contact = str(supplier.get("contact_name", "")).strip()
    company = str(supplier.get("company_name", "")).strip()
    supplier_name = first_name_of(contact) or company or "there"

    subject = f"Pricing enquiry: {len(items)} items" if len(items) != 1 else f"Pricing enquiry: {items[0]['product_name']}"
    lines = [
        f"Hi {supplier_name},",
        "",
        "Do you have the following in stock, and if so what would be your best prices?",
        "",
    ]
    for item in items:
        lines.append(f"- {item['product_name']}  (qty {item['qty']})")
    lines += ["", "Please let me know at your earliest convenience."]
    plain = "\n".join(lines)
    if your_name.strip():
        plain += f"\n\nKind regards\n{your_name.strip()}"
    body_html = "".join(
        f'<div style="font-family:Arial,\'Segoe UI\',Helvetica,sans-serif;font-size:14px;color:#3a3164;'
        f'margin:14px 0 0 0;">{html_escape(line)}</div>'
        if line.strip() else '<div style="margin:10px 0 0 0;">&nbsp;</div>'
        for line in plain.splitlines()
    )
    html = email_shell_html(body_html, width=600)
    charts = {}
    logo_png = rcg_logo_png_bytes()
    if logo_png:
        charts[RCG_LOGO_CID] = logo_png
    return {"subject": subject, "plain": plain, "html": html, "charts": charts}


def get_last_and_lowest_price_paid(conn, product_name):
    """The two purchase-history figures the change request's Price Request
    screen was to show alongside newly-received quotes ("show last valid
    price paid [and] lowest valid price paid... make this information
    clear and easy to compare against the new prices received"), so a
    person can judge an incoming quote against what's actually been paid
    before, not just against the other suppliers asked this time.

    "Paid" means a real purchase -- an actual po_items line -- not a quote
    that was only ever asked for and never ordered; quotes are already
    shown separately in the request's own replies table. £0/missing
    prices are ignored for both figures, per Yitzi's own explicit
    instruction here, matching the same convention already applied
    everywhere else a "lowest"/"last" price is computed (batch 77:
    get_product_scorecard, get_price_comparison,
    get_supplier_price_competitiveness).

    Returns {"last": entry or None, "lowest": entry or None}, each entry a
    {"price", "supplier_company_name", "date"} dict -- None for either key
    if this product has never actually been bought at a real (>£0) price.
    """
    rows = [
        r for r in get_product_purchase_history(conn, product_name)
        if r.get("price") is not None and float(r["price"] or 0) > 0
    ]
    if not rows:
        return {"last": None, "lowest": None}
    # get_product_purchase_history is already ordered newest-first.
    last_row = rows[0]
    lowest_row = min(rows, key=lambda r: float(r["price"] or 0))

    def _entry(r):
        return {
            "price": float(r["price"] or 0),
            "supplier_company_name": r.get("supplier_company_name") or "",
            "date": r.get("created_at") or "",
        }

    return {"last": _entry(last_row), "lowest": _entry(lowest_row)}


def build_price_request_summary(conn, request_id):
    """Builds the content behind both the Price Request screen's "Copy to
    Clipboard" button and its "Open in Outlook" summary email -- direct
    spec: "allow me to quickly send staff the pricing that has been
    received... all pricing received, supplier information, product
    information, last price paid [and its supplier/date], lowest valid
    price paid [and its supplier/date], relevant price request
    information... the result must be easy to paste into an email and
    look presentable... green highlighting for lowest prices where
    applicable."

    Returns None if the request no longer exists. Otherwise a dict:
    {"product_name", "qty", "cheapest": reply-dict or None (the best of
    the NEW quotes actually received on THIS request -- every reply tied
    at that same price, not just the first found, per the same
    tied-lowest rule the catalogue view uses), "last_paid"/"lowest_paid":
    from get_last_and_lowest_price_paid, "plain": str, "html": str}.
    "html" is a ready-to-use fragment -- a bordered table with alternating
    row shading, a green highlight on every cheapest-received row, plain
    inline styles throughout (no external stylesheet, since this has to
    survive being pasted into an email client) -- meant to be handed to
    set_clipboard_html_windows (Copy to Clipboard) or as the body of an
    Outlook draft (Open in Outlook).
    """
    req = get_price_request(conn, request_id)
    if not req:
        return None
    replies = req["replies"]
    currency = get_setting(conn, "currency", "GBP")
    status_labels = {"pending": "Pending", "received": "Received", "declined": "Unavailable"}

    # Genuine £0 quotes are never a real price any more than a £0 purchase
    # is (batch 77's guard, applied here too) -- they're excluded from
    # "best offer" but still shown, unhighlighted, in the full table below.
    received = [
        r for r in replies
        if r["status"] == "received" and r.get("price") is not None and float(r["price"]) > 0
    ]
    cheapest_price = min((float(r["price"]) for r in received), default=None)
    price_context = get_last_and_lowest_price_paid(conn, req["product_name"])
    cheapest = next(
        (r for r in received if float(r["price"]) == cheapest_price), None
    ) if cheapest_price is not None else None

    def _is_cheapest(r):
        return cheapest_price is not None and r["status"] == "received" and r.get("price") is not None and float(r["price"]) == cheapest_price

    def esc(s):
        return html_escape(str(s or ""))

    # -- plain text (used for the mailto: fallback and the DataFormats
    #    UnicodeText half of the clipboard payload) --
    lines = [f"Pricing quote for {req['product_name']} (qty {req['qty']})", ""]
    for r in replies:
        label = status_labels.get(r["status"], (r["status"] or "").capitalize())
        price_txt = (
            money(r["price"], currency) if r.get("price") is not None
            else ("Not available" if r["status"] == "declined" else "-")
        )
        marker = "  <-- best offer" if _is_cheapest(r) else ""
        lines.append(f"- {r['supplier_company_name'] or '(no supplier)'}: {label}, {price_txt}{marker}")
    lines.append("")
    if cheapest is not None:
        lines.append(
            f"Best offer: {money(cheapest['price'], currency)} from "
            f"{cheapest['supplier_company_name'] or '(no supplier)'} "
            f"({(cheapest.get('replied_at') or '')[:10] or 'date not recorded'})"
        )
    if price_context["last"]:
        lp = price_context["last"]
        lines.append(
            f"Last price paid: {money(lp['price'], currency)} "
            f"({lp['supplier_company_name'] or '(no supplier)'}, {(lp['date'] or '')[:10]})"
        )
    if price_context["lowest"]:
        lo = price_context["lowest"]
        lines.append(
            f"Lowest price ever paid: {money(lo['price'], currency)} "
            f"({lo['supplier_company_name'] or '(no supplier)'}, {(lo['date'] or '')[:10]})"
        )
    plain = "\n".join(lines)

    # -- HTML table (headers, alternating rows, cheapest-row highlight).
    # This is a pasteable FRAGMENT, not a full email -- used both as the
    # "Copy to Clipboard" payload (pasted inline into whatever message the
    # user is already writing) and appended straight after a plain intro
    # sentence for "Open in Outlook" -- so it deliberately carries the RCG
    # colour palette and highlighting but no logo/header/footer chrome of
    # its own, which would look out of place stitched into someone else's
    # message. Colour is the ONLY signal for "best offer" besides the text
    # label, so red-quantity's <font color> + plain-CSS redundancy pattern
    # (build_html, above) is repeated here for the same reason: Outlook's
    # older Word-based clipboard-paste converter has a history of quietly
    # dropping style attributes it doesn't like. -- RCG_GREEN as a literal
    # hex, not a CSS variable, and no border-radius/box-shadow relied on for
    # meaning (a square instead of a rounded "BEST PRICE" tag is a harmless
    # cosmetic downgrade if the converter drops the radius; losing the
    # colour entirely would not be).
    rows_html = []
    for i, r in enumerate(replies):
        label = status_labels.get(r["status"], (r["status"] or "").capitalize())
        price_txt = (
            money(r["price"], currency) if r.get("price") is not None
            else ("Not available" if r["status"] == "declined" else "—")
        )
        is_cheapest = _is_cheapest(r)
        is_pending = r["status"] == "pending"
        if is_cheapest:
            bg, colour, weight = "#e3f7e8", RCG_GREEN, "bold"
        elif is_pending:
            bg, colour, weight = "#ffffff", RCG_MUTED, "normal"
        elif r["status"] == "declined":
            bg, colour, weight = "#fbeaea", RCG_RED, "normal"
        else:
            bg, colour, weight = (RCG_ROW_BLUE if i % 2 else "#ffffff"), "#3a3164", "normal"
        best_tag = (
            f'&nbsp;<span style="background:{RCG_GREEN};color:#ffffff;font-size:10.5px;font-weight:bold;'
            f'padding:2px 7px;border-radius:9px;">BEST PRICE</span>' if is_cheapest else ""
        )
        border = f"border-left:3px solid {RCG_GREEN};" if is_cheapest else ""
        rows_html.append(
            f'<tr style="background:{bg};">'
            f'<td style="padding:8px 10px;{border}border-bottom:1px solid {RCG_LINE};">'
            f'<font color="{colour}"><span style="color:{colour};font-weight:{weight};">'
            f'{esc(r["supplier_company_name"] or "(no supplier)")}</span></font></td>'
            f'<td style="padding:8px 10px;border-bottom:1px solid {RCG_LINE};">'
            f'<font color="{colour}"><span style="color:{colour};">{esc(label)}</span></font></td>'
            f'<td style="padding:8px 10px;border-bottom:1px solid {RCG_LINE};">'
            f'<font color="{colour}"><span style="color:{colour};font-weight:{weight};">'
            f'{esc(price_txt)}</span></font>{best_tag}</td>'
            f'<td style="padding:8px 10px;border-bottom:1px solid {RCG_LINE};">'
            f'<font color="{colour}"><span style="color:{colour};">'
            f'{esc((r.get("replied_at") or "")[:10] or "—")}</span></font></td>'
            f'</tr>'
        )
    context_cards = []
    if price_context["last"]:
        lp = price_context["last"]
        context_cards.append(
            f'<td width="49%" style="background:#f9f8fc;border-left:3px solid {RCG_ACCENT_LIGHT};'
            f'padding:12px 14px;"><div style="font-size:10.5px;letter-spacing:0.04em;text-transform:uppercase;'
            f'color:{RCG_MUTED};margin-bottom:4px;">Last price paid</div>'
            f'<div style="font-size:16px;font-weight:bold;color:{RCG_INK};">{esc(money(lp["price"], currency))}</div>'
            f'<div style="font-size:11.5px;color:{RCG_MUTED};margin-top:2px;">'
            f'{esc(lp["supplier_company_name"])} &middot; {esc((lp["date"] or "")[:10])}</div></td>'
        )
    if price_context["lowest"]:
        lo = price_context["lowest"]
        context_cards.append(
            f'<td width="49%" style="background:#f9f8fc;border-left:3px solid {RCG_GREEN};'
            f'padding:12px 14px;"><div style="font-size:10.5px;letter-spacing:0.04em;text-transform:uppercase;'
            f'color:{RCG_MUTED};margin-bottom:4px;">Lowest price ever paid</div>'
            f'<div style="font-size:16px;font-weight:bold;color:{RCG_GREEN};">'
            f'{esc(money(lo["price"], currency))}</div>'
            f'<div style="font-size:11.5px;color:{RCG_MUTED};margin-top:2px;">'
            f'{esc(lo["supplier_company_name"])} &middot; {esc((lo["date"] or "")[:10])}</div></td>'
        )
    context_html = ""
    if context_cards:
        gap = '<td width="2%" style="font-size:0;line-height:0;">&nbsp;</td>' if len(context_cards) > 1 else ""
        context_html = (
            '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" '
            f'style="margin-top:14px;"><tr>{context_cards[0]}{gap}'
            f'{context_cards[1] if len(context_cards) > 1 else ""}</tr></table>'
        )
    html = (
        '<div style="font-family:Arial,\'Segoe UI\',Helvetica,sans-serif;font-size:14px;color:#3a3164;">'
        f'<p style="margin:0 0 10px 0;"><b>Pricing quote for {esc(req["product_name"])}</b> '
        f'(qty {req["qty"]})</p>'
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" '
        'style="border-collapse:collapse;font-size:13px;">'
        f'<tr><th align="left" style="background:{RCG_INK};color:#ffffff;padding:8px 10px;font-size:11px;'
        'letter-spacing:0.04em;text-transform:uppercase;">Supplier</th>'
        f'<th align="left" style="background:{RCG_INK};color:#ffffff;padding:8px 10px;font-size:11px;'
        'letter-spacing:0.04em;text-transform:uppercase;">Status</th>'
        f'<th align="left" style="background:{RCG_INK};color:#ffffff;padding:8px 10px;font-size:11px;'
        'letter-spacing:0.04em;text-transform:uppercase;">Price</th>'
        f'<th align="left" style="background:{RCG_INK};color:#ffffff;padding:8px 10px;font-size:11px;'
        'letter-spacing:0.04em;text-transform:uppercase;">Date</th>'
        '</tr>'
        + "".join(rows_html) +
        '</table>'
        + context_html
        + '</div>'
    )

    return {
        "product_name": req["product_name"], "qty": req["qty"],
        "cheapest": cheapest, "last_paid": price_context["last"], "lowest_paid": price_context["lowest"],
        "plain": plain, "html": html,
    }


def build_combined_price_request_summary(conn, request_ids):
    """Batch 162 -- Yitzi: "if im getting a qoute for Iphone ... i want to
    be able to select multiplule product qoutes and ad it to one email and
    one PDF" so staff get one message covering several products' pricing
    instead of one email per product. Stitches several requests' own
    build_price_request_summary() output together -- each keeps its own
    full table and its own "BEST PRICE" highlight (there's no single "best"
    across different products), separated by a divider. request_ids may
    contain duplicates or ids for a request that's since been deleted;
    duplicates collapse to one copy (first occurrence order kept) and a
    missing request is silently skipped rather than failing the whole
    thing -- matches build_price_request_summary's own "returns None if
    this one no longer exists" behaviour, just at the collection level.

    Returns None if none of the given ids resolve to a request still on
    file. Otherwise a dict: {"product_names": [str, ...] (in request_ids'
    own order), "count": int, "subject": str, "intro": str, "plain": str,
    "html": str} -- "html"/"plain" are ready to hand to
    set_clipboard_html_windows or as an Outlook draft body, same as a
    single summary's own "html"/"plain".
    """
    seen_ids = []
    for rid in request_ids:
        if rid not in seen_ids:
            seen_ids.append(rid)
    summaries = []
    for rid in seen_ids:
        s = build_price_request_summary(conn, rid)
        if s is not None:
            summaries.append(s)
    if not summaries:
        return None

    product_names = [s["product_name"] for s in summaries]
    if len(product_names) <= 3:
        names_for_subject = ", ".join(product_names)
    else:
        names_for_subject = f"{', '.join(product_names[:2])} and {len(product_names) - 2} more"
    subject = f"Pricing quotes: {names_for_subject}"
    intro = (
        f"Please see below pricing quotes for {pluralize(len(summaries), 'product')}: "
        f"{', '.join(product_names)}."
    )

    plain = intro + "\n\n" + ("\n\n" + ("-" * 40) + "\n\n").join(s["plain"] for s in summaries)

    # Same "no external stylesheet, plain inline styles" rule as a single
    # summary's own html fragment -- this still has to survive being pasted
    # into Outlook or another email client. Each product keeps its own
    # "Pricing quote for X" heading (already part of its own summary["html"]),
    # so the divider's only job is telling the products apart at a glance.
    divider = (
        f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" '
        f'style="margin:22px 0;"><tr><td style="border-top:1px solid {RCG_LINE};font-size:0;'
        f'line-height:0;">&nbsp;</td></tr></table>'
    )
    html = (
        '<div style="font-family:Arial,\'Segoe UI\',Helvetica,sans-serif;font-size:14px;color:#3a3164;">'
        f'<p style="margin:0 0 14px 0;">{html_escape(intro)}</p></div>'
        + divider.join(s["html"] for s in summaries)
    )

    return {
        "product_names": product_names, "count": len(summaries),
        "subject": subject, "intro": intro, "plain": plain, "html": html,
    }


def get_price_comparison(conn, product_name):
    """Every price on record for this exact product name, from every
    supplier -- combining what's actually been paid before (products.
    last_price, from real orders) with any quotes gathered through a price
    request (price_request_replies where status='received'). If the same
    supplier shows up in both, only the more recent one is kept. Sorted
    cheapest first."""
    product_name = (product_name or "").strip()
    if not product_name:
        return []
    entries = {}  # supplier_company_name -> entry dict

    for row in conn.execute(
        "SELECT p.supplier_id, p.last_price, p.updated_at, s.company_name "
        "FROM products p LEFT JOIN suppliers s ON s.id = p.supplier_id "
        "WHERE p.name = ? AND p.supplier_id != 0",
        (product_name,),
    ):
        name = row["company_name"] or f"Supplier #{row['supplier_id']}"
        entries[name] = {
            "supplier_company_name": name,
            "price": float(row["last_price"] or 0),
            "source": "Last ordered",
            "source_date": row["updated_at"] or "",
        }

    for row in conn.execute(
        "SELECT prr.supplier_company_name, prr.price, prr.replied_at "
        "FROM price_request_replies prr "
        "JOIN price_requests pr ON pr.id = prr.request_id "
        "WHERE pr.product_name = ? AND prr.status = 'received' AND prr.price IS NOT NULL",
        (product_name,),
    ):
        name = row["supplier_company_name"] or "(unknown supplier)"
        candidate = {
            "supplier_company_name": name,
            "price": float(row["price"] or 0),
            "source": "Quote received",
            "source_date": row["replied_at"] or "",
        }
        existing = entries.get(name)
        if not existing or (candidate["source_date"] or "") >= (existing["source_date"] or ""):
            entries[name] = candidate

    result = sorted(entries.values(), key=lambda e: e["price"])
    # The cheapest VALID price, ignoring any £0/missing entry (an
    # unpriced product, or a historical record with no price ever
    # recorded) -- previously this only ever checked whether the single
    # lowest-sorting entry (index 0) had a real price, so a £0 entry
    # sorting to the very front (0 is always the minimum) meant NO entry
    # ever got marked cheapest, not even the genuinely lowest real price
    # sitting right behind it. Ties (more than one supplier at the exact
    # same lowest valid price) are all marked cheapest together.
    valid_prices = [e["price"] for e in result if e["price"] > 0]
    lowest_valid = min(valid_prices) if valid_prices else None
    for e in result:
        e["is_cheapest"] = lowest_valid is not None and e["price"] == lowest_valid
    return result


def get_price_history_for_product(conn, product_name, supplier_company_name=None):
    """Every price this product has ever been bought or quoted at, newest
    first -- unlike get_price_comparison (one row per supplier, latest
    price only), this keeps every past price point so a trend, or a
    supplier's price quietly creeping up over time, is visible instead of
    only ever seeing the single latest figure. Combines real purchase
    prices (every PO line item ever raised for it) with quoted prices
    (every price request reply ever received for it). Pass
    supplier_company_name to see just one supplier's history."""
    product_name = (product_name or "").strip()
    if not product_name:
        return []
    entries = []
    for row in get_product_purchase_history(conn, product_name):
        if supplier_company_name and row["supplier_company_name"] != supplier_company_name:
            continue
        entries.append({
            "date": row["created_at"] or "",
            "supplier_company_name": row["supplier_company_name"] or "",
            "price": float(row["price"] or 0),
            "source": "Ordered",
            "reference": row["po_ref"] or "",
        })
    for row in conn.execute(
        "SELECT prr.supplier_company_name, prr.price, prr.replied_at "
        "FROM price_request_replies prr JOIN price_requests pr ON pr.id = prr.request_id "
        "WHERE pr.product_name = ? AND prr.status = 'received' AND prr.price IS NOT NULL",
        (product_name,),
    ):
        name = row["supplier_company_name"] or "(unknown supplier)"
        if supplier_company_name and name != supplier_company_name:
            continue
        entries.append({
            "date": row["replied_at"] or "",
            "supplier_company_name": name,
            "price": float(row["price"] or 0),
            "source": "Quoted",
            "reference": "",
        })
    entries.sort(key=lambda e: e["date"] or "", reverse=True)
    return entries


def get_product_scorecard(conn, product_name):
    """Everything shown on a product's scorecard -- the same idea as
    get_supplier_scorecard, but for one product: order/spend summary,
    monthly spend trend, a per-supplier breakdown (who it's actually been
    bought from, and at what price), and the full price history (ordered +
    quoted, already used elsewhere for this product) for a price-over-time
    view."""
    product_name = (product_name or "").strip()
    rows = [dict(r) for r in conn.execute(
        "SELECT po.po_ref AS po_ref, po.supplier_company_name AS supplier_company_name, "
        "       po.created_at AS created_at, po.fx_rate AS fx_rate, "
        "       pi.qty AS qty, pi.price AS price, pi.code AS code "
        "FROM po_items pi JOIN purchase_orders po ON po.id = pi.po_id "
        "WHERE pi.product = ? AND po.deleted = 0 "
        "ORDER BY po.created_at DESC, po.id DESC",
        (product_name,),
    )]

    def line_spend(r):
        return float(r.get("qty") or 0) * float(r.get("price") or 0) * float(r.get("fx_rate") or 1)

    po_count = len(rows)
    total_qty = sum(float(r["qty"] or 0) for r in rows)
    total_spend = sum(line_spend(r) for r in rows)
    # A historical line item with a £0 buy price (the price was simply
    # never recorded, not that it was actually free) still counts as a
    # real purchase -- po_count/total_qty above deliberately include it --
    # but it must never drag the average PRICE down. Excluding it from
    # avg_price means excluding its qty from the denominator too, not just
    # its (already-zero) contribution to the numerator, since weighting by
    # an inflated total_qty would still understate the true average even
    # with a correct numerator.
    priced_rows = [r for r in rows if float(r.get("price") or 0) > 0]
    priced_qty = sum(float(r["qty"] or 0) for r in priced_rows)
    priced_spend = sum(line_spend(r) for r in priced_rows)
    avg_price = (priced_spend / priced_qty) if priced_qty else 0.0
    first_ordered_at = rows[-1]["created_at"] if rows else ""
    last_ordered_at = rows[0]["created_at"] if rows else ""

    monthly = {}
    for r in rows:
        period = (r["created_at"] or "")[:7]
        if not period:
            continue
        entry = monthly.setdefault(period, {"period": period, "total": 0.0, "po_count": 0})
        entry["total"] += line_spend(r)
        entry["po_count"] += 1
    monthly_list = [monthly[k] for k in sorted(monthly)]

    by_supplier = {}
    for r in rows:
        name = r["supplier_company_name"] or "(no supplier)"
        entry = by_supplier.setdefault(name, {
            "supplier": name, "total_qty": 0.0, "total_spend": 0.0, "times_ordered": 0,
            "last_price": None, "last_ordered_at": "",
        })
        entry["total_qty"] += float(r["qty"] or 0)
        entry["total_spend"] += line_spend(r)
        entry["times_ordered"] += 1
        # rows are newest-first, so the first hit per supplier is already
        # that supplier's most recent ORDER for this product -- but its
        # last VALID price specifically skips forward past any more recent
        # £0/unrecorded row to the most recent one that actually has a
        # real price, so a supplier's scorecard entry never shows "last
        # price: £0.00" just because the latest order for it happened to
        # be an old import with no price on record.
        if not entry["last_ordered_at"]:
            entry["last_ordered_at"] = r["created_at"] or ""
        price = float(r["price"] or 0)
        if entry["last_price"] is None and price > 0:
            entry["last_price"] = price
    supplier_rows = sorted(by_supplier.values(), key=lambda e: e["total_spend"], reverse=True)

    price_history = get_price_history_for_product(conn, product_name)
    # Ignore £0/missing prices here too -- a historical record with no
    # price on file must never look like the "lowest price ever paid".
    priced = [h for h in price_history if h.get("price") is not None and h["price"] > 0]
    lowest_entry = min(priced, key=lambda h: h["price"]) if priced else None

    return {
        "product_name": product_name,
        "po_count": po_count,
        "total_qty": total_qty,
        "total_spend": total_spend,
        "avg_price": avg_price,
        "first_ordered_at": first_ordered_at,
        "last_ordered_at": last_ordered_at,
        "monthly": monthly_list,
        "by_supplier": supplier_rows,
        "price_history": price_history,
        "lowest_price": lowest_entry["price"] if lowest_entry else None,
        "lowest_price_entry": lowest_entry,
    }


def find_recent_price_request(conn, product_name, days=None):
    """The most recent price_requests row for this exact product (any
    supplier, any status) -- used to flag "you already asked about this"
    before firing off what might be a duplicate request. By default looks
    at any point in time (not just a recent window), since a price asked
    for months ago is still worth being flagged rather than silently
    re-asked. Pass days=N to only count one created within the last N
    days instead. Returns None if there isn't a matching one."""
    product_name = (product_name or "").strip()
    if not product_name:
        return None
    if days is not None:
        cutoff = (datetime.now() - timedelta(days=days)).isoformat(timespec="seconds")
        row = conn.execute(
            "SELECT * FROM price_requests WHERE product_name = ? AND created_at >= ? "
            "ORDER BY created_at DESC LIMIT 1",
            (product_name, cutoff),
        ).fetchone()
    else:
        row = conn.execute(
            "SELECT * FROM price_requests WHERE product_name = ? ORDER BY created_at DESC LIMIT 1",
            (product_name,),
        ).fetchone()
    return dict(row) if row else None


# ============================================================
# 10. Zoho Books purchase order import / export
# ============================================================
#
# Zoho Books exports one CSV row per PO line item, with the PO-level fields
# (Reference#, Vendor Name, Purchase Order Status, etc.) repeated on every
# row of that PO. Import groups rows back into one PO per Reference# (falling
# back to Purchase Order Number if Reference# is blank). This never
# overwrites or auto-merges anything: an exact reference match (after
# ignoring case/punctuation) is treated as already-imported and skipped, a
# close-but-not-exact match is only ever surfaced for the user to confirm,
# and a vendor name with no known match is left for the user to map — once
# mapped, that mapping is remembered for future imports.

def _normalize_ref(ref):
    """Normalize a PO reference for comparison. Strips punctuation/case, and
    also strips a leading "PO" -- this app's own po_ref always starts with
    the literal prefix "PO:" (see make_po_ref), but Zoho's Reference# field
    only ever contains the meaningful part after that prefix, so without
    this the two would never be recognised as the same order."""
    norm = re.sub(r"[^A-Z0-9]+", "", (ref or "").upper())
    if norm.startswith("PO") and len(norm) > 2:
        norm = norm[2:]
    return norm


def parse_zoho_export_csv(path):
    """Parse a Zoho Books Purchase Order export CSV into one grouped dict per
    PO: {reference, po_number, vendor_name, status, order_date, purchase_owner,
    currency, items:[...]}. Items with a QuantityOrdered of 0 or less are
    skipped (Zoho sometimes keeps a zero-quantity placeholder line for a
    fully-cancelled item). purchase_owner comes from Zoho's own
    "CF.Purchase Owner" custom field when the export has it -- import_zoho_po
    prefers this over deriving an owner from the PO reference's prefix
    letters. currency comes from Zoho's "Currency Code" column; Zoho's PO
    export has no exchange-rate data, so import_zoho_po can't set a real
    fx_rate for a non-home-currency order -- it falls back to 1:1 like any
    other path with no rate info, correctable later by editing the PO."""
    groups = {}
    order = []
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            ref = (row.get("Reference#") or "").strip() or (row.get("Purchase Order Number") or "").strip()
            if not ref:
                continue
            if ref not in groups:
                groups[ref] = {
                    "reference": ref,
                    "po_number": (row.get("Purchase Order Number") or "").strip(),
                    "vendor_name": (row.get("Vendor Name") or "").strip(),
                    "status": (row.get("Purchase Order Status") or "").strip(),
                    "order_date": (row.get("Purchase Order Date") or "").strip(),
                    "purchase_owner": (row.get("CF.Purchase Owner") or "").strip(),
                    "currency": (row.get("Currency Code") or "").strip(),
                    "items": [],
                }
                order.append(ref)
            item_name = (row.get("Item Name") or "").strip() or (row.get("Item Desc") or "").strip()
            if not item_name:
                continue
            try:
                qty = float(row.get("QuantityOrdered") or 0)
            except ValueError:
                qty = 0
            if qty <= 0:
                continue
            try:
                price = float(row.get("Item Price") or 0)
            except ValueError:
                price = 0
            groups[ref]["items"].append({"product": item_name, "qty": qty, "price": price, "code": ""})
    return [groups[k] for k in order]


def get_zoho_vendor_map(conn):
    return {r["zoho_vendor_name"]: r["supplier_id"] for r in conn.execute("SELECT * FROM zoho_vendor_map")}


def get_zoho_vendor_names_by_supplier(conn):
    """Reverse of get_zoho_vendor_map: {supplier_id: zoho_vendor_name} --
    the exact vendor name Zoho already knows this supplier by, so exports
    back to Zoho can use Zoho's own name rather than risking a second,
    slightly-different vendor record being created on import (e.g. a
    supplier imported as "Currys Business" but entered here as "Currys").
    Where more than one Zoho name has ever mapped to the same supplier
    (rare -- usually a rename that was re-matched), the most recently
    remembered one wins, since that reflects the current match."""
    by_supplier = {}
    # Batch 157: seq instead of SQLite's own implicit rowid -- see its
    # column comment in _ensure_schema_ddl() for why. Same ordering, works
    # identically once this table lives on Postgres instead of SQLite.
    for r in conn.execute("SELECT * FROM zoho_vendor_map ORDER BY seq ASC"):
        by_supplier[r["supplier_id"]] = r["zoho_vendor_name"]
    return by_supplier


def set_zoho_vendor_map(conn, zoho_vendor_name, supplier_id):
    """Remember a Zoho vendor name -> supplier match so future imports of
    that same vendor name don't need to be asked again."""
    # Batch 157: seq is only ever set on a genuine fresh INSERT (one more
    # than the current highest), never touched by the ON CONFLICT ... DO
    # UPDATE path below -- mirroring exactly how a real SQLite rowid never
    # changes when a row is merely updated, only when a new one is created.
    conn.execute(
        "INSERT INTO zoho_vendor_map(zoho_vendor_name, supplier_id, seq) "
        "VALUES (?, ?, (SELECT COALESCE(MAX(seq), 0) + 1 FROM zoho_vendor_map)) "
        "ON CONFLICT(zoho_vendor_name) DO UPDATE SET supplier_id=excluded.supplier_id",
        (zoho_vendor_name, supplier_id),
    )
    conn.commit()


def _levenshtein(a, b):
    """Edit distance (single-character insertions/deletions/substitutions
    needed to turn a into b). Used for PO reference matching instead of a
    generic similarity ratio, because references are structured IDs where
    two DIFFERENT orders can easily look superficially similar (e.g. two
    unrelated orders both using a "-9999" placeholder suffix) -- a ratio
    gets fooled by that, while edit distance only flags references that are
    almost character-for-character the same, i.e. an actual typo/transcription
    difference rather than a coincidentally similar-looking different order."""
    if a == b:
        return 0
    if not a:
        return len(b)
    if not b:
        return len(a)
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i] + [0] * len(b)
        for j, cb in enumerate(b, 1):
            cur[j] = min(
                prev[j] + 1,
                cur[j - 1] + 1,
                prev[j - 1] + (0 if ca == cb else 1),
            )
        prev = cur
    return prev[-1]


def plan_zoho_import(conn, zoho_pos, max_edit_distance=1):
    """Classify parsed Zoho PO groups against what's already in this app.

    Returns {"new": [...], "review": [...], "existing": [...]}:
      - existing: reference normalizes to an exact match with a PO already
        here -> always skipped automatically, never re-imported.
      - review: reference is within max_edit_distance character edits of an
        existing PO's reference (after normalizing away case/punctuation) --
        e.g. one digit added, removed, or mistyped, such as the same order
        logged in Zoho with a small transcription slip. This is deliberately
        an absolute edit-distance check, not a percentage-similarity one:
        two genuinely different references can share a lot of characters
        (a common "-9999" placeholder suffix, a similar date pattern) and
        still not be the same order, so only near-exact matches get flagged.
      - new: no meaningful match -> safe to import as a new historical PO.

    Each entry also carries "vendor_supplier_id": the supplier this vendor
    name is already mapped to (via a remembered zoho_vendor_map entry), or
    None if the user still needs to map that vendor name.
    """
    existing_by_norm = {}
    for r in conn.execute("SELECT po_ref FROM purchase_orders"):
        existing_by_norm.setdefault(_normalize_ref(r["po_ref"]), []).append(r["po_ref"])

    vendor_map = get_zoho_vendor_map(conn)

    new_list, review_list, existing_list = [], [], []
    for raw in zoho_pos:
        po = dict(raw)
        po["vendor_supplier_id"] = vendor_map.get(po["vendor_name"])
        norm = _normalize_ref(po["reference"])
        if norm in existing_by_norm:
            po["matched_ref"] = existing_by_norm[norm][0]
            po["match_ratio"] = 1.0
            existing_list.append(po)
            continue
        best_dist, best_ref = None, None
        for enorm, refs in existing_by_norm.items():
            if abs(len(norm) - len(enorm)) > max_edit_distance:
                continue  # cheap pre-filter -- can't be within distance if lengths differ too much
            dist = _levenshtein(norm, enorm)
            if best_dist is None or dist < best_dist:
                best_dist, best_ref = dist, refs[0]
        if best_dist is not None and best_dist <= max_edit_distance:
            po["matched_ref"] = best_ref
            longest = max(len(norm), len(best_ref and _normalize_ref(best_ref) or ""), 1)
            po["match_ratio"] = 1 - (best_dist / longest)
            review_list.append(po)
        else:
            new_list.append(po)
    return {"new": new_list, "review": review_list, "existing": existing_list}


def import_zoho_po(conn, zoho_po, supplier_id, status=None, commit=True):
    """Create one new historical PO from a resolved Zoho PO group (vendor
    already mapped to supplier_id by the caller). Returns the saved PO, or
    None if it had no importable (positive-quantity) line items."""
    supplier = None
    if supplier_id:
        row = conn.execute("SELECT * FROM suppliers WHERE id=?", (supplier_id,)).fetchone()
        supplier = dict(row) if row else None

    items = [dict(i) for i in zoho_po["items"] if i["qty"] > 0]
    if not items:
        return None

    # The PO reference itself encodes an order date (see parse_date_from_po_ref)
    # and, checked against real Zoho exports, that date is often more accurate
    # than Zoho's own "Purchase Order Date" field, which can lag the real order
    # date by anywhere from 1 to 9 days due to data-entry delay on Zoho's side.
    # So prefer the ref date, fall back to Zoho's order_date, then to "now".
    ref_date = parse_date_from_po_ref(zoho_po.get("reference"))
    order_date = (zoho_po.get("order_date") or "").strip()
    if ref_date:
        created_at = f"{ref_date}T00:00:00"
    elif order_date:
        created_at = f"{order_date}T00:00:00"
    else:
        created_at = now_iso()

    use_status = status or get_setting(conn, "zoho_import_default_status", "Sent")
    if use_status not in STATUS_CHOICES:
        use_status = "Sent"

    po = {
        "po_ref": zoho_po["reference"],
        "status": use_status,
        "supplier_id": supplier_id or 0,
        "supplier_company_name": (supplier or {}).get("company_name") or zoho_po.get("vendor_name", ""),
        "supplier_contact_name": (supplier or {}).get("contact_name", ""),
        "supplier_email": (supplier or {}).get("email", ""),
        "supplier_cc_emails": (supplier or {}).get("cc_emails", ""),
        "supplier_phone": (supplier or {}).get("phone", ""),
        "supplier_address": (supplier or {}).get("address", ""),
        "delivery_label": "",
        "delivery_name": "",
        "delivery_address": "",
        "business_name": get_setting(conn, "business_name", ""),
        "company_number": get_setting(conn, "company_number", ""),
        "currency": zoho_po.get("currency") or get_setting(conn, "currency", "GBP"),
        "invoice_name": get_setting(conn, "invoice_name", ""),
        "invoice_address": get_setting(conn, "invoice_address", ""),
        "your_name": get_setting(conn, "your_name", ""),
        "bcc_emails": "",
        "notes": ("Imported from Zoho Books" + (f" (PO {zoho_po['po_number']})" if zoho_po.get("po_number") else "")).strip(),
        # Not a real PO_FIELDS column -- save_po reads this straight off the
        # dict to prefer Zoho's own "purchase owner" field over deriving one
        # from the reference prefix, without it becoming a stray DB column.
        "_zoho_purchase_owner": zoho_po.get("purchase_owner", ""),
        # Zoho's own "Purchase Order Status" (Draft/Issued/Billed/Partially
        # Billed) -- billing status, not payment status (Zoho's PO export
        # has no payment data), kept as extra context for credit limit
        # tracking rather than a source of truth for what's been paid.
        "zoho_po_status": zoho_po.get("status", ""),
    }
    # Learn this prefix -> owner pairing (if it's new) before saving, so
    # Settings > Purchase Owners is populated straight from real Zoho data
    # instead of only ever being typed in by hand.
    learn_purchase_owner_prefix_from_zoho(conn, zoho_po["reference"], zoho_po.get("purchase_owner", ""), commit=commit)
    return save_po(conn, po, items, status=use_status, event="imported from Zoho Books",
                    record_product_memory=True, created_at=created_at, commit=commit)


def backfill_order_dates_from_ref(conn, commit=True):
    """One-time repair for POs imported before ref-based date parsing
    existed, where created_at was set purely from Zoho's own "Purchase
    Order Date" field (which testing showed can lag the real order date by
    1-9 days). For every non-deleted PO whose reference encodes a
    parseable date (see parse_date_from_po_ref) and whose currently stored
    created_at date differs from it, corrects the date part of created_at
    to the ref-derived date (keeping whatever time-of-day is already
    there). Returns the number of POs updated."""
    rows = conn.execute("SELECT id, po_ref, created_at FROM purchase_orders WHERE deleted=0").fetchall()
    updated = 0
    for r in rows:
        ref_date = parse_date_from_po_ref(r["po_ref"])
        if not ref_date:
            continue
        created_at = r["created_at"] or ""
        current_date = created_at[:10]
        if current_date == ref_date:
            continue
        time_part = created_at[10:] or "T00:00:00"
        new_created_at = f"{ref_date}{time_part}"
        conn.execute("UPDATE purchase_orders SET created_at=? WHERE id=?", (new_created_at, r["id"]))
        updated += 1
    if commit and updated:
        conn.commit()
    return updated


def export_pos_to_zoho_csv(conn, path, date_from=None, date_to=None, status=None, deleted=False):
    """Write a Zoho Books-importable Purchase Order CSV (one row per line
    item) for POs in the given date range (YYYY-MM-DD strings, inclusive).

    Rebuilt (change request section 12) to reuse the exact same
    comprehensive, ~79-column data-building logic as the automatic daily
    Zoho export workbook (build_zoho_export_rows / ZOHO_COLUMN_SPECS,
    defined below this function) instead of its own separate, much
    thinner 16-column version, which "misses much of the required
    information" per Yitzi's own description -- every column Zoho's own
    PO import format expects (cost centre, VAT treatment, ship-to address,
    tax amounts, the "CF.Purchase Owner" custom field, etc.) is now
    populated here exactly as it already is on the daily export, using
    whatever's actually on file (the Settings > Zoho Export defaults, the
    PO's own resolved purchase_owner, the Zoho-remembered vendor name) and
    leaving a field blank only when there's genuinely nothing to put there
    -- nothing invented. The corrected bare-then-displayed PO number
    format (batch 76) is used automatically here, same as everywhere else
    that reads po_ref, since po_ref itself has carried no "PO:" prefix
    since that fix.

    Same filters as before this rebuild: an inclusive date range, an exact
    PO status, or deleted POs instead of live ones. "Purchase Order
    Status" still comes from the zoho_export_default_status setting
    (Draft by default) via build_zoho_export_rows, same as the daily
    export. Returns the number of POs written (not rows -- a PO with
    several line items still counts once, matching this function's
    original return convention)."""
    sql = "SELECT * FROM purchase_orders WHERE deleted=?"
    params = [1 if deleted else 0]
    if date_from:
        sql += " AND substr(created_at,1,10) >= ?"
        params.append(date_from)
    if date_to:
        sql += " AND substr(created_at,1,10) <= ?"
        params.append(date_to)
    if status:
        sql += " AND status=?"
        params.append(status)
    sql += " ORDER BY created_at"
    pos = [dict(r) for r in conn.execute(sql, params)]

    rows = build_zoho_export_rows(conn, pos)
    headers = [spec[0] for spec in ZOHO_COLUMN_SPECS]
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        for row in rows:
            writer.writerow([row.get(h, "") for h in headers])
    return len(pos)


# ---- Daily Zoho export (full-column Excel workbook + reminder email) ----
#
# Zoho Books' own Purchase Order export/import format has far more columns
# than the simple CSV above -- most of them accounting/business settings
# (cost centre, tax treatment, ship-to address) that this app has no way to
# know per order, so every one of them comes from an editable Settings >
# Zoho Export default (see DEFAULT_SETTINGS's zoho_col_* keys) rather than
# being guessed. Each spec below is (header, kind, extra):
#   "po"       -- extra is a callable(po) -> value
#   "item"     -- extra is a callable(po, item) -> value
#   "setting"  -- extra is the settings key to read
#   "fixed"    -- extra is a literal value, the same on every row
#   "blank"    -- always ""
# "computed_tax", "computed_item_total", and "zoho_grand_total" are also
# used below, all special-cased directly in build_zoho_export_rows.
ZOHO_COLUMN_SPECS = [
    ("Purchase Order ID", "blank", None),
    ("Purchase Order Date", "po", lambda po: (po.get("created_at") or "")[:10]),
    ("Delivery Date", "blank", None),
    # Yitzi: "in collem D of Zoho export it asks for Purchase Order Number
    # it should be the samer as Reference# collum E so same info in both" --
    # both columns just carry the PO's own reference now.
    ("Purchase Order Number", "po", lambda po: po.get("po_ref", "")),
    ("Reference#", "po", lambda po: po.get("po_ref", "")),
    ("Purchase Order Status", "setting", "zoho_export_default_status"),
    ("Vendor Name", "po", lambda po: po.get("supplier_company_name", "")),
    ("Vendor Number", "blank", None),
    ("Company Registration Number", "blank", None),
    ("Purchase Order VAT Treatment", "setting", "zoho_col_vat_treatment"),
    ("Is Inclusive Tax", "setting", "zoho_col_is_inclusive_tax"),
    ("Currency Code", "po", lambda po: po.get("currency", "GBP")),
    ("Exchange Rate", "setting", "zoho_col_exchange_rate"),
    ("Template Name", "setting", "zoho_col_template_name"),
    ("Reference No", "po", lambda po: po.get("po_ref", "")),
    ("Delivery Instructions", "setting", "zoho_col_delivery_instructions"),
    ("Terms & Conditions", "setting", "zoho_col_terms_conditions"),
    ("Shipment preference", "setting", "zoho_col_shipment_preference"),
    ("Expected Arrival Date", "blank", None),
    ("Account", "setting", "zoho_col_account"),
    ("Account Code", "setting", "zoho_col_account_code"),
    ("Item Price", "item", lambda po, item: f"{float(item.get('price', 0)):.2f}"),
    ("Item Name", "item", lambda po, item: display_product(item)),
    ("Product ID", "blank", None),
    ("Item Desc", "blank", None),
    # Header is plain "Quantity", NOT "QuantityOrdered" -- see
    # ZOHO_QUANTITY_HEADER in build_zoho_export_rows's docstring for why:
    # Zoho's own real export calls this column "QuantityOrdered", but its
    # IMPORT auto-mapper doesn't recognise that as its own "Quantity" field,
    # so it was silently going unmapped and defaulting every line's
    # quantity to 1 -- the real root cause of the whole "Total" saga below,
    # confirmed by Yitzi manually mapping it during a real import.
    ("Quantity", "item", lambda po, item: item.get("qty", 0)),
    ("QuantityCancelled", "fixed", "0.00"),
    ("QuantityReceived", "fixed", "0.00"),
    ("QuantityBilled", "fixed", "0.00"),
    ("Usage unit", "setting", "zoho_col_usage_unit"),
    ("Discount Type", "setting", "zoho_col_discount_type"),
    ("Is Discount Before Tax", "setting", "zoho_col_is_discount_before_tax"),
    ("Discount", "fixed", "0.00"),
    ("Discount Amount", "fixed", "0.00"),
    ("Tax ID", "setting", "zoho_col_tax_id"),
    ("Item Tax", "setting", "zoho_col_item_tax"),
    ("Item Tax %", "setting", "zoho_col_item_tax_pct"),
    ("Item Tax Amount", "computed_tax", None),
    ("Item Tax Type", "setting", "zoho_col_item_tax_type"),
    ("Item Exemption Code", "setting", "zoho_col_item_exemption_code"),
    ("Item Type", "setting", "zoho_col_item_type"),
    ("Acquisition/Rev. Charge VAT Name", "setting", "zoho_col_acq_vat_name"),
    ("Acquisition/Rev. Charge VAT Percentage", "setting", "zoho_col_acq_vat_pct"),
    ("Item Total", "computed_item_total", None),
    # NOT po["total"] (that's the NET total; this needs to be the VAT-
    # inclusive, QUANTITY-inclusive PO grand total). See
    # ZOHO_QUANTITY_HEADER in the long comment on build_zoho_export_rows --
    # the real root cause of the whole "Total" saga (Batches 88-90) was the
    # "QuantityOrdered" header above not auto-mapping to Zoho's own
    # "Quantity" import field, silently defaulting every line's quantity to
    # 1. Now that the header is fixed and Yitzi confirmed a real import
    # with Quantity correctly mapped, Zoho's own "system generated total"
    # check turned out to be quantity-inclusive all along -- there was
    # never a quantity-blind quirk in Zoho's validator itself.
    ("Total", "zoho_grand_total", None),
    ("Adjustment", "fixed", "0.00"),
    ("Adjustment Description", "setting", "zoho_col_adjustment_description"),
    ("Entity Discount Percent", "fixed", "0.00"),
    ("Entity Discount Amount", "fixed", "0.000"),
    ("Discount Account", "setting", "zoho_col_discount_account"),
    ("Discount Account Code", "setting", "zoho_col_discount_account_code"),
    ("1. Cost Centre", "setting", "zoho_col_cost_centre"),
    ("2. Cost Allocation", "setting", "zoho_col_cost_allocation"),
    ("6. Group Cost Allocation (Finance Team Only)", "setting", "zoho_col_group_cost_allocation"),
    ("5. Profit Centre (Finance Team Only)", "setting", "zoho_col_profit_centre"),
    ("3. Location", "setting", "zoho_col_location"),
    ("8. Marketing Tag", "setting", "zoho_col_marketing_tag"),
    ("7. Person (Finance Team Only)", "setting", "zoho_col_person"),
    ("4. Budgets", "setting", "zoho_col_budgets"),
    ("Project ID", "setting", "zoho_col_project_id"),
    ("Project Name", "setting", "zoho_col_project_name"),
    ("Payment Terms", "setting", "zoho_col_payment_terms"),
    ("Payment Terms Label", "setting", "zoho_col_payment_terms_label"),
    ("Attention", "setting", "zoho_col_attention"),
    ("Address", "setting", "zoho_col_address"),
    ("City", "setting", "zoho_col_city"),
    ("State", "setting", "zoho_col_state"),
    ("Country", "setting", "zoho_col_country"),
    ("Code", "setting", "zoho_col_postcode"),
    ("Phone", "setting", "zoho_col_phone"),
    ("Deliver To Customer", "setting", "zoho_col_deliver_to_customer"),
    ("Recipient Address", "blank", None),
    ("Recipient City", "blank", None),
    ("Recipient State", "blank", None),
    ("Recipient Country", "blank", None),
    ("Recipient Postal Code", "blank", None),
    ("Recipient Phone", "blank", None),
    ("Submitted By", "blank", None),
    ("Approved By", "blank", None),
    ("Submitted Date", "blank", None),
    ("Approved Date", "blank", None),
    # Per-PO, not a flat setting -- each PO already carries its own resolved
    # purchase_owner (Zoho's own field on import, or derived from the PO
    # reference's prefix via Settings > Purchase Owners -- see
    # resolve_purchase_owner). zoho_col_purchase_owner is kept only as the
    # fallback for a PO whose owner couldn't be resolved at all, just below.
    #
    # Header is plain "Purchase Owner", not "CF.Purchase Owner" -- Yitzi:
    # "All zoho exports The collum should not say CF.Purchase Owner just
    # Purchase Owner". No other custom-field column in this export carries
    # a "CF." prefix either, so this only ever matched this one column.
    # NOTE: this is the OUTGOING export column only. The unrelated
    # parse_zoho_export_csv (reading a real Zoho Books export back in) still
    # looks for "CF.Purchase Owner" there, since that's Zoho's own genuine
    # custom-field naming convention on ITS output, not this app's choice.
    ("Purchase Owner", "po", lambda po: po.get("purchase_owner", "")),
]


def _zoho_item_tax_pct(settings, item):
    """The Item Tax %/name a single po_items row exports with -- shared by
    build_zoho_export_rows for both the real "Item Tax %" column and the
    "Total" grand total (see ZOHO_TOTAL_IS_GROSS_PER_PO), so the two can
    never drift apart. A line ticked "Margin VAT" (po_items.margin_vat)
    uses the separate zoho_margin_vat_item_tax/_pct default instead of the
    normal Item Tax columns, on that line only -- everything else about the
    PO (VAT Treatment, currency, etc.) is unaffected."""
    is_margin_vat = bool(item.get("margin_vat"))
    if is_margin_vat:
        item_tax_name = settings.get("zoho_margin_vat_item_tax", "No VAT")
        key = "zoho_margin_vat_item_tax_pct"
    else:
        item_tax_name = settings.get("zoho_col_item_tax", "")
        key = "zoho_col_item_tax_pct"
    try:
        tax_pct = float(settings.get(key, "") or 0)
    except ValueError:
        tax_pct = 0.0
    return is_margin_vat, item_tax_name, tax_pct


def build_zoho_export_rows(conn, pos):
    """One row per PO line item, in ZOHO_COLUMN_SPECS order, ready to write
    straight to an Excel workbook (see export_pos_to_zoho_excel). Vendor
    Name uses Zoho's own remembered name for that supplier when one's on
    record (see get_zoho_vendor_names_by_supplier), so the export matches
    the vendor Zoho already has rather than risking a near-duplicate vendor
    being created on the Zoho side.

    ZOHO_QUANTITY_HEADER / ZOHO_TOTAL_IS_GROSS_PER_PO: the "Total" column is
    the real VAT-and-quantity-inclusive PO grand total --
    sum(Item Total + Item Tax Amount) per PO, repeated on every line -- and
    the "Quantity" column header (see ZOHO_COLUMN_SPECS) is plain
    "Quantity", not "QuantityOrdered". This is the resolution of a six-pass
    investigation, and it's worth reading in full: the real root cause
    turned out to be neither the Total formula nor a Zoho validator quirk
    at all, both of which were chased hard across five earlier passes.

    Pass 1: reverse-engineered a formula from Zoho's "given total"/"system
    generated total" mismatch error on 8 real skipped POs, matched those 8
    numbers to the penny, and shipped a quantity-blind Total. Coincidentally
    "worked" for the wrong reason -- see Pass 6.

    Pass 2: Yitzi's screenshots showed real bulk imports recording
    Quantity=1 on every line regardless of the real quantity ordered. Item
    Price, Item Total, Total, and every "0.00" fixed placeholder column
    were being written into the .xlsx as literal TEXT strings, not real
    numbers, while QuantityOrdered was already numeric -- fixed via
    _zoho_cell_value/export_pos_to_zoho_excel. This fix is harmless and
    still in effect, but Pass 6 proves it was never the actual cause of the
    Quantity=1 symptom.

    Pass 3: Yitzi sent a real Purchase Order export straight out of his own
    Zoho account -- e.g. reference MR230126-9999, four lines with Item
    Totals 225/450/300/864 (sum 1839) and Item Tax Amounts 45/90/60/172.80
    (sum 367.80): every one of its four rows carries Total = 2206.80,
    exactly 1839 + 367.80. Total was changed to sum(Item Total + Item Tax
    Amount) per PO -- correct then, and (per Pass 6) still correct now.

    Pass 4: that Pass-3 formula got rejected by Zoho's import validator on 6
    more real POs. given/system ratios (10.0, 100.0, 75.0, 5.6487, 2.0,
    34.6) matched each PO's line quantity, which looked like conclusive
    proof of a quantity-blind validator quirk. "Total" was left blank
    instead, on the theory that a blank cell gives the validator nothing to
    disagree with.

    Pass 5: blank didn't work either -- a blank cell reads as 0, and the
    same check ran and failed again on the same POs. "Total" was reinstated
    to Pass 1's quantity-blind formula, since the check was now proven
    mandatory and unskippable.

    Pass 6 (this one, the real fix): asked Yitzi to check Zoho's import
    column-mapping screen directly. He confirmed "QuantityOrdered" was NOT
    auto-matching to Zoho's own "Quantity" import field -- Zoho's exporter
    calls this column "QuantityOrdered", but its importer's field is
    labelled "Quantity", and the auto-mapper doesn't bridge the two. Left
    unmapped, every line silently defaulted to Quantity=1 -- on the ACTUAL
    STORED PO RECORD, not just in some separate validator. That single fact
    explains everything: Item Price/Amount always looked "right" because
    Amount = Rate * 1 also looks plausible at a glance; the real "system
    generated total" was ALWAYS computing Quantity * Rate * (1 + tax%)
    correctly, it just always got quantity=1 fed into it, making it look
    quantity-blind by coincidence in every single test across Passes 1, 4,
    and 5. Yitzi then manually mapped Quantity during a real import and
    re-ran it: the given/system values from Pass 5's quantity-blind Total
    came back EXACTLY REVERSED from every previous test -- e.g. reference
    YJ250826-1143, given 518.40 (this app's Pass-5 quantity-blind value),
    system 51,840.00, which is exactly (162+93+177)*100*1.2 -- the real
    quantity-inclusive total, confirmed to the penny across all 7 real POs
    in that test including one never seen before (YJ240826-1717, given
    771.60, system 3,086.40 = 643*4*1.2). There never was a quantity-blind
    quirk in Zoho's validator; there was only ever an unmapped column
    silently zeroing out Quantity everywhere at once. Fixed by renaming the
    export header to plain "Quantity" (so it auto-maps without requiring
    manual remapping on every import) and reverting "Total" to Pass 3's
    real formula, which is now CONFIRMED against a live import, not
    inferred. Pass 2's numeric-cell-type fix stays in place regardless (it's
    real and harmless, just not the cause of any of this), and Pass 4/5's
    blank/quantity-blind Total logic is fully removed, not left dead."""
    settings = get_all_settings(conn)
    zoho_names_by_supplier = get_zoho_vendor_names_by_supplier(conn)
    rows = []
    if not pos:
        return rows
    placeholders = ",".join("?" * len(pos))
    items_by_po = {}
    for r in conn.execute(
        f"SELECT * FROM po_items WHERE po_id IN ({placeholders}) ORDER BY po_id, position",
        [po["id"] for po in pos],
    ):
        items_by_po.setdefault(r["po_id"], []).append(dict(r))

    # First pass: the real VAT-and-quantity-inclusive grand total per PO
    # (see ZOHO_TOTAL_IS_GROSS_PER_PO above) -- Pass 3's original formula,
    # confirmed correct again by Pass 6. Uses the exact same per-line tax
    # logic the second pass uses for the real "Item Tax %" column, so the
    # two can never disagree with each other.
    check_total_by_po = {}
    for po in pos:
        total = 0.0
        for item in items_by_po.get(po["id"], []):
            _, _, tax_pct = _zoho_item_tax_pct(settings, item)
            line_total = float(item.get("qty", 0)) * float(item.get("price", 0))
            total += line_total * (1 + tax_pct / 100)
        check_total_by_po[po["id"]] = total

    for po in pos:
        for item in items_by_po.get(po["id"], []):
            item_total = float(item.get("qty", 0)) * float(item.get("price", 0))
            is_margin_vat, item_tax_name, tax_pct = _zoho_item_tax_pct(settings, item)
            tax_amount = round(item_total * tax_pct / 100, 2) if tax_pct else ""
            row = {}
            for header, kind, extra in ZOHO_COLUMN_SPECS:
                if header == "Item Tax" and is_margin_vat:
                    row[header] = item_tax_name
                elif header == "Item Tax %" and is_margin_vat:
                    row[header] = f"{tax_pct:.2f}"
                elif kind == "po":
                    row[header] = extra(po)
                elif kind == "item":
                    row[header] = extra(po, item)
                elif kind == "setting":
                    row[header] = settings.get(extra, "")
                elif kind == "fixed":
                    row[header] = extra
                elif kind == "computed_tax":
                    row[header] = tax_amount
                elif kind == "computed_item_total":
                    row[header] = f"{item_total:.2f}"
                elif kind == "zoho_grand_total":
                    row[header] = f"{check_total_by_po.get(po['id'], 0.0):.2f}"
                else:
                    row[header] = ""
            if not row.get("Purchase Owner"):
                row["Purchase Owner"] = settings.get("zoho_col_purchase_owner", "")
            zoho_name = zoho_names_by_supplier.get(po.get("supplier_id"))
            if zoho_name:
                row["Vendor Name"] = zoho_name
            rows.append(row)
    return rows


def default_zoho_export_path(now=None):
    """Where the daily Zoho export workbook is saved -- kept (not deleted
    after sending) as a natural local record of what's been sent each day,
    under the app's own folder rather than a throwaway temp file."""
    now = now or datetime.now()
    return ZOHO_EXPORT_DIR / f"zoho_export_{now.strftime('%Y%m%d_%H%M%S')}.xlsx"


_ZOHO_NUMERIC_CELL_RE = re.compile(r"^-?\d+(\.\d+)?$")


def _zoho_cell_value(value):
    """A plain numeric-looking string (e.g. the "3.75"/"281.25"/"0.00"
    build_zoho_export_rows writes for Item Price/Item Total/Total/every
    "fixed": "0.00" quantity-and-discount placeholder) becomes a REAL
    Excel number here instead of text -- everything else (dates, names,
    statuses, "true"/"false", already-blank cells) passes through
    untouched. Yitzi sent screenshots of real imports where Quantity was
    silently dropping to 1 on every line no matter what QuantityOrdered
    said (a real 75-unit line importing as 1); QuantityOrdered was already
    a genuine numeric Excel cell (see build_zoho_export_rows), but Item
    Price/Item Total/Total and every "0.00" placeholder were being written
    as literal text strings (openpyxl data_type "s", not "n") -- exactly
    the column Zoho's Rate x Quantity math depends on. A text "Rate" cell
    is a very plausible reason Zoho's importer would fall back to just
    displaying the raw text rather than actually multiplying it by
    Quantity to compute Amount."""
    if isinstance(value, str) and _ZOHO_NUMERIC_CELL_RE.match(value):
        return float(value)
    return value


def export_pos_to_zoho_excel(conn, path, pos=None):
    """Writes the daily Zoho export workbook (one row per line item, every
    column Zoho's own Purchase Order import expects) for the given POs, or
    every not-yet-exported Sent PO if pos isn't given. Returns the list of
    PO refs actually written (empty list if there's nothing outstanding)."""
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter

    if pos is None:
        pos = report_pos_needing_zoho_export(conn)
    if not pos:
        return []

    rows = build_zoho_export_rows(conn, pos)
    headers = [spec[0] for spec in ZOHO_COLUMN_SPECS]

    wb = Workbook()
    ws = wb.active
    ws.title = "Purchase Orders"
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in rows:
        ws.append([_zoho_cell_value(row.get(h, "")) for h in headers])
    for i, header in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(i)].width = max(10, min(28, len(str(header)) + 4))
    wb.save(path)
    return [po["po_ref"] for po in pos]


def build_zoho_export_reminder_email(conn, count, now=None):
    """Subject/plain text for the daily Zoho export email. Placeholders:
    {count} (the bare number), {count_label} ("1 purchase order" / "8
    purchase orders", correctly singular/plural -- what the default body
    template actually uses), {date}. This one attaches the Excel workbook
    itself as a real, visible attachment (see open_zoho_export_email_windows)
    rather than a matching PDF -- confirmed scope: PDF attachments are only
    for the "content" emails, and this one already has its own attachment to
    carry the substance. It also never goes through Outlook's clipboard-paste
    path (no "Copy" button anywhere in the Zoho export flow), so unlike
    build_html/build_price_request_summary it's free to use the real cid:
    logo header (email_brand_header_html) like every other redesigned email
    -- Yitzi: "in the zoho export email is also missing logo" -- rather than
    the text-only fallback those two clipboard-exposed builders are stuck
    with."""
    now = now or datetime.now()
    fields = {
        "count": str(count), "count_label": pluralize(count, "purchase order"),
        "date": now.strftime("%d/%m/%Y"),
    }
    subject_tmpl = get_setting(conn, "zoho_export_email_subject", DEFAULT_SETTINGS["zoho_export_email_subject"])
    body_tmpl = get_setting(conn, "zoho_export_email_body", DEFAULT_SETTINGS["zoho_export_email_body"])
    try:
        subject = subject_tmpl.format(**fields)
    except (KeyError, IndexError):
        subject = subject_tmpl
    try:
        plain = body_tmpl.format(**fields)
    except (KeyError, IndexError):
        plain = body_tmpl
    body_html = "".join(
        f'<div style="font-size:14px;color:{RCG_INK};margin:0 0 10px 0;">{html_escape(line)}</div>'
        if line.strip() else '<div style="margin:0 0 10px 0;">&nbsp;</div>'
        for line in plain.splitlines()
    )
    header_html = email_brand_header_html()
    footer_html = email_brand_footer_html()
    html = f"""
    <html>
    <body>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#f3f1f9;">
    <tr><td align="center" style="padding:24px 12px;">
    <table role="presentation" width="600" cellpadding="0" cellspacing="0" border="0" style="width:600px;max-width:600px;background:#ffffff;font-family:Arial,'Segoe UI',Helvetica,sans-serif;">
    <tr><td>{header_html}</td></tr>
    <tr><td style="padding:16px 36px 0 36px;">{body_html}</td></tr>
    <tr><td>{footer_html}</td></tr>
    </table>
    </td></tr>
    </table>
    </body>
    </html>
    """
    return {"subject": subject, "plain": plain, "html": html}


# ---- Zoho Books products/items (catalog import/export) ----
#
# Same "ask once, remember it" vendor-matching approach as PO import, reusing
# the same zoho_vendor_map table -- so mapping a vendor while importing items
# also covers that vendor for PO import (and vice versa).

def parse_zoho_items_csv(path):
    """Parse a Zoho Books Item export CSV into [{name, price, vendor_name,
    status}, ...]. Purchase Rate arrives as e.g. "GBP 96.85"; the currency
    prefix is stripped and the number is used as the product's price."""
    items = []
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            name = (row.get("Item Name") or "").strip()
            if not name:
                continue
            rate_raw = (row.get("Purchase Rate") or "").strip().replace(",", "")
            m = re.search(r"[\d.]+", rate_raw)
            price = float(m.group(0)) if m else 0.0
            items.append({
                "name": name,
                "price": price,
                "vendor_name": (row.get("Vendor") or "").strip(),
                "status": (row.get("Status") or "").strip(),
            })
    return items


def plan_zoho_item_import(conn, zoho_items):
    """Classify parsed Zoho items against the existing product catalog.
    Returns {"new": [...], "changed": [...], "unchanged": [...]}. An item
    whose vendor name isn't yet mapped to a supplier has
    vendor_supplier_id=None and needs the user to resolve it (map to an
    existing supplier, or leave as "any supplier") before it can be imported;
    a blank Zoho vendor is treated as "any supplier" (supplier_id 0)
    straight away. 'changed' entries match an already-catalogued product for
    that supplier by name but at a different price -- importing them
    refreshes the price, nothing else changes. 'unchanged' entries are a
    100% match (same supplier, same name, same price already on file) --
    these are left alone entirely rather than re-imported as if they were a
    fresh row, so a re-run of the same export doesn't touch anything that
    hasn't actually changed."""
    vendor_map = get_zoho_vendor_map(conn)
    # Keyed by (supplier_id, normalized_name) -> a LIST of candidates, not a
    # single row: post-Batch-151 tidy-up fix -- a supplier can genuinely
    # stock both "Galaxy S25" and "Galaxy S25+" (same normalized text once
    # punctuation is stripped, but _product_names_match's own variant-marker
    # check treats them as different products, same rule
    # find_close_matching_products/upsert_product_manual/etc. all follow).
    # Collapsing straight to one row per key the way this used to would
    # silently drop one of the two, and Zoho's price for one variant would
    # get matched against the other's catalog row.
    by_key = {}
    for p in conn.execute("SELECT * FROM products"):
        by_key.setdefault((p["supplier_id"], _normalize_product_name(p["name"])), []).append(dict(p))

    new_list, changed_list, unchanged_list = [], [], []
    for raw in zoho_items:
        item = dict(raw)
        if not item["vendor_name"]:
            item["vendor_supplier_id"] = 0
        else:
            item["vendor_supplier_id"] = vendor_map.get(item["vendor_name"])
        if item["vendor_supplier_id"] is not None:
            key = (item["vendor_supplier_id"], _normalize_product_name(item["name"]))
            candidates = by_key.get(key, [])
            match = next((c for c in candidates if _product_names_match(c["name"], item["name"])), None)
            if match:
                item["matched_product_id"] = match["id"]
                if round(float(match["last_price"] or 0), 2) == round(float(item["price"] or 0), 2):
                    unchanged_list.append(item)
                else:
                    changed_list.append(item)
                continue
        new_list.append(item)
    return {"new": new_list, "changed": changed_list, "unchanged": unchanged_list}


def import_zoho_item(conn, zoho_item, supplier_id, commit=True):
    """Create a new catalog product, or refresh the price of a matching
    existing one, from a resolved Zoho item (vendor already mapped to
    supplier_id by the caller; 0 means 'any supplier'). Zoho's item export
    has no product-code column, so an existing product's own code (if it's
    ever had one set by hand) is preserved rather than blanked out by every
    re-import."""
    supplier_id = supplier_id or 0
    existing = conn.execute(
        "SELECT id, code FROM products WHERE supplier_id=? AND name=?",
        (supplier_id, zoho_item["name"]),
    ).fetchone()
    existing_code = existing["code"] if existing else ""
    upsert_product_manual(
        conn, existing["id"] if existing else None, supplier_id, existing_code or "", zoho_item["name"],
        zoho_item["price"], commit=commit,
    )


ZOHO_ITEM_EXPORT_HEADERS = ["Item Name", "Purchase Rate", "Vendor", "Product Type"]


def export_products_to_zoho_csv(conn, path, supplier_id=None):
    """Write the product catalog out as a Zoho Books-importable Item CSV."""
    sql = ("SELECT p.*, s.company_name AS vendor_name FROM products p "
           "LEFT JOIN suppliers s ON p.supplier_id = s.id")
    params = []
    if supplier_id is not None:
        sql += " WHERE p.supplier_id=?"
        params.append(supplier_id)
    sql += " ORDER BY p.name COLLATE NOCASE"
    rows = [dict(r) for r in conn.execute(sql, params)]
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(ZOHO_ITEM_EXPORT_HEADERS)
        for r in rows:
            writer.writerow([r["name"], f"GBP {float(r['last_price']):.2f}", r.get("vendor_name") or "", "goods"])
    return len(rows)


def list_zoho_import_history(conn, limit=200):
    """Every PO ever pulled in via a Zoho import, most recently imported
    first -- built from the po_events log (each import writes an
    'imported from Zoho Books' event), so it needs no extra schema and
    shows both the order's own date and when it was actually imported."""
    rows = conn.execute(
        "SELECT po.po_ref AS po_ref, po.supplier_company_name AS supplier, po.total AS total, "
        "po.status AS status, po.created_at AS order_date, MAX(ev.at) AS imported_at "
        "FROM purchase_orders po JOIN po_events ev ON ev.po_id = po.id "
        "WHERE ev.event = 'imported from Zoho Books' "
        "GROUP BY po.id ORDER BY imported_at DESC LIMIT ?",
        (limit,),
    ).fetchall()
    return [dict(r) for r in rows]


# ---- Zoho vendor (supplier) import ----
#
# Zoho Books/Inventory exports vendor contacts as their own CSV, separate
# from the Purchase Order and Item exports above -- and the exact column
# headers can vary a bit by plan/region, unlike the PO/Item exports (whose
# headers were confirmed against a real file from this account). So each
# field we care about is matched against a list of likely header spellings,
# case-insensitively, rather than one exact string that might be wrong.
_ZOHO_VENDOR_FIELD_CANDIDATES = {
    "company_name": ["Company Name", "Vendor Name", "Display Name"],
    "contact_name": ["First Name"],
    "last_name": ["Last Name"],
    "email": ["EmailID", "Email", "Email Address"],
    "phone": ["Phone", "Work Phone", "MobilePhone", "Mobile Phone"],
    "address_street": ["Billing Street", "Billing Address", "Street", "Billing Street2"],
    "address_city": ["Billing City", "City"],
    "address_state": ["Billing State", "State"],
    "address_country": ["Billing Country", "Country"],
    "address_zip": ["Billing Code", "Billing Zip Code", "Zip Code", "Postal Code"],
}


def _resolve_zoho_vendor_headers(fieldnames):
    """Match each target field above to the best-fitting column actually in
    the file, exact spelling match first, then substring -- so a real
    export's header text doesn't have to be guessed exactly right up
    front. Each source column is only ever used for one field."""
    available = {(fn or "").strip().lower(): fn for fn in fieldnames if fn}
    resolved = {}
    used = set()
    for field, candidates in _ZOHO_VENDOR_FIELD_CANDIDATES.items():
        found = None
        for cand in candidates:
            key = cand.strip().lower()
            if key in available and available[key] not in used:
                found = available[key]
                break
        if not found:
            for cand in candidates:
                key = cand.strip().lower()
                for lower_fn, fn in available.items():
                    if fn in used:
                        continue
                    if key in lower_fn or lower_fn in key:
                        found = fn
                        break
                if found:
                    break
        if found:
            resolved[field] = found
            used.add(found)
    return resolved


def parse_zoho_vendor_csv(path):
    """Parse a Zoho Books/Inventory Vendor export CSV into
    [{company_name, contact_name, email, phone, address}, ...].
    contact_name combines First Name + Last Name (if both are present) so
    the rest of the app -- which stores one "contact name" field and
    greets by first name only in emails -- can use it as-is. Raises
    ValueError if the file doesn't look like a Zoho vendor export at all
    (no column that resolves to a company/vendor name)."""
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        resolved = _resolve_zoho_vendor_headers(reader.fieldnames or [])
        if not resolved.get("company_name"):
            raise ValueError(
                "Couldn't find a company/vendor name column in that file -- is this a Zoho Vendor export?"
            )

        def get(row, field):
            col = resolved.get(field)
            return (row.get(col) or "").strip() if col else ""

        vendors = []
        for row in reader:
            company_name = get(row, "company_name")
            if not company_name:
                continue
            contact_name = " ".join(p for p in (get(row, "contact_name"), get(row, "last_name")) if p)
            address = ", ".join(
                p for p in (
                    get(row, "address_street"), get(row, "address_city"),
                    get(row, "address_state"), get(row, "address_country"), get(row, "address_zip"),
                ) if p
            )
            vendors.append({
                "company_name": company_name,
                "contact_name": contact_name,
                "email": get(row, "email"),
                "phone": get(row, "phone"),
                "address": address,
            })
    return vendors


def plan_zoho_supplier_import(conn, zoho_vendors, min_ratio=0.6):
    """Classify parsed Zoho vendor rows against existing suppliers.
    Returns {"new": [...], "changed": [...], "unchanged": [...], "review": [...]}.

    - unchanged: already matched (via a remembered zoho_vendor_map entry, or
      an exact company-name match) to a supplier whose stored contact/
      email/phone/address are identical to what's in the file -- nothing
      to do.
    - changed: same match, but at least one detail differs -- importing
      refreshes those details. A blank value in the file never overwrites
      an existing value (Zoho exports are often missing fields this app
      already has filled in by hand).
    - review: no exact match, but the vendor name is a close match to an
      existing supplier's name (e.g. "Currys" vs "Currys Business Sales")
      -- needs a human decision: the same supplier under a slightly
      different name (just map it), or a genuinely different company
      (create as new).
    - new: no match at all, close or exact -- safe to create as a new
      supplier automatically.

    Every entry carries "matched_supplier_id" (the supplier it will update,
    or None for new/review)."""
    vendor_map = get_zoho_vendor_map(conn)
    suppliers = list_suppliers(conn)
    by_id = {s["id"]: s for s in suppliers}
    by_name = {s["company_name"].strip().lower(): s for s in suppliers if s["company_name"].strip()}

    new_list, changed_list, unchanged_list, review_list = [], [], [], []
    for raw in zoho_vendors:
        v = dict(raw)
        name = v["company_name"]
        supplier = by_id.get(vendor_map.get(name)) or by_name.get(name.strip().lower())
        if supplier is not None:
            v["matched_supplier_id"] = supplier["id"]
            diffs = {}
            for field in ("contact_name", "email", "phone", "address"):
                new_val = (v.get(field) or "").strip()
                old_val = (supplier.get(field) or "").strip()
                if new_val and new_val != old_val:
                    diffs[field] = new_val
            v["diffs"] = diffs
            (changed_list if diffs else unchanged_list).append(v)
            continue

        na = _normalize_product_name(name)
        best_ratio, best_supplier = 0.0, None
        for s in suppliers:
            ns = _normalize_product_name(s["company_name"])
            if not ns:
                continue
            ratio = difflib.SequenceMatcher(None, na, ns).ratio()
            if na in ns or ns in na:
                ratio = max(ratio, 0.85)
            if ratio > best_ratio:
                best_ratio, best_supplier = ratio, s
        v["matched_supplier_id"] = None
        if best_supplier is not None and best_ratio >= min_ratio:
            v["closest_supplier_id"] = best_supplier["id"]
            v["closest_supplier_name"] = best_supplier["company_name"]
            v["match_ratio"] = best_ratio
            review_list.append(v)
        else:
            new_list.append(v)
    return {"new": new_list, "changed": changed_list, "unchanged": unchanged_list, "review": review_list}


def import_zoho_vendor_as_new(conn, zoho_vendor):
    """Create a new supplier from a resolved Zoho vendor row and remember
    the Zoho name -> supplier mapping so future imports (and PO/item
    imports that reference the same vendor name) recognise it
    automatically. Always passes the FULL current supplier list to
    replace_suppliers, per its upsert-by-company_name contract."""
    suppliers = list_suppliers(conn)
    suppliers.append({
        "company_name": zoho_vendor["company_name"],
        "contact_name": zoho_vendor.get("contact_name", ""),
        "email": zoho_vendor.get("email", ""),
        "cc_emails": "",
        "phone": zoho_vendor.get("phone", ""),
        "address": zoho_vendor.get("address", ""),
        "active": True,
    })
    replace_suppliers(conn, suppliers)
    new_sup = get_supplier_by_name(conn, zoho_vendor["company_name"])
    set_zoho_vendor_map(conn, zoho_vendor["company_name"], new_sup["id"])
    return new_sup


def apply_zoho_vendor_update(conn, zoho_vendor, supplier_id):
    """Refresh an existing supplier's contact details from a matched Zoho
    vendor row -- only fields the file actually has a value for are
    touched, and the Zoho name -> supplier mapping is (re)remembered so
    this vendor is recognised automatically from now on."""
    suppliers = list_suppliers(conn)
    for s in suppliers:
        if s["id"] == supplier_id:
            for field in ("contact_name", "email", "phone", "address"):
                new_val = (zoho_vendor.get(field) or "").strip()
                if new_val:
                    s[field] = new_val
            break
    replace_suppliers(conn, suppliers)
    set_zoho_vendor_map(conn, zoho_vendor["company_name"], supplier_id)


# ============================================================
# 11. Procurement savings tracking -- record a saving (negotiated price,
#     switching supplier, a bulk discount, avoided cost, and so on) against
#     a PO, a product, or both, as it happens, so the YTD figure is built
#     up through the year instead of reconstructed retrospectively. Where
#     there's a clear before/after order to compare, the amount is worked
#     out from the product's own price history across every order it's
#     ever appeared on (not just typed in by hand) -- a plain amount can
#     still be entered directly for cases like avoided cost that don't
#     have a matching order to compare against. A negative amount records
#     a price increase rather than a saving, so the running total stays
#     honest either way.
# ============================================================

def get_previous_price_for_product(conn, product_name, exclude_po_ref=None, before_at=None):
    """The most recent price actually paid for this product before the
    order currently being compared against -- across every supplier, since
    a saving might come from switching supplier rather than a repeat order
    with the same one. Returns the full history row (po_ref, supplier,
    price, qty, created_at) or None if there's no earlier order on record
    to compare against.

    before_at (an ISO created_at string) restricts this to orders that
    genuinely happened earlier -- without it, excluding just the PO's own
    po_ref isn't enough: if a *later* order for the same product happens to
    have been entered first, skipping only the exact po_ref would still
    pick that later order as the "previous" price, which is backwards."""
    for h in get_product_purchase_history(conn, product_name):
        if exclude_po_ref and h["po_ref"] == exclude_po_ref:
            continue
        if before_at and (h["created_at"] or "") >= before_at:
            continue
        return h
    return None


def list_savings_categories(conn, active_only=True):
    q = "SELECT * FROM savings_categories"
    if active_only:
        q += " WHERE active=1"
    q += " ORDER BY sort_order, id"
    return [dict(r) for r in conn.execute(q)]


def add_savings_category(conn, name):
    name = (name or "").strip()
    if not name:
        raise ValueError("Enter a category name.")
    existing = conn.execute(
        "SELECT id, active FROM savings_categories WHERE name=? COLLATE NOCASE", (name,)
    ).fetchone()
    if existing:
        if not existing["active"]:
            conn.execute("UPDATE savings_categories SET active=1 WHERE id=?", (existing["id"],))
            conn.commit()
        return
    max_order = conn.execute("SELECT COALESCE(MAX(sort_order), -1) AS m FROM savings_categories").fetchone()["m"]
    conn.execute("INSERT INTO savings_categories(name, sort_order, active) VALUES (?, ?, 1)", (name, max_order + 1))
    conn.commit()


def rename_savings_category(conn, category_id, new_name):
    new_name = (new_name or "").strip()
    if not new_name:
        raise ValueError("Enter a category name.")
    row = conn.execute("SELECT name FROM savings_categories WHERE id=?", (category_id,)).fetchone()
    if not row:
        return
    old_name = row["name"]
    # Tidy-up-pass fix (post-Batch-151): the UPDATE below can raise
    # sqlite3.IntegrityError if new_name collides with another category's
    # UNIQUE name -- a completely ordinary, expected user error (typing a
    # name that's already taken). Left unguarded, that raise used to leave
    # this connection's implicit transaction open (SQLite auto-rolls-back
    # only the one failed statement, not the transaction wrapper Python's
    # sqlite3 module tracks) -- and the NEXT, entirely unrelated write
    # anywhere else in the app that opens its own explicit transaction
    # (e.g. save_po's conn.execute("BEGIN")) would then fail with the
    # confusing "cannot start a transaction within a transaction" instead
    # of doing its actual job. Wrapping in try/except+rollback, same
    # pattern save_po/merge_products/merge_suppliers already use for their
    # own explicit transactions, so a caught, ordinary validation error
    # here can never poison some other, unrelated save later in the
    # session.
    try:
        conn.execute("UPDATE savings_categories SET name=? WHERE id=?", (new_name, category_id))
        conn.execute("UPDATE savings SET category=? WHERE category=?", (new_name, old_name))
        conn.commit()
    except Exception:
        conn.rollback()
        raise


def set_savings_category_active(conn, category_id, active):
    conn.execute("UPDATE savings_categories SET active=? WHERE id=?", (1 if active else 0, category_id))
    conn.commit()


def _compute_saving_amount(data):
    """Shared by record_saving/update_saving: works out the saved amount
    (and the previous/new unit price to store alongside it) from whatever
    the caller supplied -- a clear before/after price pair if there is
    one, otherwise a plain typed-in amount. Also returns has_pair, so the
    caller can require a written basis for the cases where the number
    isn't self-evidently backed by two real prices (see record_saving)."""
    qty = float(data.get("qty") or 0) or 1
    previous_price = data.get("previous_price")
    new_price = data.get("new_price")
    has_pair = previous_price not in (None, "") and new_price not in (None, "")
    if has_pair:
        previous_price = float(previous_price)
        new_price = float(new_price)
        amount = round((previous_price - new_price) * qty, 2)
    else:
        previous_price = float(previous_price) if previous_price not in (None, "") else None
        new_price = float(new_price) if new_price not in (None, "") else None
        amount = round(float(data.get("amount") or 0), 2)
    return previous_price, new_price, amount, has_pair


def saving_basis_text(row, currency):
    """A one-line, always-present explanation of how a saving's amount was
    worked out -- so the figure stays meaningful and demonstrable over
    time rather than just a number, per Yitzi's feedback that there should
    "always be a clear and consistent basis" behind it. That's either the
    actual before/after unit prices being compared (and the quantity they
    were applied to), when there's a real order on record to compare
    against, or the note the person entered by hand when there isn't --
    record_saving/update_saving now require that note precisely so this
    line is never blank. Used everywhere a saving is listed: the Savings
    page table and both periodic report emails."""
    previous_price, new_price = row.get("previous_price"), row.get("new_price")
    if previous_price not in (None, "") and new_price not in (None, ""):
        qty = row.get("qty") or 1
        return f"{money(float(previous_price), currency)} -> {money(float(new_price), currency)} (qty {float(qty):g})"
    return (row.get("note") or "").strip() or "(no basis recorded)"


def _require_saving_basis(has_pair, amount, note):
    """A flat typed-in amount has no automatic before/after comparison to
    show its working, so -- unlike a per-unit saving, where the two prices
    speak for themselves -- it needs a note explaining where the figure
    came from before it's recorded. Raises ValueError if one's missing."""
    if not has_pair and amount != 0 and not note:
        raise ValueError(
            "Add a note explaining the basis for this saving -- there's no before/after price to show it "
            "automatically, so a short reason (e.g. \"quote came in £40 under budget\") is what keeps the "
            "figure meaningful later."
        )


def record_saving(conn, data):
    """Logs a saving (or, with a negative amount, a price increase)
    against a PO, a product, or both. Returns the amount actually
    recorded."""
    category = (data.get("category") or "").strip()
    if not category:
        raise ValueError("Choose a savings category.")
    po_ref = (data.get("po_ref") or "").strip()
    product = (data.get("product") or "").strip()
    if not po_ref and not product:
        raise ValueError("Link the saving to a PO, a product, or both.")
    previous_price, new_price, amount, has_pair = _compute_saving_amount(data)
    note = (data.get("note") or "").strip()
    _require_saving_basis(has_pair, amount, note)
    qty = float(data.get("qty") or 0)
    now = now_iso()
    conn.execute(
        "INSERT INTO savings(category, po_ref, product, supplier_company_name, qty, "
        "previous_price, new_price, amount, note, created_at, deleted) "
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0)",
        (
            category, po_ref, product, (data.get("supplier_company_name") or "").strip(), qty,
            previous_price, new_price, amount, note, now,
        ),
    )
    conn.commit()
    return amount


def update_saving(conn, saving_id, data):
    category = (data.get("category") or "").strip()
    if not category:
        raise ValueError("Choose a savings category.")
    po_ref = (data.get("po_ref") or "").strip()
    product = (data.get("product") or "").strip()
    if not po_ref and not product:
        raise ValueError("Link the saving to a PO, a product, or both.")
    previous_price, new_price, amount, has_pair = _compute_saving_amount(data)
    note = (data.get("note") or "").strip()
    _require_saving_basis(has_pair, amount, note)
    qty = float(data.get("qty") or 0)
    conn.execute(
        "UPDATE savings SET category=?, po_ref=?, product=?, supplier_company_name=?, qty=?, "
        "previous_price=?, new_price=?, amount=?, note=? WHERE id=?",
        (
            category, po_ref, product, (data.get("supplier_company_name") or "").strip(), qty,
            previous_price, new_price, amount, note, saving_id,
        ),
    )
    conn.commit()
    return amount


def delete_saving(conn, saving_id):
    conn.execute("UPDATE savings SET deleted=1 WHERE id=?", (saving_id,))
    conn.commit()


def list_savings(conn, date_from=None, date_to=None, category=None, product=None, po_ref=None, deleted=False):
    # date(...) rather than a raw string compare -- created_at has a time
    # component (e.g. "2026-08-16T13:47:16") while date_to is often just a
    # plain "yyyy-mm-dd" from a date picker, and comparing those as text
    # would wrongly exclude anything from date_to's own day (a bare date
    # sorts before any timestamp on the same day).
    q = "SELECT * FROM savings WHERE deleted=?"
    params = [1 if deleted else 0]
    if date_from:
        q += " AND date(created_at) >= date(?)"
        params.append(date_from.isoformat() if hasattr(date_from, "isoformat") else date_from)
    if date_to:
        q += " AND date(created_at) <= date(?)"
        params.append(date_to.isoformat() if hasattr(date_to, "isoformat") else date_to)
    if category:
        q += " AND category = ?"
        params.append(category)
    if product:
        q += " AND product = ?"
        params.append(product)
    if po_ref:
        q += " AND po_ref = ?"
        params.append(po_ref)
    q += " ORDER BY created_at DESC"
    return [dict(r) for r in conn.execute(q, params)]


def savings_fiscal_year_bounds(conn, now=None):
    """Year-to-date window for the procurement savings KPI, respecting the
    configurable financial year start month (1=January, i.e. the plain
    calendar year, by default -- Settings > Savings & Scorecards)."""
    now = now or datetime.now()
    try:
        start_month = int(get_setting(conn, "savings_fiscal_year_start_month", "1") or 1)
    except ValueError:
        start_month = 1
    start_month = max(1, min(start_month, 12))
    year = now.year if now.month >= start_month else now.year - 1
    start = datetime(year, start_month, 1)
    end = datetime(year + 1, start_month, 1) - timedelta(seconds=1)
    return start, end


def get_savings_summary(conn, date_from=None, date_to=None):
    """Total saved (or, if negative, net lost to price increases) for a
    period, broken down by category/product/supplier -- the same data
    used for both the Dashboard tile and the Reports breakdown."""
    rows = list_savings(conn, date_from=date_from, date_to=date_to)
    total = round(sum(r["amount"] for r in rows), 2)
    by_category, by_product, by_supplier = {}, {}, {}
    for r in rows:
        cat = r["category"] or "Uncategorised"
        by_category[cat] = by_category.get(cat, 0) + r["amount"]
        if r["product"]:
            by_product[r["product"]] = by_product.get(r["product"], 0) + r["amount"]
        supp = r["supplier_company_name"] or "(no supplier)"
        by_supplier[supp] = by_supplier.get(supp, 0) + r["amount"]
    return {
        "total": total,
        "count": len(rows),
        "by_category": sorted(by_category.items(), key=lambda kv: -kv[1]),
        "by_product": sorted(by_product.items(), key=lambda kv: -kv[1])[:15],
        "by_supplier": sorted(by_supplier.items(), key=lambda kv: -kv[1])[:15],
        "rows": rows,
    }


def get_savings_ytd(conn, now=None):
    start, end = savings_fiscal_year_bounds(conn, now)
    summary = get_savings_summary(conn, date_from=start, date_to=end)
    summary["period_start"] = start
    summary["period_end"] = end
    return summary


# ============================================================
# 12. Purchase owner tracking -- who actually raised each PO (Yitzi,
#     another staff member, etc), so spend and pricing can be looked at
#     per-person as well as per-supplier. Two sources, in priority order:
#       1. Zoho's own "CF.Purchase Owner" field, when importing a Zoho
#          export that has it filled in.
#       2. The PO reference's own prefix letters (e.g. "YJ", "MR") against
#          a Settings-editable mapping (Settings > Purchase Owners) -- the
#          only option for POs raised directly in this app, which has never
#          carried a Zoho-style owner field of its own, and a safety net
#          for Zoho POs where the field was left blank.
#     Resolution happens centrally in save_po(), so every code path that
#     creates or re-saves a PO gets this automatically -- nothing about the
#     existing Zoho PO import flow (parse_zoho_export_csv/plan_zoho_import/
#     the Import/Export page UI) needed to change for this to work.
# ============================================================

_PO_REF_PREFIX_RE = re.compile(r'^(?:PO:)?([A-Za-z]+)')


def extract_po_ref_prefix(po_ref):
    """The letters right at the start of a PO reference (after the "PO:"
    if present, before the first digit) -- "YJ" from "PO:YJ140826-9101",
    "MR" from "MR060126-N320", etc. Returns "" if the reference doesn't
    start with any letters at all."""
    if not po_ref:
        return ""
    m = _PO_REF_PREFIX_RE.match(po_ref.strip())
    return m.group(1).upper() if m else ""


def list_purchase_owner_prefixes(conn):
    return [dict(r) for r in conn.execute(
        "SELECT * FROM purchase_owner_prefixes ORDER BY sort_order, id"
    )]


def set_purchase_owner_prefix(conn, prefix, name):
    """Adds a new prefix -> name mapping, or renames the person already
    mapped to that prefix if it exists (matched case-insensitively, since
    PO refs aren't reliably one case or the other)."""
    prefix = (prefix or "").strip().upper()
    name = (name or "").strip()
    if not prefix or not name:
        raise ValueError("Enter both a prefix and a name.")
    existing = conn.execute(
        "SELECT id FROM purchase_owner_prefixes WHERE prefix = ? COLLATE NOCASE", (prefix,)
    ).fetchone()
    if existing:
        conn.execute("UPDATE purchase_owner_prefixes SET name=? WHERE id=?", (name, existing["id"]))
    else:
        max_order = conn.execute(
            "SELECT COALESCE(MAX(sort_order), -1) AS m FROM purchase_owner_prefixes"
        ).fetchone()["m"]
        conn.execute(
            "INSERT INTO purchase_owner_prefixes(prefix, name, sort_order) VALUES (?, ?, ?)",
            (prefix, name, max_order + 1),
        )
    conn.commit()


def delete_purchase_owner_prefix(conn, prefix_id):
    conn.execute("DELETE FROM purchase_owner_prefixes WHERE id=?", (prefix_id,))
    conn.commit()


def resolve_purchase_owner(conn, po_ref, zoho_owner_name=None):
    """The purchase owner for a PO -- Zoho's own field if it was supplied
    (from an import), otherwise derived from the PO reference's own prefix
    letters via the Settings > Purchase Owners mapping. Returns "" if
    neither source resolves to anything (e.g. an unmapped prefix, or a
    one-off manually typed reference with no recognisable prefix at all)
    -- never guesses."""
    zoho_owner_name = (zoho_owner_name or "").strip()
    if zoho_owner_name:
        return zoho_owner_name
    prefix = extract_po_ref_prefix(po_ref)
    if not prefix:
        return ""
    row = conn.execute(
        "SELECT name FROM purchase_owner_prefixes WHERE prefix = ? COLLATE NOCASE", (prefix,)
    ).fetchone()
    return row["name"] if row else ""


def learn_purchase_owner_prefix_from_zoho(conn, po_ref, zoho_owner_name, commit=True):
    """When a Zoho PO import carries its own "CF.Purchase Owner" value and
    the reference's prefix letters (eg. "YJ") aren't in the Settings >
    Purchase Owners table yet, remembers that prefix -> name pairing
    automatically -- so it shows up in Settings ready-made, and so later
    POs sharing the same prefix (including ones where the Zoho export left
    the owner field blank) resolve to the right person without it having
    to be typed in by hand first. Never touches a prefix that's already
    mapped, even to a different name -- overwriting one is always a
    deliberate Settings edit, not something an import should do silently."""
    zoho_owner_name = (zoho_owner_name or "").strip()
    if not zoho_owner_name:
        return
    prefix = extract_po_ref_prefix(po_ref)
    if not prefix:
        return
    existing = conn.execute(
        "SELECT id FROM purchase_owner_prefixes WHERE prefix = ? COLLATE NOCASE", (prefix,)
    ).fetchone()
    if existing:
        return
    max_order = conn.execute(
        "SELECT COALESCE(MAX(sort_order), -1) AS m FROM purchase_owner_prefixes"
    ).fetchone()["m"]
    conn.execute(
        "INSERT INTO purchase_owner_prefixes(prefix, name, sort_order) VALUES (?, ?, ?)",
        (prefix, zoho_owner_name, max_order + 1),
    )
    if commit:
        conn.commit()


def backfill_purchase_owners(conn, commit=True):
    """Fills in purchase_owner for every existing PO that doesn't have one
    yet, using its reference prefix against the current mapping -- for POs
    saved before this feature existed, or before a mapping covered their
    prefix. Never overwrites a purchase_owner that's already set, and
    leaves a PO alone if its prefix still isn't mapped, so it's always
    safe to re-run (e.g. right after adding a new mapping). Returns how
    many POs were updated."""
    updated = 0
    rows = conn.execute(
        "SELECT id, po_ref FROM purchase_orders WHERE purchase_owner IS NULL OR purchase_owner = ''"
    ).fetchall()
    for r in rows:
        owner = resolve_purchase_owner(conn, r["po_ref"])
        if owner:
            conn.execute("UPDATE purchase_orders SET purchase_owner=? WHERE id=?", (owner, r["id"]))
            updated += 1
    if commit:
        conn.commit()
    return updated


def report_spend_by_purchase_owner(conn, date_from=None, date_to=None, deleted=False):
    """Order count/spend per purchase owner -- the "By Purchase Owner"
    Reports tab's summary table, same shape as report_spend_by_supplier."""
    sql = (
        "SELECT purchase_owner AS owner, COUNT(*) AS po_count, "
        "COALESCE(SUM(base_total), 0) AS total, COALESCE(AVG(base_total), 0) AS avg_po_value "
        "FROM purchase_orders WHERE deleted=?"
    )
    params = [1 if deleted else 0]
    if date_from:
        sql += " AND date(created_at) >= date(?)"
        params.append(date_from)
    if date_to:
        sql += " AND date(created_at) <= date(?)"
        params.append(date_to)
    sql += " GROUP BY owner ORDER BY total DESC"
    return [dict(r) for r in conn.execute(sql, params)]


def report_price_by_owner(conn, months=None):
    """For every product bought by more than one purchase owner, compares
    what each of them most recently paid for it -- so it's visible who's
    buying at higher prices and who's getting the better deals, the same
    idea as get_supplier_price_competitiveness but grouped by person
    instead of by supplier. One row per (product, owner); sorted so the
    biggest overpay against that product's cheapest owner comes first."""
    if months is None:
        try:
            months = int(get_setting(conn, "scorecard_price_window_months", "12") or 12)
        except ValueError:
            months = 12
    cutoff = (datetime.now() - timedelta(days=max(months, 1) * 30)).isoformat()
    rows = conn.execute(
        # pi.price * po.fx_rate, same reasoning as get_supplier_price_competitiveness --
        # compares prices in the home currency so a foreign-currency order
        # doesn't look artificially cheap or expensive next to a home-currency one.
        "SELECT pi.product AS product, po.purchase_owner AS owner, pi.price * po.fx_rate AS price, "
        "po.created_at AS created_at FROM po_items pi JOIN purchase_orders po ON po.id = pi.po_id "
        "WHERE po.deleted = 0 AND po.created_at >= ? AND po.purchase_owner != '' "
        "ORDER BY po.created_at DESC",
        (cutoff,),
    ).fetchall()
    latest = {}
    for r in rows:
        key = (r["product"], r["owner"])
        if key not in latest:  # newest-first, so the first hit per key is the latest price
            latest[key] = float(r["price"] or 0)
    by_product = {}
    for (product, owner), price in latest.items():
        by_product.setdefault(product, {})[owner] = price

    out = []
    for product, by_owner in by_product.items():
        if len(by_owner) < 2:
            continue
        cheapest_owner = min(by_owner, key=by_owner.get)
        cheapest_price = by_owner[cheapest_owner]
        for owner, price in by_owner.items():
            out.append({
                "product": product, "owner": owner, "price": price,
                "cheapest_owner": cheapest_owner, "cheapest_price": cheapest_price,
                "difference": price - cheapest_price,
            })
    out.sort(key=lambda r: r["difference"], reverse=True)
    return out


def purchase_owner_prefix_is_unmapped(conn, po_ref):
    """True when a PO reference has an extractable prefix (e.g. "YJ" in
    PO:YJ140826-0933) but that prefix has no entry in
    purchase_owner_prefixes -- the kind of gap worth warning about (a new
    starter's initials not added yet), as opposed to a reference with no
    recognisable prefix at all, which isn't something to nag about."""
    prefix = extract_po_ref_prefix(po_ref)
    if not prefix:
        return False
    row = conn.execute(
        "SELECT 1 FROM purchase_owner_prefixes WHERE prefix = ? COLLATE NOCASE", (prefix,)
    ).fetchone()
    return row is None

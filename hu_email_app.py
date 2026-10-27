"""
Hungarian cold-EMAIL sender — live, one-lead-at-a-time review app.

Browser UI shaped like the DM sender's own Owner-Emails flow (dm_app.py's
/email_load, /email_next): load an Excel, generate one email at a time,
review/edit it, then explicitly click Send — one real Gmail API send per
click, nothing goes out on its own.

DESIGN — why this is a separate file from dm_app.py / dm_app_hu.py:
It reuses hu_email_gen.py's per-lead pipeline (scrape, screenshot, guarded
opener, subject/body) instead of duplicating it. hu_email_gen.py itself only
IMPORTS dm_app_hu (which patches dm_app in memory, for this process only) —
nothing in this file, or in hu_email_gen.py, ever writes to dm_app.py or
dm_app_hu.py on disk. The DM sender's own files and its own Flask app
(dm_app.py, port 5002) are completely untouched by any of this.

Sends via gmail_send.py (the Gmail API), not the old mail.google.com compose
deep-link — that trick can only carry a plain-text body, and the whole point
of this version is a real embedded hero-image screenshot of the lead's own
site, which needs a real MIME message.

Doing one lead at a time (rather than the batch script's 5 parallel workers)
also sidesteps the Groq rate-limit problem the batch run hit: a human
reviewing and sending each email is far slower than the 8s pacing the batch
script needs to enforce artificially.

Run via START_EMAIL_HU.bat, or directly:
    py hu_email_app.py
Then open http://localhost:5003
"""

import json as _json
import os
import re as _re
import threading as _threading
from concurrent.futures import ThreadPoolExecutor as _ThreadPoolExecutor
from datetime import datetime, timedelta
from pathlib import Path as _Path

from flask import Flask, request, jsonify, render_template
import openpyxl

import hu_email_gen as heg       # the batch generator — reused for its pipeline
import hu_followup               # follow-up copy + schedule (no I/O of its own)
import hu_email_mech as mech     # Autószerelő mode: its own opener, body, follow-ups
import hu_email_ajanlat as ajanlat   # Job-ad mode: opener from the ad, no scraping
import gmail_send                # real Gmail API sending, hero image included
import dm_app                    # only for copy_to_clipboard

app = Flask(__name__)

PORT = 5003

# ── Domain warmup — a fresh sending domain/account, not a Groq quota ───────
# milan@landscale.net has zero prior sending history. Gmail's technical daily
# cap on a paid Workspace account (2000 recipients/day) is nowhere near the
# real constraint here — a brand-new domain suddenly sending at that volume
# is one of the most reliable ways to get flagged as spam regardless of the
# account tier. This is "domain warming": ramp volume up over weeks so
# inbox providers build trust in the sender. Ramps are keyed off calendar
# days since THAT ACCOUNT's first recorded send (read from the sheet
# itself, not a constant you'd have to remember to bump by hand — see
# _current_daily_cap), and differ by account type: a paid Workspace domain
# can keep climbing, but a free @gmail.com's safe cold-email ceiling is
# ~25/day permanently (2026 deliverability consensus: woodpecker.co,
# smartlead.ai, prospeo.io) — its official 500/day cap is irrelevant, and
# pushing volume on a free account risks the account itself, with no
# admin console behind it to appeal through.
_RAMP_SCHEDULES = {
    # Ceiling raised 30 -> 50 on request (2026-09-22). The reasoning for the
    # old 30 still stands and is kept here on purpose, because it is the thing
    # to re-read before anyone raises it again:
    #
    #   Workspace's TECHNICAL cap is 2000 recipients/day, but that number is
    #   irrelevant to cold mail. Practitioner consensus for a single seasoned
    #   inbox sending cold is 20-50/day, with the conservative end (20-30) for
    #   a solo sender whose domain has no other traffic to average against.
    #   A low-volume sender is MORE fragile on spam complaints, not less:
    #   Gmail wants the reported rate under 0.30% (ideally 0.10%), so at 30
    #   sends/day a single "report spam" click is already ~3% for that day.
    #   Volume was never this funnel's bottleneck either ([[outreach-
    #   bottleneck-is-the-close]]: ~40 positive replies produced 1 booking).
    #
    # 50 is the TOP of that consensus band, not past it, and this inbox has
    # been sending daily since 2026-08-05, so it is a seasoned inbox rather
    # than a warming one. What has NOT changed: landscale.net is the only
    # sending domain here, it has no other traffic to dilute a complaint
    # against, and there is no second domain to fall back on if it gets
    # burnt. Past ~50/inbox the honest answer is another domain and inbox,
    # not a bigger number on this one.
    #
    # The two new bands exist so the jump is a glide rather than a cliff: a
    # 30 -> 50 step overnight is exactly the volume spike filters score on.
    "workspace": [
        (7,  10),   # week 1: up to 10/day
        (14, 20),   # week 2: up to 20/day
        (21, 25),   # week 3: up to 25/day
        (35, 30),   # weeks 4-5: 30/day
        (49, 40),   # weeks 6-7: 40/day
        (None, 50),  # week 8+: 50/day — deliberate stop, not a technical limit
    ],
    "free_gmail": [
        (7,  10),   # week 1: up to 10/day
        (14, 20),   # week 2: up to 20/day
        (None, 25),  # week 3+: 25/day — permanent ceiling on free Gmail
    ],
}


def _ramp_for(account: str) -> list:
    acct_type = gmail_send.ACCOUNTS.get(account, {}).get("type", "workspace")
    return _RAMP_SCHEDULES[acct_type]

_state = {
    "excel_path":  None,
    "wb":          None,
    "ws":          None,
    "cols":        {},
    "current_row": None,
    "current_lead": None,   # dict with email/subject/body for the pending review
    "current_hero": None,   # raw JPEG bytes for the pending review, or None
    "account": gmail_send.DEFAULT_ACCOUNT,   # which sending identity is selected in the UI
    # "builders" = the original pitch (hu_email_gen), "mech" = autószerelők
    # (hu_email_mech), "ajanlat" = companies hiring a quote writer
    # (hu_email_ajanlat). Chosen per loaded sheet in the UI; everything else
    # (cap, follow-up schedule, reply checks) is shared.
    "mode": "builders",
    # sent_today is NOT tracked here — it's read fresh from the Excel file's
    # timestamped "Sent" markers each time (_todays_send_times), so the
    # daily cap survives an app restart instead of quietly resetting to 0.
}


def _resolve_account(data: dict) -> str:
    """Account key from the request body, falling back to whatever was last
    selected — same pattern as subject/body already being editable-but-
    defaulted in /send. Unknown keys fall back too rather than 500ing on a
    stale/garbled client value."""
    acct = (data.get("account") or "").strip()
    if acct in gmail_send.ACCOUNTS:
        _state["account"] = acct
        return acct
    return _state["account"]

_SENT_VALUES = ("yes", "y", "sent", "1", "true", "skip")
# Matches both the plain old marker ("Sent", "Skip" — rows sent before this
# timestamped version existed), the timestamped-only form ("Sent 2026-08-05
# 14:32" — written before multi-account support existed, always the
# landscale account), and the current timestamped+account form
# ("Sent 2026-08-06 10:15 [gmail]"). Group 3 is the account key, or None for
# the two older forms (both of which only ever happened on "landscale").
_SENT_PREFIX_RE = _re.compile(r"^(sent|skip|yes|y|1|true)\b", _re.IGNORECASE)
_SEND_TS_RE = _re.compile(
    r"^sent\s+(\d{4}-\d{2}-\d{2})\s+(\d{2}:\d{2})(?:\s*\[(\w+)\])?", _re.IGNORECASE)


def _is_handled(value: str) -> bool:
    """True for any cell value meaning 'already sent or skipped', old bare
    form or new timestamped form alike."""
    return bool(_SENT_PREFIX_RE.match((value or "").strip()))


# ─── Legal targeting guard (Hungary: Grt. §6 is strict opt-in) ─────────────
# Hungarian advertising law requires PRIOR EXPRESS CONSENT to send advertising
# email to a natural person — and an egyéni vállalkozó is a natural person, as
# is any named-individual address (kovacs.janos@ceg.hu), regardless of it being
# printed on a company website. A company's own functional contact point
# (info@, iroda@, office@) is the category that can be approached cold.
#
# That distinction is invisible in the sheet and easy to lose track of at 20-30
# sends a day, so it's computed per lead and shown in the review UI before the
# Send button — the human in the loop is the enforcement point, exactly like it
# already is for the copy itself. Deliberately a WARNING, not a hard block:
# plenty of real functional addresses won't be in the list below
# (cegnev@cegnev.hu is common for small HU firms), and only the person reading
# the actual website can tell those apart from a personal one.
_FUNCTIONAL_LOCALPARTS = {
    "info", "iroda", "office", "contact", "kapcsolat", "hello", "mail", "email",
    "ugyfelszolgalat", "ertekesites", "sales", "rendeles", "megrendeles",
    "titkarsag", "admin", "posta", "cegugy", "arajanlat", "ajanlat", "szerviz",
    "support", "help", "keszlet", "webshop", "shop", "boltunk", "munka",
    # Added after profiling a real 150-lead list, where each of these turned
    # out to be a company mailbox that the first version wrongly flagged.
    "ajanlatkeres", "arajanlatkeres", "kivitelezes", "bolt", "uzlet", "szalon",
    "berles", "beszerzes", "penzugy", "szamlazas", "marketing", "web", "cegunk",
}

# Free/consumer mailbox providers. An address here can still be a company's
# only contact point, but it can never be evidence of one: nothing about
# "kissjanos@gmail.com" identifies a company, so the natural-person reading
# is the safe default and the local-part heuristics below are skipped.
_FREE_MAIL_DOMAINS = {
    "gmail.com", "googlemail.com", "freemail.hu", "citromail.hu", "indamail.hu",
    "t-online.hu", "vipmail.hu", "yahoo.com", "yahoo.co.uk", "hotmail.com",
    "outlook.com", "outlook.hu", "live.com", "icloud.com", "protonmail.com",
    "chello.hu", "invitel.hu", "upcmail.hu", "digikabel.hu", "externet.hu",
}

_ACCENT_MAP = str.maketrans("áéíóöőúüűÁÉÍÓÖŐÚÜŰ", "aeiooouuuAEIOOOUUU")


def _norm_local(s: str) -> str:
    return _re.sub(r"[^a-z]", "", (s or "").lower().translate(_ACCENT_MAP))


def _address_risk(email: str) -> str:
    """'' if the address looks like a company contact point, else a Hungarian
    warning string for the review UI.

    Three ways an address clears:
      1. a known functional local part (info@, iroda@, ajanlatkeres@ ...)
      2. the local part IS the company — 'bacshaz@bacshaz.hu',
         'sanglass@sanglass.hu', 'trior@tri-or.hu'. A mailbox named after
         the firm is the firm's mailbox, not a person's, and small Hungarian
         companies use this shape constantly.
      3. the local part is a city — 'eger@palatinusfurdoszoba.hu',
         'kaposvar@webepker.hu' — i.e. a branch mailbox.
    All three are skipped outright on a free consumer provider, where the
    domain says nothing about a company either way.
    """
    addr = (email or "").strip().lower()
    if "@" not in addr:
        return ""
    local_raw, domain = addr.split("@", 1)
    base = _norm_local(local_raw)

    warning = ("Ez nem tipikus céges cím (info@, iroda@, office@). Ha egyéni "
               "vállalkozóé vagy egy konkrét személyé, a Grt. 6. § szerint "
               "előzetes hozzájárulás kellene a reklámlevélhez. Nézd meg a "
               "weboldalt: ha a cég hivatalos kapcsolattartó címe, mehet — ha "
               "egy magánszemélyé, inkább hagyd ki.")

    if domain in _FREE_MAIL_DOMAINS:
        return ("Ingyenes postafiók ({}), nem saját céges domain. Magyar "
                "szakiparban ez jellemzően egyéni vállalkozó vagy "
                "magánszemély, akinek a Grt. 6. § szerint előzetes "
                "hozzájárulás nélkül nem küldhető reklámlevél. Ez a "
                "leadlista legkockázatosabb kategóriája — hagyd ki, hacsak "
                "nem látod a weboldalon kifejezetten cégként "
                "megadva.").format(domain)

    if not base or base in _FUNCTIONAL_LOCALPARTS:
        return ""

    # 2 — local part is (a prefix of) the company's own domain, or vice versa.
    domain_root = _norm_local(domain.rsplit(".", 2)[0] if domain.count(".") > 1
                              else domain.split(".", 1)[0])
    if domain_root and (base == domain_root
                        or base.startswith(domain_root)
                        or domain_root.startswith(base)):
        return ""

    # 3 — branch mailbox named after a city.
    # via heg, which already imports dm_app_hu — see this file's design note.
    if base in {_norm_local(c) for c in heg.dm_app_hu._CITY_RAGOS}:
        return ""

    return warning


# ─── Excel I/O — live workbook, one column added for send-tracking ─────────

def _load_excel(path: str):
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    headers = [str(c.value or "").strip().lower() for c in ws[1]]

    cols = {}
    for key, needles in heg._COL_RULES:
        for i, h in enumerate(headers):
            if i in cols.values():
                continue
            if any(n in h for n in needles):
                cols[key] = i
                break

    # The loose "url" needle in heg._COL_RULES grabs a "Facebook URL" column
    # when it comes first, which hides the real Website column. A header
    # that literally says website/weboldal/honlap always wins.
    real_site = next((i for i, h in enumerate(headers)
                      if any(n in h for n in ("website", "weboldal", "honlap"))), None)
    if real_site is not None:
        cols["website"] = real_site

    sent_idx = next((i for i, h in enumerate(headers)
                      if "elküldve" in h or "küldve" in h or "kuldve" in h or "sent" in h), None)
    if sent_idx is None:
        ws.cell(row=1, column=len(headers) + 1, value="Elküldve?")
        headers.append("elküldve?")
        sent_idx = len(headers) - 1
        wb.save(path)
    cols["sent"] = sent_idx

    # Where the lead came from ("Listed on Q-Service network page"). Only the
    # mechanics opener reads it, as its fallback fact.
    src = next((i for i, h in enumerate(headers)
                if "verified" in h or "forrás" in h or "forras" in h), None)
    if src is not None:
        cols["source"] = src

    # Ajánlat mode only: the job ad's position and what they quote for.
    for key, needles in (("position",   ("pozíció", "pozicio", "position")),
                         ("quote_what", ("mit árajánlatoz", "mit arajanlatoz")),
                         ("opener_custom", ("nyitómondat", "nyitomondat"))):
        idx = next((i for i, h in enumerate(headers) if any(n in h for n in needles)), None)
        if idx is not None:
            cols[key] = idx

    # Follow-up bookkeeping. Added on load like "Elküldve?" already was, so
    # an existing sheet mid-campaign gains them without anything to migrate:
    # rows sent before this feature simply have blank thread ids, which
    # _followup_queue treats as "cannot verify replies, ask before sending".
    #   szál       "<threadId>|<Message-ID header>" — threading + reply checks
    #   válaszolt  "Igen 2026-08-14" once a reply is seen; stops the sequence
    #   követés    "F1 2026-08-15 10:00 [landscale]; F2 ..." — one line per step
    added = False
    for key, title, needles in (
        ("thread",  "Szál",         ("szál", "szal", "thread")),
        ("replied", "Válaszolt?",   ("válaszolt", "valaszolt", "replied")),
        ("fups",    "Követés",      ("követés", "kovetes", "followup", "follow-up")),
        # "Nem kézbesíthető 2026-08-12" once a bounce is seen. Its own column
        # rather than a value inside "Válaszolt?" so the sheet stays readable
        # by eye — a bounce and a reply are opposite outcomes and shouldn't
        # share a cell.
        ("bounced", "Kézbesítés",   ("kézbesít", "kezbesit", "bounce")),
    ):
        idx = next((i for i, h in enumerate(headers) if any(n in h for n in needles)), None)
        if idx is None:
            ws.cell(row=1, column=len(headers) + 1, value=title)
            headers.append(title.lower())
            idx = len(headers) - 1
            added = True
        cols[key] = idx
    if added:
        wb.save(path)
    return wb, ws, cols


def _save_wb():
    import time as _t
    wb, path = _state["wb"], _state["excel_path"]
    if wb is None or not path:
        return
    for _ in range(6):
        try:
            wb.save(path)
            return
        except PermissionError:
            _t.sleep(0.4)
    raise PermissionError(
        "Nem tudok írni az Excel fájlba — nyitva van Excelben, vagy még "
        "szinkronizál a OneDrive. Zárd be, várj egy kicsit, aztán próbáld újra.")


def _mark(row: int, value: str):
    ws, cols = _state["ws"], _state["cols"]
    ws.cell(row=row, column=cols["sent"] + 1, value=value)
    _save_wb()


_MECH_NICHE_WORDS = ("autószerel", "autoszerel", "autószerviz", "autoszerviz",
                     "mechanic", "car repair")


def _max_step() -> int:
    """Last follow-up step for the current mode. Ajánlat stops after two."""
    return ajanlat.MAX_STEP if _state.get("mode") == "ajanlat" else hu_followup.MAX_STEP


def _row_in_mode(ws_row, cols) -> bool:
    """One workbook can hold both campaigns (hun_all_email_leads 10 08.xlsx
    has mechanics AND builders). Mechanics mode only sees mechanic rows and
    builders mode only sees the rest, so neither email or follow-up ever
    reaches the other trade. Sheets with no Niche column are all one mode.
    Ajánlat mode only sees rows with a Pozíció, the one fact its opener needs."""
    ni = cols.get("niche")
    if _state.get("mode") == "ajanlat":
        pi = cols.get("position")
        return pi is not None and bool(str(ws_row[pi].value or "").strip())
    if ni is None:
        return True
    niche = str(ws_row[ni].value or "").lower()
    is_mech = any(w in niche for w in _MECH_NICHE_WORDS)
    return is_mech if _state.get("mode") == "mech" else not is_mech


def _upcoming_leads(ws, cols, limit=6) -> list:
    """The next N unhandled leads, in sheet order. Feeds the prefetcher."""
    out = []
    for row in ws.iter_rows(min_row=2):
        excel_row = row[0].row

        def v(key, _row=row):
            i = cols.get(key)
            return str(_row[i].value or "").strip() if i is not None else ""

        email = v("email")
        if not email or "@" not in email or _is_handled(v("sent")):
            continue
        if not _row_in_mode(row, cols):
            continue
        out.append({
            "row": excel_row, "email": email, "website": v("website"),
            "owner": v("owner"), "name": v("name"),
            "city": heg._clean_city(v("city")), "niche": v("niche"),
            "source": v("source"),
            "position": v("position"), "quote_what": v("quote_what"),
            "opener_custom": v("opener_custom"),
        })
        if len(out) >= limit:
            break
    return out


def _next_lead(ws, cols, skip_row=None) -> dict | None:
    for row in ws.iter_rows(min_row=2):
        excel_row = row[0].row
        if skip_row is not None and excel_row == skip_row:
            continue

        def v(key):
            i = cols.get(key)
            return str(row[i].value or "").strip() if i is not None else ""

        email = v("email")
        if not email or "@" not in email:
            continue
        if _is_handled(v("sent")):
            continue
        if not _row_in_mode(row, cols):
            continue
        return {
            "row":     excel_row,
            "email":   email,
            "website": v("website"),
            "owner":   v("owner"),
            "name":    v("name"),
            "city":    heg._clean_city(v("city")),
            "niche":   v("niche"),
            "source":  v("source"),
            "position":   v("position"),
            "quote_what": v("quote_what"),
            "opener_custom": v("opener_custom"),
        }
    return None


_SENT_NOT_SKIP_RE = _re.compile(r"^(sent|yes|y|1|true)\b", _re.IGNORECASE)


def _counts(ws, cols):
    rows = [r for r in ws.iter_rows(min_row=2) if _row_in_mode(r, cols)]
    total = sum(1 for r in rows if "@" in str(r[cols["email"]].value or ""))
    sent = sum(1 for r in rows
               if _SENT_NOT_SKIP_RE.match(str(r[cols["sent"]].value or "").strip()))
    return total, sent


def _marker_account(m) -> str:
    """Account key from a matched _SEND_TS_RE — markers written before
    multi-account support existed have no [tag] and were always sent from
    landscale, so that's the correct default, not an arbitrary one."""
    return (m.group(3) or gmail_send.DEFAULT_ACCOUNT).lower()


# Follow-up markers inside the "Követés" cell: "F2 2026-08-15 10:04 [gmail]".
# Several can share one cell, separated by "; ", so this is used with
# finditer rather than match.
_FUP_TS_RE = _re.compile(
    r"F(\d)\s+(\d{4}-\d{2}-\d{2})\s+(\d{2}:\d{2})(?:\s*\[(\w+)\])?", _re.IGNORECASE)


def _todays_send_times(ws, cols, account: str) -> list:
    """Every real send today on ONE account — first touches AND follow-ups.

    Read from the Excel file rather than in-memory state, so the daily cap
    survives closing and reopening the app. Each sending identity warms up on
    its own reputation, so landscale's history must never count against
    gmail's cap or vice versa.

    Follow-ups are counted here deliberately: they are ordinary outbound mail
    as far as Gmail and the receiving server are concerned, so a cap that
    ignored them would quietly send double the intended volume on any day the
    queue was busy — exactly the spike the warmup ramp exists to prevent."""
    today = datetime.now().date()
    bi = cols.get("bounced")
    times = []
    for r in ws.iter_rows(min_row=2):
        # A bounced address never reached an inbox, so it never consumed any
        # of the day's real delivered volume — the thing the warmup ramp is
        # actually rationing. Excluding it gives the day's quota back, which
        # is what you'd want after five dead addresses out of twelve.
        #
        # Worth being clear about what this does NOT mean: a bounce is worse
        # for sender reputation than a delivered email, not better. This
        # frees the CAP, it does not make the bounce harmless — see the
        # bounce-rate warning in _delivery_stats.
        if bi is not None and str(r[bi].value or "").strip():
            continue

        raw = str(r[cols["sent"]].value or "")
        m = _SEND_TS_RE.match(raw.strip())
        if m and _marker_account(m) == account:
            try:
                dt = datetime.strptime(f"{m.group(1)} {m.group(2)}", "%Y-%m-%d %H:%M")
                if dt.date() == today:
                    times.append(dt)
            except ValueError:
                pass

        fi = cols.get("fups")
        if fi is None:
            continue
        for fm in _FUP_TS_RE.finditer(str(r[fi].value or "")):
            acct = (fm.group(4) or gmail_send.DEFAULT_ACCOUNT).lower()
            if acct != account:
                continue
            try:
                dt = datetime.strptime(f"{fm.group(2)} {fm.group(3)}", "%Y-%m-%d %H:%M")
                if dt.date() == today:
                    times.append(dt)
            except ValueError:
                pass
    return _merge_logged_sends(times, account)


# Cross-SHEET send log. The builders' list and the mechanics' list are two
# different workbooks sending from the same inbox, and the cap above only
# sees the workbook that is loaded. Without this, 30 sends from one sheet and
# 30 from the other would both pass a 50/day cap. The log is per account and
# only keeps the last few days.
_SEND_LOG_PATH = _Path(__file__).parent / "send_log.json"


def _log_send(account: str) -> None:
    try:
        log = _json.loads(_SEND_LOG_PATH.read_text(encoding="utf-8"))
    except Exception:
        log = {}
    cutoff = (datetime.now() - timedelta(days=3)).strftime("%Y-%m-%d")
    entries = [t for t in log.get(account, []) if t[:10] >= cutoff]
    entries.append(datetime.now().strftime("%Y-%m-%d %H:%M"))
    log[account] = entries
    try:
        _SEND_LOG_PATH.write_text(_json.dumps(log, indent=1), encoding="utf-8")
    except Exception:
        pass   # the cap is not worth failing a send over


def _merge_logged_sends(sheet_times: list, account: str) -> list:
    """Sheet timestamps plus any logged sends today that this sheet does not
    account for (i.e. ones made from the other workbook). Matched per minute
    so a send present in both is counted once."""
    try:
        log = _json.loads(_SEND_LOG_PATH.read_text(encoding="utf-8"))
    except Exception:
        return sorted(sheet_times)
    today = datetime.now().strftime("%Y-%m-%d")
    from collections import Counter
    sheet_c = Counter(t.strftime("%Y-%m-%d %H:%M") for t in sheet_times)
    log_c = Counter(t for t in log.get(account, []) if t.startswith(today))
    out = list(sheet_times)
    for minute, n in log_c.items():
        for _ in range(max(0, n - sheet_c.get(minute, 0))):
            out.append(datetime.strptime(minute, "%Y-%m-%d %H:%M"))
    return sorted(out)


# Warmup start dates, persisted OUTSIDE any workbook.
#
# _first_send_date() used to scan only the loaded worksheet, despite a
# docstring promising it spanned workbooks. Real consequence: the landscale
# account first sent on 2026-08-05, but opening a fresh lead sheet on
# 2026-08-24 made day 20 of the ramp look like day 1, throttling a
# nineteen-day-old sending domain back to the 10/day beginner cap. Starting a
# new sheet silently reset the warmup every time.
#
# The ramp is a property of the SENDING DOMAIN, not of whichever spreadsheet
# happens to be open, so the start date has to outlive the workbook.
_WARMUP_STATE_PATH = _Path(__file__).parent / "warmup_state.json"


def _load_warmup_state() -> dict:
    try:
        return _json.loads(_WARMUP_STATE_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _remember_warmup_start(account: str, d) -> None:
    """Record the earliest send date ever seen for this account. Only ever
    moves EARLIER, so re-opening an old workbook can correct the record but
    a new one can never push the ramp back to day 1."""
    state = _load_warmup_state()
    iso = d.isoformat()
    if state.get(account) is None or iso < state[account]:
        state[account] = iso
        try:
            _WARMUP_STATE_PATH.write_text(_json.dumps(state, indent=2), encoding="utf-8")
        except Exception:
            pass   # a warmup cap is not worth failing a send over


def _first_send_date(ws, cols, account: str):
    """Earliest timestamped 'Sent' marker for ONE account — that account's
    real first-send date, however many workbooks/restarts ago that was.
    Returns a date, or None if that account has never sent yet."""
    earliest = None
    for r in ws.iter_rows(min_row=2):
        raw = str(r[cols["sent"]].value or "")
        m = _SEND_TS_RE.match(raw.strip())
        if not m or _marker_account(m) != account:
            continue
        try:
            d = datetime.strptime(m.group(1), "%Y-%m-%d").date()
        except ValueError:
            continue
        if earliest is None or d < earliest:
            earliest = d

    # Reconcile with the cross-workbook record: whichever is earlier wins, and
    # the result is written back so the next sheet inherits it.
    if earliest is not None:
        _remember_warmup_start(account, earliest)
    stored = _load_warmup_state().get(account)
    if stored:
        try:
            stored_d = datetime.strptime(stored, "%Y-%m-%d").date()
            earliest = stored_d if earliest is None or stored_d < earliest else earliest
        except ValueError:
            pass
    return earliest


def _current_daily_cap(ws, cols, account: str) -> int:
    """Walks the account's own ramp schedule (workspace vs free_gmail — see
    _RAMP_SCHEDULES) using calendar days since THIS account's first recorded
    send — so each sending identity climbs its own ramp, and a long-running
    landscale.net doesn't let a day-1 gmail account borrow its cap, or vice
    versa."""
    ramp = _ramp_for(account)
    first = _first_send_date(ws, cols, account)
    if first is None:
        return ramp[0][1]   # no history yet: week-1 cap
    day_number = (datetime.now().date() - first).days + 1   # first day counts as day 1
    for cutoff, cap in ramp:
        if cutoff is None or day_number <= cutoff:
            return cap
    return ramp[-1][1]


def _send_gate(ws, cols, account: str):
    """Returns None if a send is allowed right now, or a Hungarian error
    string if that account's daily cap is hit. Read fresh from the Excel
    file every time (see _todays_send_times), not from an in-memory
    counter, so it holds even across an app restart."""
    cap = _current_daily_cap(ws, cols, account)
    times = _todays_send_times(ws, cols, account)
    if len(times) >= cap:
        label = gmail_send.ACCOUNTS.get(account, {}).get("label", account)
        return (f"Elérted a mai {cap} leveles limitet ezen a fiókon ({label}). "
                f"Ez új fióknál/domainnél szándékos — még nincs küldési "
                f"előélete, és a Gmail/más szolgáltatók gyanakvóbbak egy "
                f"hirtelen nagy mennyiségre, mint amennyit a fiók technikai "
                f"limitje megengedne. A limit magától nő, ahogy telnek a "
                f"napok. Válassz másik fiókot, vagy folytasd holnap.")
    return None


# ─── Prefetch ─────────────────────────────────────────────────────────────
# The app used to generate a lead only when you clicked "Következő lead", so
# every click cost the full pipeline: fetch the site, launch a browser,
# screenshot it, then a Groq vision call that has been measured anywhere from
# 1.5s to 66s. Twenty leads meant twenty of those waits back to back, which is
# how sending twenty emails turned into half an hour of sitting there.
#
# Nothing about that work needs to happen while you wait. Reviewing and
# sending one email takes a person 20-60 seconds, which is roughly one
# pipeline run — so generating the NEXT few leads in the background while you
# read the current one hides the latency almost entirely. First click of a
# session still pays for one run; after that it should be instant.
#
# Two workers, not five: the batch script needs five to saturate Groq's TPM
# budget across 95 leads, but here the human is the rate limiter and more
# workers would just burn quota generating leads that get skipped.
_PREFETCH_WORKERS = 2
_prefetch = {
    "cache": {},          # row -> processed lead dict
    "inflight": set(),    # rows currently being generated
    "lock": _threading.Lock(),
    "pool": None,
}


def _prefetch_pool():
    if _prefetch["pool"] is None:
        _prefetch["pool"] = _ThreadPoolExecutor(max_workers=_PREFETCH_WORKERS,
                                                 thread_name_prefix="prefetch")
    return _prefetch["pool"]


def _generate_lead(lead: dict) -> dict:
    """One full pipeline run + the browser cleanup that the crash fix requires.

    Every thread that calls into hu_email_gen gets its own Playwright browser
    via threading.local(), so every thread must also close it — otherwise the
    prefetch pool leaks a Chromium per worker exactly like the request threads
    used to."""
    try:
        if _state.get("mode") == "mech":
            processed = mech.process_lead(lead)
        elif _state.get("mode") == "ajanlat":
            processed = ajanlat.process_lead(lead)
        else:
            processed = heg.process_lead(lead)
    finally:
        try:
            heg._close_browser()
        except Exception as e:
            print(f"    (prefetch browser cleanup failed: {str(e)[:80]})")
    processed.pop("hero_image", None)
    return processed


def _prefetch_worker(lead: dict):
    row = lead["row"]
    try:
        processed = _generate_lead(lead)
        with _prefetch["lock"]:
            _prefetch["cache"][row] = processed
    except Exception as e:
        print(f"    (prefetch failed for row {row}: {str(e)[:110]})")
    finally:
        with _prefetch["lock"]:
            _prefetch["inflight"].discard(row)


def _kick_prefetch(ws, cols, ahead=3):
    """Queue up the next few unhandled leads that aren't cached or running."""
    try:
        for lead in _upcoming_leads(ws, cols, limit=ahead + 4):
            row = lead["row"]
            # The lead currently open for review is still "unhandled" in the
            # sheet until it's sent or skipped, so without this it gets picked
            # up and generated a SECOND time in the background while the user
            # reads the first copy — pure wasted Groq quota.
            if row == _state.get("current_row"):
                continue
            # Same cheap checks the route runs. Generating a lead that will be
            # blocked on arrival is exactly the wasted minute this whole
            # prefetch layer exists to remove.
            if _precheck(lead["email"]):
                continue
            with _prefetch["lock"]:
                if row in _prefetch["cache"] or row in _prefetch["inflight"]:
                    continue
                if len(_prefetch["inflight"]) >= ahead:
                    return
                _prefetch["inflight"].add(row)
            _prefetch_pool().submit(_prefetch_worker, lead)
    except Exception as e:
        print(f"    (prefetch scheduling failed: {str(e)[:110]})")


def _take_prefetched(row: int):
    with _prefetch["lock"]:
        return _prefetch["cache"].pop(row, None)


def _clear_prefetch():
    """Called when a different workbook is loaded — cached rows are keyed by
    row number, which means nothing once the sheet changes."""
    with _prefetch["lock"]:
        _prefetch["cache"].clear()
        _prefetch["inflight"].clear()


# ─── Cheap pre-checks, run BEFORE the expensive pipeline ──────────────────
# Ordering bug this fixes: the address checks used to run on the RESULT of
# generation, so a lead that was never sendable still cost a site fetch, a
# screenshot and a vision call first. These need nothing but the address
# string, so they belong in front of all of it.
#
# t-online.hu is here on evidence, not suspicion: five t-online-family
# addresses went out in one day and five bounces came back. Magyar Telekom
# wound the mailbox service down, but the domain still publishes MX, so
# nothing at the DNS layer catches it — the send has to be skipped by name.
_DEAD_PROVIDERS = {
    "t-online.hu": "A t-online.hu postafiók-szolgáltatás megszűnt. Ezekre a "
                   "címekre küldött levelek visszapattannak, és a "
                   "visszapattanás rontja a domain hírnevét. Hagyd ki.",
}


def _precheck(email: str) -> str:
    """Blocking-level reason to skip this lead without generating anything,
    or '' to proceed."""
    addr = (email or "").strip().lower()
    if "@" not in addr:
        return "Nincs érvényes e-mail cím ebben a sorban."
    domain = addr.split("@", 1)[1]
    for dead, msg in _DEAD_PROVIDERS.items():
        if domain == dead or domain.endswith("." + dead):
            return msg
    return ""


# ─── Follow-up queue ──────────────────────────────────────────────────────

def _row_thread(ws, cols, row) -> tuple:
    """(threadId, Message-ID) from the 'Szál' cell, or ('','') for a lead
    first-touched before this column existed."""
    raw = str(ws.cell(row=row, column=cols["thread"] + 1).value or "").strip()
    if not raw:
        return "", ""
    parts = raw.split("|", 1)
    return parts[0].strip(), (parts[1].strip() if len(parts) > 1 else "")


def _steps_done(ws, cols, row) -> int:
    """Highest follow-up step recorded for this lead. Highest, not count, so
    a hand-edited or partially-written cell can never rewind the sequence and
    resend a step the lead already got."""
    cell = str(ws.cell(row=row, column=cols["fups"] + 1).value or "")
    steps = [int(m.group(1)) for m in _FUP_TS_RE.finditer(cell)]
    return max(steps) if steps else 0


def _last_followup_date(ws, cols, row):
    """Date of the most recent follow-up, or None. Feeds due_step's second
    clock, which is what stops a stale lead getting two steps in one day."""
    cell = str(ws.cell(row=row, column=cols["fups"] + 1).value or "")
    dates = []
    for m in _FUP_TS_RE.finditer(cell):
        try:
            dates.append(datetime.strptime(m.group(2), "%Y-%m-%d").date())
        except ValueError:
            pass
    return max(dates) if dates else None


def _has_replied(ws, cols, row) -> bool:
    return bool(str(ws.cell(row=row, column=cols["replied"] + 1).value or "").strip())


def _has_bounced(ws, cols, row) -> bool:
    bi = cols.get("bounced")
    if bi is None:
        return False
    return bool(str(ws.cell(row=row, column=bi + 1).value or "").strip())


def _mark_bounced(ws, cols, row, when=None):
    when = when or datetime.now()
    ws.cell(row=row, column=cols["bounced"] + 1,
            value=f"Nem kézbesíthető {when:%Y-%m-%d}")


def _delivery_stats(ws, cols) -> dict:
    """Sent / bounced / replied totals, plus a bounce-rate verdict.

    Bounce rate is the one campaign metric with a hard industry line rather
    than a soft benchmark: under 2% is normal, over 3% is the accepted
    "stop and clean the list" threshold, and over 5% is where mailbox
    providers start treating the sender itself as the problem — repeatedly
    mailing addresses that don't exist is exactly the pattern a spammer
    working from a scraped list produces, so it is read as evidence about
    the sender, not about the list."""
    sent = bounced = replied = 0
    for r in ws.iter_rows(min_row=2):
        row = r[0].row
        if "@" not in str(r[cols["email"]].value or ""):
            continue
        if not _SEND_TS_RE.match(str(r[cols["sent"]].value or "").strip()):
            continue
        sent += 1
        if _has_bounced(ws, cols, row):
            bounced += 1
        if _has_replied(ws, cols, row):
            replied += 1

    rate = (bounced / sent * 100) if sent else 0.0
    if not sent or rate < 2:
        level, msg = "ok", ""
    elif rate < 5:
        level = "warn"
        msg = (f"{rate:.0f}%-os visszapattanási arány ({bounced}/{sent}). "
               f"3% fölött már érdemes ellenőrizni a listát.")
    else:
        level = "bad"
        msg = (f"{rate:.0f}%-os visszapattanási arány ({bounced}/{sent}) — ez "
               f"nagyon magas. 5% fölött a szolgáltatók már a FELADÓT tekintik "
               f"gyanúsnak, nem a listát: nem létező címekre küldeni pont az a "
               f"minta, amit egy lekapart listáról dolgozó spammer csinál. "
               f"Ellenőriztesd a címeket, mielőtt tovább küldesz.")
    return {"sent": sent, "bounced": bounced, "replied": replied,
            "bounce_rate": round(rate, 1), "level": level, "message": msg}


def _followup_queue(ws, cols) -> list:
    """Every lead whose next follow-up is due today or overdue.

    Four things exclude a lead, in order: no first-touch timestamp (nothing
    to count from), a recorded reply, the sequence already finished, or the
    next step's day not having arrived yet. Skipped leads never enter at all
    since they carry no 'Sent' timestamp."""
    out = []
    for r in ws.iter_rows(min_row=2):
        row = r[0].row
        email = str(r[cols["email"]].value or "")
        if "@" not in email:
            continue
        m = _SEND_TS_RE.match(str(r[cols["sent"]].value or "").strip())
        if not m:
            continue
        # A bounced address cannot receive a follow-up either — chasing a
        # mailbox that does not exist just generates more bounces, which is
        # the exact signal already hurting the domain.
        if _has_replied(ws, cols, row) or _has_bounced(ws, cols, row):
            continue
        # Each mode writes its own follow-up wording, so it only queues its
        # own trade's rows.
        if not _row_in_mode(r, cols):
            continue
        try:
            first_sent = datetime.strptime(m.group(1), "%Y-%m-%d").date()
        except ValueError:
            continue
        done = _steps_done(ws, cols, row)
        step = hu_followup.due_step(first_sent, done,
                                    last_followup=_last_followup_date(ws, cols, row))
        if not step:
            continue
        # Ajánlat mode sends two follow-ups, not the shared schedule's three.
        if step > _max_step():
            continue
        thread_id, msg_id = _row_thread(ws, cols, row)
        out.append({
            "row": row,
            "email": email,
            "name": str(r[cols["name"]].value or "") if "name" in cols else "",
            "first_sent": first_sent.isoformat(),
            "days_since": (datetime.now().date() - first_sent).days,
            "step": step,
            "step_label": hu_followup.step_label(step),
            "account": _marker_account(m),
            "has_thread": bool(thread_id),
        })
    # Oldest first: a lead 14 days out has a colder thread than one at 3 days,
    # so it is the one that loses most by waiting another day.
    out.sort(key=lambda d: (-d["days_since"], d["row"]))
    return out


def _mark_replied(ws, cols, row, when=None):
    when = when or datetime.now()
    ws.cell(row=row, column=cols["replied"] + 1,
            value=f"Igen {when:%Y-%m-%d}")


# ─── Routes ──────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("email_hu.html")


@app.route("/load", methods=["POST"])
def load():
    try:
        data = request.get_json() or {}
        path = (data.get("path") or "").strip().strip('"')
        if not path:
            # "leads_template.xlsx" (the 1-row example file _make_template()
            # writes) matches "lead" too, so it must be excluded by name here
            # — a bare blank submit was silently loading that instead of the
            # real list, producing exactly one fake lead ("Példa Tetőfedő
            # Kft.") and then a premature "Nincs több lead" right after.
            here = _Path(__file__).parent
            for f in here.glob("*.xlsx"):
                name = f.name.lower()
                if "template" in name:
                    continue
                if any(k in name for k in ("lead", "email", "hun")):
                    path = str(f)
                    break
        if not path or not os.path.isfile(path):
            return jsonify({"error": f"Excel fájl nem található: {path}"}), 400

        wb, ws, cols = _load_excel(path)
        if "email" not in cols:
            return jsonify({"error": "Nem találok Email oszlopot ebben a fájlban."}), 400

        account = _resolve_account(data)
        mode = data.get("mode") if data.get("mode") in ("mech", "ajanlat") else "builders"
        _state.update({"excel_path": path, "wb": wb, "ws": ws, "cols": cols,
                       "current_row": None, "mode": mode})
        # Row-keyed cache is meaningless against a different sheet.
        _clear_prefetch()
        # Start generating immediately, so the first "Következő lead" click
        # lands on an already-finished lead instead of starting the wait.
        _kick_prefetch(ws, cols)
        total, sent = _counts(ws, cols)
        sent_today = len(_todays_send_times(ws, cols, account))
        return jsonify({"ok": True, "total": total, "sent": sent, "sent_today": sent_today,
                         "daily_cap": _current_daily_cap(ws, cols, account), "account": account,
                         "mode": mode,
                         "accounts": [{"key": k, "label": v["label"]} for k, v in gmail_send.ACCOUNTS.items()],
                         "path": path})
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/next", methods=["POST"])
def next_lead():
    try:
        ws, cols = _state["ws"], _state["cols"]
        if ws is None:
            return jsonify({"error": "Előbb tölts be egy Excel fájlt."}), 400

        # A lead is only "used up" by an explicit Send or Skip (see below) —
        # clicking Next again on a lead you haven't decided on yet would
        # otherwise silently abandon it with no record either way.
        if _state["current_row"] is not None:
            return jsonify({"error": "Előbb küldd el vagy hagyd ki az aktuális levelet."}), 400

        lead = _next_lead(ws, cols)
        if not lead:
            total, sent = _counts(ws, cols)
            return jsonify({"error": "Nincs több lead — kész!", "total": total, "sent": sent}), 200

        # CHEAP CHECKS FIRST. Anything decidable from the address string alone
        # runs before a site fetch, a browser launch and a vision call — the
        # old order spent ~60s of work on leads that were never sendable.
        blocked = _precheck(lead["email"])
        if blocked:
            _state["current_row"] = lead["row"]
            _state["current_lead"] = None
            _kick_prefetch(ws, cols)
            total, sent = _counts(ws, cols)
            return jsonify({
                "lead": {"email": lead["email"], "name": lead.get("name", ""),
                          "city": lead.get("city", ""), "niche": lead.get("niche", ""),
                          "subject": "", "body": ""},
                "blocked": blocked, "total": total, "sent": sent,
                "sent_today": len(_todays_send_times(ws, cols, _state["account"])),
                "daily_cap": _current_daily_cap(ws, cols, _state["account"]),
            })

        _state["current_row"] = lead["row"]

        # Already generated in the background while the previous lead was
        # being reviewed? Then this returns instantly. See the prefetch block
        # above for why this exists.
        processed = _take_prefetched(lead["row"])
        if processed is None:
            processed = _generate_lead(lead)

        # Top the queue back up for the leads after this one.
        _kick_prefetch(ws, cols)

        # Mechanics mode found no real detail on the site: nothing honest to
        # write, so same shape as a precheck block (skip only).
        if processed.get("blocked"):
            _state["current_lead"] = None
            total, sent = _counts(ws, cols)
            account = _state["account"]
            return jsonify({
                "lead": processed, "blocked": processed["blocked"],
                "total": total, "sent": sent,
                "sent_today": len(_todays_send_times(ws, cols, account)),
                "daily_cap": _current_daily_cap(ws, cols, account),
            })

        _state["current_lead"] = {
            "email": processed["email"], "subject": processed["subject"], "body": processed["body"],
        }
        _state["current_hero"] = None

        account = _state["account"]
        total, sent = _counts(ws, cols)
        return jsonify({
            "lead": processed,
            "address_warning": _address_risk(processed["email"]),
            "email_on_site": processed.get("email_on_site", "unknown"),
            "total": total, "sent": sent, "sent_today": len(_todays_send_times(ws, cols, account)),
            "daily_cap": _current_daily_cap(ws, cols, account),
        })
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500


def _check_replies_sweep(ws, cols) -> dict:
    """Ask Gmail which of the open threads now contain an inbound message,
    and record those in the 'Válaszolt?' column. Shared by the automatic
    sweep (/followups, run on every load — see below) and the manual
    "Válaszok ellenőrzése" button (/check_replies), so there is exactly one
    place this logic lives.

    Only walks leads that are still open (sent, no recorded reply, sequence
    unfinished) — re-checking a closed lead costs an API call and can change
    nothing."""
    checked = found = skipped = bounces = 0
    errors = []
    for r in ws.iter_rows(min_row=2):
        row = r[0].row
        if "@" not in str(r[cols["email"]].value or ""):
            continue
        m = _SEND_TS_RE.match(str(r[cols["sent"]].value or "").strip())
        if not m or _has_replied(ws, cols, row) or _has_bounced(ws, cols, row):
            continue
        if _steps_done(ws, cols, row) >= _max_step():
            continue
        thread_id, _ = _row_thread(ws, cols, row)
        if not thread_id:
            skipped += 1          # first-touched before thread ids existed
            continue
        try:
            checked += 1
            status = gmail_send.thread_status(thread_id, _marker_account(m))
            if status == "reply":
                _mark_replied(ws, cols, row)
                found += 1
            elif status == "bounce":
                _mark_bounced(ws, cols, row)
                bounces += 1
        except Exception as e:
            errors.append(f"sor {row}: {str(e)[:80]}")
    if found or bounces:
        _save_wb()
    return {"checked": checked, "found": found, "bounces": bounces,
            "skipped_no_thread": skipped, "errors": errors[:5]}


@app.route("/followups", methods=["POST"])
def followups():
    """The due-today follow-up queue.

    Runs the live Gmail reply sweep FIRST, automatically, every single time
    this loads — not something that depends on remembering to click
    "Válaszok ellenőrzése" or on manually noting down who replied. You never
    write down a reply, so the app has to go find out for itself before it
    ever shows you a queue; a queue built straight from the sheet's own
    (possibly stale) 'Válaszolt?' column could list someone who already
    answered an hour ago.

    The manual button/route still exists for an explicit re-check with its
    own feedback message, but nothing about correctness depends on it being
    clicked."""
    try:
        ws, cols = _state["ws"], _state["cols"]
        if ws is None:
            return jsonify({"error": "Előbb tölts be egy Excel fájlt."}), 400
        # The sweep is skippable for exactly one caller: the manual
        # "Válaszok ellenőrzése" button, which has just run this same sweep
        # itself via /check_replies and only needs the queue rebuilt. Without
        # this it swept twice per click, doubling Gmail traffic for nothing.
        # Every other caller leaves it on, so the queue is never built from a
        # stale 'Válaszolt?' column.
        do_sweep = (request.get_json(silent=True) or {}).get("sweep", True)
        sweep = _check_replies_sweep(ws, cols) if do_sweep else None
        queue = _followup_queue(ws, cols)
        needs_auth = sorted(a for a in gmail_send.ACCOUNTS
                            if gmail_send.missing_scopes(a))
        return jsonify({
            "ok": True, "queue": queue, "count": len(queue),
            "needs_reauth": needs_auth,
            "delivery": _delivery_stats(ws, cols),
            "schedule": [{"step": s, "days": d, "label": l}
                         for s, d, l in hu_followup.SCHEDULE],
            "auto_check": sweep,
        })
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/check_replies", methods=["POST"])
def check_replies():
    """Manual re-run of the same sweep /followups already runs automatically
    on every load. Kept as an explicit action for its own status message
    ("3 ellenőrizve, 1 új válasz...") when you want to force a re-check
    without reloading the whole queue."""
    try:
        ws, cols = _state["ws"], _state["cols"]
        if ws is None:
            return jsonify({"error": "Előbb tölts be egy Excel fájlt."}), 400
        sweep = _check_replies_sweep(ws, cols)
        account = _state["account"]
        return jsonify({"ok": True, **sweep,
                         "delivery": _delivery_stats(ws, cols),
                         "sent_today": len(_todays_send_times(ws, cols, account)),
                         "daily_cap": _current_daily_cap(ws, cols, account)})
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/followup_preview", methods=["POST"])
def followup_preview():
    """Draft for one queued follow-up, for review before it goes out. Same
    human-in-the-loop shape as the first touch: nothing sends without a
    second, explicit click."""
    try:
        ws, cols = _state["ws"], _state["cols"]
        if ws is None:
            return jsonify({"error": "Előbb tölts be egy Excel fájlt."}), 400
        data = request.get_json() or {}
        row = int(data.get("row") or 0)
        if not row:
            return jsonify({"error": "Hiányzó sor."}), 400
        if _has_replied(ws, cols, row):
            return jsonify({"error": "Erre a leadre már érkezett válasz."}), 400

        m = _SEND_TS_RE.match(str(ws.cell(row=row, column=cols["sent"] + 1).value or "").strip())
        if not m:
            return jsonify({"error": "Ez a lead még nem kapott első levelet."}), 400
        account = _marker_account(m)
        step = _steps_done(ws, cols, row) + 1
        if step > _max_step():
            return jsonify({"error": "A sorozat már befejeződött ennél a leadnél."}), 400

        thread_id, msg_id = _row_thread(ws, cols, row)

        # Same live Gmail check as send_followup — catches a reply that
        # landed after the last "Válaszok ellenőrzése" click, so the draft
        # never even opens for a lead that has since answered.
        #
        # FAILS CLOSED: if the Gmail call itself errors (expired token,
        # missing scope, network blip, rate limit), that is NOT the same as
        # "no reply" — it means the reply status is simply unknown, and
        # unknown must block the draft rather than wave it through. Treating
        # a failed check as "clear to proceed" would silently defeat the
        # entire point of checking live in the first place.
        if thread_id:
            try:
                status = gmail_send.thread_status(thread_id, account)
            except Exception as e:
                return jsonify({"error": "Nem sikerült ellenőrizni a Gmail szálat "
                                          "(lehet lejárt a bejelentkezés), ezért "
                                          "biztonságból nem mutatom a piszkozatot: "
                                          f"{str(e)[:100]}"}), 400
            if status == "reply":
                _mark_replied(ws, cols, row)
                _save_wb()
                return jsonify({"error": "Időközben válasz érkezett erre a levélre."}), 400
            if status == "bounce":
                _mark_bounced(ws, cols, row)
                _save_wb()
                return jsonify({"error": "Ez a cím időközben visszapattant."}), 400

        name  = str(ws.cell(row=row, column=cols["name"] + 1).value or "") if "name" in cols else ""
        first = heg._first_name_for({"owner": "", "name": name})

        # Reuse the ORIGINAL subject so the thread holds together. If the
        # thread id predates this feature there is nothing to read it from,
        # and the caller is warned rather than silently starting a new thread.
        subject, warn = "", ""
        if thread_id:
            try:
                subject = hu_followup.followup_subject(
                    gmail_send.thread_subject(thread_id, account))
            except Exception as e:
                warn = f"Nem tudom kiolvasni az eredeti tárgyat: {str(e)[:70]}"
        if not subject:
            if _state.get("mode") == "mech":
                subject = "Re: az árajánlatokról"
            elif _state.get("mode") == "ajanlat":
                pos = (str(ws.cell(row=row, column=cols["position"] + 1).value or "")
                       if "position" in cols else "")
                subject = ("Re: " + ajanlat.ajanlat_subject(pos)
                           if ajanlat.clean_position(pos) else "Re: a pozícióról")
            else:
                subject = "Re: az árkérésekről"
            warn = warn or ("Ehhez a leadhez nincs eltárolt szál (a levél még a "
                            "követés bevezetése előtt ment ki), így ez új szálként "
                            "érkezhet meg, nem válaszként.")

        return jsonify({
            "ok": True, "row": row, "step": step,
            "step_label": hu_followup.step_label(step),
            "email": str(ws.cell(row=row, column=cols["email"] + 1).value or ""),
            "name": name, "account": account,
            "subject": subject,
            "body": (mech.followup_body(step, first) if _state.get("mode") == "mech"
                     else ajanlat.followup_body(
                         step, first,
                         str(ws.cell(row=row, column=cols["niche"] + 1).value or "")
                         if "niche" in cols else "")
                     if _state.get("mode") == "ajanlat"
                     else hu_followup.step_body(step, first)),
            "warning": warn,
        })
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/send_followup", methods=["POST"])
def send_followup():
    """Send one reviewed follow-up into its existing thread."""
    try:
        ws, cols = _state["ws"], _state["cols"]
        if ws is None:
            return jsonify({"error": "Előbb tölts be egy Excel fájlt."}), 400
        data = request.get_json() or {}
        row = int(data.get("row") or 0)
        if not row:
            return jsonify({"error": "Hiányzó sor."}), 400

        # Re-checked at send time, not just at preview time: the queue may
        # have been built minutes ago and a reply can land in between.
        if _has_replied(ws, cols, row):
            return jsonify({"error": "Erre a leadre már érkezett válasz — nem küldök utánkövetést."}), 400

        m = _SEND_TS_RE.match(str(ws.cell(row=row, column=cols["sent"] + 1).value or "").strip())
        if not m:
            return jsonify({"error": "Ez a lead még nem kapott első levelet."}), 400
        account = _marker_account(m)          # same sender as the first touch
        step = _steps_done(ws, cols, row) + 1
        if step > _max_step():
            return jsonify({"error": "A sorozat már befejeződött ennél a leadnél."}), 400

        # Same-day guard. The queue is rebuilt from the sheet, but a stale
        # browser tab, a double click, or a retry after a slow response can
        # POST the same row twice. By then _steps_done has already counted the
        # first marker, so the second call is NOT a harmless duplicate of F1:
        # it sends F2 minutes after F1, to someone who has heard nothing back.
        # No step of the schedule ever falls on the same day as another, so a
        # follow-up already recorded today means this is a repeat.
        last_fu = _last_followup_date(ws, cols, row)
        if last_fu and last_fu == datetime.now().date():
            return jsonify({"error": "Erre a leadre ma már ment utánkövetés, "
                                      "nem küldök másodikat ugyanaznap."}), 400

        blocked = _send_gate(ws, cols, account)
        if blocked:
            return jsonify({"error": blocked}), 429

        thread_id, msg_id = _row_thread(ws, cols, row)

        # Live check against Gmail itself, not just the sheet's "Válaszolt?"
        # cell. That cell only updates when someone clicks "Válaszok
        # ellenőrzése" — skip that click (or a reply lands in the gap between
        # a click and this send) and the sheet still says "no reply" while
        # Gmail already has one. Sending a follow-up to someone who just
        # replied is the one mistake this whole queue exists to prevent, so
        # it can't depend on a separate manual step being remembered.
        #
        # FAILS CLOSED: an error from the Gmail call (expired token, missing
        # scope, network blip, rate limit) means "unknown", not "no reply".
        # The original version caught the exception and fell through to
        # status="" — i.e. treated a FAILED safety check as a PASSED one and
        # sent anyway. That's the one bug that would make this whole feature
        # worthless, so a failed check now blocks the send instead of
        # silently clearing it.
        if thread_id:
            try:
                status = gmail_send.thread_status(thread_id, account)
            except Exception as e:
                return jsonify({"error": "Nem sikerült ellenőrizni a Gmail szálat "
                                          "(lehet lejárt a bejelentkezés) — "
                                          "biztonságból nem küldök utánkövetést, amíg "
                                          f"ez nem működik: {str(e)[:100]}"}), 400
            if status == "reply":
                _mark_replied(ws, cols, row)
                _save_wb()
                return jsonify({"error": "Időközben válasz érkezett erre a levélre — "
                                          "nem küldök utánkövetést."}), 400
            if status == "bounce":
                _mark_bounced(ws, cols, row)
                _save_wb()
                return jsonify({"error": "Ez a cím időközben visszapattant — "
                                          "nem küldök utánkövetést."}), 400

        to      = (data.get("email") or str(ws.cell(row=row, column=cols["email"] + 1).value or "")).strip()
        subject = (data.get("subject") or "").strip() or "Re:"
        body    = data.get("body") or ""
        if not body.strip():
            return jsonify({"error": "Üres levél."}), 400

        gmail_send.send_email(to=to, subject=subject, plain_body=body,
                               account=account, thread_id=thread_id,
                               in_reply_to=msg_id)
        _log_send(account)

        cell = str(ws.cell(row=row, column=cols["fups"] + 1).value or "").strip()
        marker = f"F{step} {datetime.now():%Y-%m-%d %H:%M} [{account}]"
        ws.cell(row=row, column=cols["fups"] + 1,
                value=f"{cell}; {marker}" if cell else marker)
        _save_wb()

        return jsonify({"ok": True, "row": row, "step": step, "account": account,
                         "sent_today": len(_todays_send_times(ws, cols, account)),
                         "daily_cap": _current_daily_cap(ws, cols, account),
                         "remaining": len(_followup_queue(ws, cols))})
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/mark_replied", methods=["POST"])
def mark_replied():
    """Manual override — mark a lead as replied without an API round-trip.

    Needed for the leads first-touched before thread ids were recorded (their
    replies can't be detected automatically), and as an escape hatch whenever
    a reply arrives somewhere the API won't see it, e.g. they phoned instead."""
    try:
        ws, cols = _state["ws"], _state["cols"]
        if ws is None:
            return jsonify({"error": "Előbb tölts be egy Excel fájlt."}), 400
        row = int((request.get_json() or {}).get("row") or 0)
        if not row:
            return jsonify({"error": "Hiányzó sor."}), 400
        _mark_replied(ws, cols, row)
        _save_wb()
        return jsonify({"ok": True, "row": row,
                         "remaining": len(_followup_queue(ws, cols))})
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/account_stats", methods=["POST"])
def account_stats():
    """Lets the UI refresh the 'Ma: X/cap' figure the instant the account
    dropdown changes, without waiting for the next Next/Send round-trip."""
    try:
        ws, cols = _state["ws"], _state["cols"]
        if ws is None:
            return jsonify({"error": "Előbb tölts be egy Excel fájlt."}), 400
        account = _resolve_account(request.get_json() or {})
        return jsonify({"ok": True, "account": account,
                         "sent_today": len(_todays_send_times(ws, cols, account)),
                         "daily_cap": _current_daily_cap(ws, cols, account)})
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/send", methods=["POST"])
def send():
    """One real Gmail API send — only ever called from the Küldés button
    click in the UI, one lead at a time. Subject/body can be edited in the
    browser first; whatever the client posts here is what actually goes
    out, not necessarily what /next originally generated."""
    try:
        row = _state["current_row"]
        if row is None or _state["current_lead"] is None:
            return jsonify({"error": "Nincs elküldésre váró levél."}), 400

        ws, cols = _state["ws"], _state["cols"]
        data = request.get_json() or {}
        account = _resolve_account(data)

        blocked = _send_gate(ws, cols, account)
        if blocked:
            return jsonify({"error": blocked}), 429

        to      = (data.get("email") or _state["current_lead"]["email"]).strip()
        subject = (data.get("subject") or _state["current_lead"]["subject"]).strip()
        body    = (data.get("body") or _state["current_lead"]["body"])

        res = gmail_send.send_email(
            to=to, subject=subject, plain_body=body,
            hero_image=_state["current_hero"], account=account)
        msg_id = res["id"]
        _log_send(account)

        # Timestamped + account-tagged marker so the daily cap can tell
        # landscale's history apart from gmail's — see _todays_send_times /
        # _send_gate / _marker_account.
        _mark(row, f"Sent {datetime.now():%Y-%m-%d %H:%M} [{account}]")
        # Thread + Message-ID, so the follow-up engine can later ask whether
        # they replied and can post its follow-ups INTO this conversation
        # rather than starting a new one. Stored on the same row, written
        # before _save_wb() runs in _mark above... which has already run, so
        # this needs its own save.
        ws.cell(row=row, column=cols["thread"] + 1,
                value=f"{res.get('thread_id','')}|{res.get('message_id','')}")
        _save_wb()
        _state["current_row"] = None
        _state["current_lead"] = None
        _state["current_hero"] = None

        total, sent = _counts(ws, cols)
        sent_today = len(_todays_send_times(ws, cols, account))
        return jsonify({"ok": True, "message_id": msg_id, "account": account,
                         "total": total, "sent": sent, "sent_today": sent_today,
                         "daily_cap": _current_daily_cap(ws, cols, account)})
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/skip", methods=["POST"])
def skip():
    try:
        ws, cols = _state["ws"], _state["cols"]
        row = _state["current_row"]
        if ws is None:
            return jsonify({"error": "Nincs betöltött fájl."}), 400
        if row is not None:
            _mark(row, "Skip")
        _state["current_row"] = None
        _state["current_lead"] = None
        _state["current_hero"] = None
        total, sent = _counts(ws, cols)
        return jsonify({"ok": True, "total": total, "sent": sent})
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    if not dm_app.API_KEY:
        print("GROQ_API_KEY is not set — check the .env file in this folder.")
    print(f"Hungarian email sender running on http://localhost:{PORT}")
    app.run(port=PORT, debug=False)

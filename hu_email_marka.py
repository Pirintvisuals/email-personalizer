"""
Márkaszerviz mode for the HU email sender (hu_email_app.py).

Leads are small independent workshops that specialise in one make (or one
group: VW-csoport, BMW+MINI). Input is the lead finder's
leads-markaszerviz-<date>.xlsx. Only rows with Confidence high/medium and a
published email are used.

THE OPENER is the only personalised line. It lives in the sheet's
"Nyitómondat" column, written once per file from the row's Proof quote and
Personal hook, and read by a human in Excel before anything goes out (same
idea as ajánlat mode: no AI writes copy at send time). A row without one
falls back to "Láttam, hogy Önök {make} autókra szakosodtak."

THE BODY (user's direction, 2026-11-04): greeting, personal line, "I've
built you a quoting program", one line on what it does, one line on why
them, then ask permission to send it. Short, no over-explaining. Never call
it free (it isn't) and no pilot/test-period framing. Do not edit the body
without asking.

SENDING RULES specific to this mode (enforced in hu_email_app.py):
  * plain text, no image, no link
  * never a second first-email to any address: a global contacted list
    (marka_contacted.json) plus every address in the older HU lead files
  * opt-outs go on a permanent do-not-contact list (do_not_contact.json)
  * ONE follow-up, 5 working days after the first email, in the same thread
"""

import json
import os
import re
from datetime import date, timedelta
from pathlib import Path

import openpyxl

MAX_STEP = 1
FOLLOWUP_WORKDAYS = 5
STEP_LABEL = f"Emlékeztető (+{FOLLOWUP_WORKDAYS} munkanap)"

OK_CONFIDENCE = ("high", "medium")

_HERE = Path(__file__).parent
CONTACTED_PATH = _HERE / "marka_contacted.json"
DNC_PATH = _HERE / "do_not_contact.json"

# Older Hungarian lead files to check against. Any address that appears
# anywhere in them counts as contacted, even rows that were never sent:
# a few lost leads is cheaper than a second "first" email.
_OLD_FILE_DIRS = (
    Path(r"C:\Users\pirin\OneDrive\Asztali gép"),
    Path(r"C:\Users\pirin\Downloads"),
)
_OLD_FILE_RE = re.compile(r"hun|magyar|ajanlat|mech", re.IGNORECASE)
_EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+")


# ── Subject: short natural make, hyphenated the Hungarian way ─────────────
# Groups and pairs are checked before the single makes they contain.


def subject_prefix(make_cell: str) -> str:
    # "(+ Mercedes, Audi)" lists side makes; the subject names the main one.
    m = re.sub(r"\(\+[^)]*\)", "", (make_cell or "")).lower()
    has = lambda *w: any(x in m for x in w)
    if has("fiat", "alfa", "lancia"):
        return "olasz autós "
    if has("transporter", "commercial"):
        return "VW-haszonjármű-"
    if has("gte"):
        return "VW- és Tesla-"
    if has("peugeot", "citroën", "citroen"):
        return "francia autós "
    if has("mercedes") and has("bmw"):
        return "Mercedes- és BMW-"
    if has("mercedes"):
        return "Mercedes-"
    if has("kia") and has("hyundai"):
        return "Kia- és Hyundai-"
    if has("kia"):
        return "Kia-"
    if has("renault") and has("dacia"):
        return "Renault- és Dacia-"
    if has("toyota", "lexus"):
        return "Toyota- és Lexus-" if has("lexus") else "Toyota-"
    if has("škoda", "skoda") and not has("vw", "volkswagen", "audi", "seat", "cupra"):
        return "Škoda-"
    if has("vw", "volkswagen", "audi", "škoda", "skoda", "seat", "cupra"):
        return "VW-csoportos "
    for words, prefix in ((("bmw", "mini"), "BMW-"), (("opel",), "Opel-"),
                          (("ford",), "Ford-"), (("mazda",), "Mazda-"),
                          (("honda",), "Honda-")):
        if has(*words):
            return prefix
    first = re.split(r"[,(+/]", make_cell or "")[0].strip()
    return f"{first}-" if first else ""


def subject(make_cell: str) -> str:
    p = subject_prefix(make_cell)
    return f"Árajánlat-készítő {p}szervizeknek" if p else "Árajánlat-készítő autószervizeknek"


def make_label(make_cell: str) -> str:
    """Short make name for grouping in the summary ('BMW', 'VW-csoport')."""
    p = subject_prefix(make_cell).strip().rstrip("-").replace("- és ", "/")
    return {"VW-csoportos": "VW-csoport", "francia autós": "Francia",
            "olasz autós": "Olasz"}.get(p, p) or "?"


OPENERS_PATH = _HERE / "marka_openers.json"


def openers() -> dict:
    """address -> hand-checked opening line. Kept outside the leads sheet
    because the lead finder rewrites that file in place (2026-11-04: it
    added 24 rows and wiped the Nyitómondat column). The sheet's own
    Nyitómondat cell still wins when it has one."""
    try:
        return {k.lower(): v for k, v in
                json.loads(OPENERS_PATH.read_text(encoding="utf-8")).items()}
    except (FileNotFoundError, ValueError):
        return {}


def fallback_opener(make_cell: str) -> str:
    first = re.split(r"[,(+/]", make_cell or "")[0].strip() or "egy"
    return f"Láttam, hogy Önök {first} autókra szakosodtak."


# ── The email ─────────────────────────────────────────────────────────────

def body(opener: str) -> str:
    return "\n\n".join([
        "Jó napot kívánok!",
        opener,
        "Összeraktam Önöknek egy árajánlat-készítő programot. Átvizsgálás "
        "után a munkafelvevő beírja az alvázszámot és a munkát, a program "
        "pedig elkészíti a tételes ajánlatot. Kiküldés előtt mindig Önök "
        "nézik át.",
        "Azért Önöknek készítettem, mert egy márkára szakosodott szervizben "
        "működik a legjobban: amit egy motornál egyszer kiválasztottak, azt "
        "megjegyzi, így legközelebb alig kell beírni valamit.",
        "Elküldhetem?",
        "Ha nem érdekli, elég egy szó, és többé nem keresem.",
        "Üdvözlettel:\nPirint Milán",
    ])


def followup_body() -> str:
    return "\n\n".join([
        "Jó napot kívánok!",
        "Csak felhozom az előző levelem. Ha érdekli, elküldöm a programot, "
        "amit Önöknek összeraktam. Ha nem, egy szó elég.",
        "Üdvözlettel:\nPirint Milán",
    ])


# ── Follow-up timing: working days, not calendar days ─────────────────────

def workdays_between(start: date, end: date) -> int:
    """Mon-Fri days after `start`, up to and including `end`."""
    n, d = 0, start
    while d < end:
        d += timedelta(days=1)
        if d.weekday() < 5:
            n += 1
    return n


def due_step(first_sent: date, steps_done: int, today: date = None) -> int:
    if first_sent is None or steps_done >= MAX_STEP:
        return 0
    today = today or date.today()
    return 1 if workdays_between(first_sent, today) >= FOLLOWUP_WORKDAYS else 0


# ── Reply types and what happens next ─────────────────────────────────────

REPLY_TYPES = {
    "interested": ("Érdekli",   "Bemutató elküldése"),
    "not_now":    ("Most nem",  "Később újra, kézzel"),
    "no":         ("Nem",       "Lezárva"),
    "optout":     ("Leiratkozott", "Lezárva, tiltólistán"),
}


# ── Contacted / do-not-contact lists ──────────────────────────────────────

def _load_set(path: Path) -> set:
    try:
        return {e.lower() for e in json.loads(path.read_text(encoding="utf-8"))}
    except (FileNotFoundError, ValueError):
        return set()


def _save_set(path: Path, s: set) -> None:
    path.write_text(json.dumps(sorted(s), ensure_ascii=False, indent=1),
                    encoding="utf-8")


def dnc() -> set:
    return _load_set(DNC_PATH)


def add_dnc(email: str) -> None:
    s = dnc()
    s.add(email.strip().lower())
    _save_set(DNC_PATH, s)


def contacted() -> dict:
    """address -> {"sent": "Sent 2026-11-04 11:36 [landscale]",
                   "thread": "<threadId>|<Message-ID>"}.

    This file, not the sheet, is the record of who got a first email. The
    sheet's markers can be lost (on 2026-11-04 all 29 were overwritten,
    most likely by the sheet being saved from Excel after the app wrote it),
    so on every load the app writes these back into the sheet (see
    hu_email_app._reconcile_marka). Old files were a plain list of
    addresses; those load as entries with no send data."""
    try:
        raw = json.loads(CONTACTED_PATH.read_text(encoding="utf-8"))
    except (FileNotFoundError, ValueError):
        return {}
    if isinstance(raw, list):
        return {e.lower(): {} for e in raw}
    return {k.lower(): v for k, v in raw.items()}


def _write_contacted(d: dict) -> None:
    CONTACTED_PATH.write_text(json.dumps(d, ensure_ascii=False, indent=1),
                              encoding="utf-8")


def add_contacted(email: str, sent: str = "", thread: str = "") -> None:
    d = contacted()
    d[email.strip().lower()] = {"sent": sent, "thread": thread}
    _write_contacted(d)


def update_contacted(email: str, **fields) -> None:
    d = contacted()
    d.setdefault(email.strip().lower(), {}).update(fields)
    _write_contacted(d)


def old_file_addresses(exclude: str = "") -> dict:
    """address -> file name, for every address in the older HU lead files.
    The márkaszerviz files themselves are skipped (their sends are in
    marka_contacted.json), so a sheet never blocks its own rows."""
    out = {}
    ex = os.path.normcase(os.path.abspath(exclude)) if exclude else ""
    for d in _OLD_FILE_DIRS:
        if not d.is_dir():
            continue
        for f in d.glob("*.xlsx"):
            name = f.name
            if name.startswith("~$") or "markaszerviz" in name.lower():
                continue
            if not _OLD_FILE_RE.search(name):
                continue
            if ex and os.path.normcase(str(f.resolve())) == ex:
                continue
            try:
                wb = openpyxl.load_workbook(f, read_only=True)
            except Exception:
                continue
            try:
                for ws in wb:
                    for row in ws.iter_rows(values_only=True):
                        for v in row:
                            if isinstance(v, str) and "@" in v:
                                for m in _EMAIL_RE.findall(v):
                                    out.setdefault(m.lower(), name)
            finally:
                wb.close()
    return out


def contact_block(email: str, old: dict) -> str:
    """Why this address must not get a first email, or ''."""
    e = (email or "").strip().lower()
    if e in dnc():
        return "Tiltólistán van (leiratkozott). Soha többé nem keressük."
    if e in contacted():
        return "Ezt a címet már megkerestük egy korábbi márkaszervizes körben."
    if e in old:
        return f"Ez a cím szerepel egy régebbi magyar leadlistában ({old[e]}), nem írunk neki újra."
    return ""


# ── Per-lead pipeline (same shape as the other modes) ─────────────────────

def process_lead(lead: dict) -> dict:
    tag = lead.get("name") or lead.get("email") or f"row {lead['row']}"
    print(f"  [marka {lead['row']}] {tag}")
    lead["niche_resolved"] = lead.get("make", "")
    lead["hero_image"] = None
    lead["error"] = ""
    lead["email_on_site"] = "unknown"
    lead["blocked"] = lead.get("blocked", "")
    if lead["blocked"]:
        lead["subject"] = lead["body"] = lead["opener"] = ""
        return lead
    opener = ((lead.get("opener_custom") or "").strip()
              or openers().get((lead.get("email") or "").strip().lower(), "")
              or fallback_opener(lead.get("make", "")))
    lead["opener"] = opener
    lead["subject"] = subject(lead.get("make", ""))
    lead["body"] = body(opener)
    print(f"    -> {opener[:90]}")
    return lead

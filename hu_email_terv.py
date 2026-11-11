"""
Tervolvasó mode for the HU email sender (hu_email_app.py).

Leads: general contractors, turnkey / type-house builders, developer-
builders and cost-estimating offices, anyone who prices whole buildings
from plans. Input is any .xlsx with an email column; default path is
Leads\\Hungary\\leads-tervolvaso.xlsx.

THE OFFER (user, 2026-11-11): a 2-minute video of the program working on a
real family house's plans, offered right at the top of the email; the ask
is "Elküldhetem a videót?". Still says the program is in development
("amit fejlesztek"). The building in the how-it-works line is per lead
(sheet column "Épülettípus": "egy társasház", "egy sorház vagy ikerház"),
default "egy ház". Do not send until the video exists.
No numbers/percentages, no client references, "AI" not mentioned.

Everything else is Márka mode, unchanged: plain text, no link, no image;
never a second first email (own terv_contacted.json + the shared
do-not-contact list; older HU lead files are NOT checked, see
old_file_addresses); three follow-ups on
the shared +3 / +7 / +14 day schedule in the same thread; reply tagging.
The opener comes from the sheet's "Nyitómondat" column or
terv_openers.json, read by a human before sending, else a plain fallback.
"""

import json
import random
from pathlib import Path

import hu_email_marka as marka

OK_CONFIDENCE = marka.OK_CONFIDENCE

_HERE = Path(__file__).parent
CONTACTED_PATH = _HERE / "terv_contacted.json"
OPENERS_PATH = _HERE / "terv_openers.json"


def subject(_unused: str = "") -> str:
    return "Mennyiségkimutatás a tervekből"


def openers() -> dict:
    try:
        return {k.lower(): v for k, v in
                json.loads(OPENERS_PATH.read_text(encoding="utf-8")).items()}
    except (FileNotFoundError, ValueError):
        return {}


def fallback_opener(_niche: str = "") -> str:
    return "Láttam, hogy Önök tervek alapján építenek házakat."


# ── The email ─────────────────────────────────────────────────────────────

DEFAULT_BUILDING = "egy ház"


def body(opener: str, building: str = "") -> str:
    """building: what this lead builds, indefinite singular ("egy társasház",
    "egy sorház vagy ikerház"), from the sheet's Épülettípus column."""
    building = (building or "").strip() or DEFAULT_BUILDING
    return "\n\n".join([
        "Jó napot kívánok!",
        opener,
        "Készítettem egy 2 perces videót egy programról, amit fejlesztek: "
        "egy valódi családi ház tervein mutatja meg, hogyan működik.",
        f"A program {building} tervrajzaiból (alaprajz, metszetek, "
        "homlokzatok) elkészíti a mennyiségkimutatás első változatát. Minden "
        "szám mellé odaírja, melyik tervlapról jött, és külön jelzi, ami "
        "hiányzik vagy ellentmond. Így a kolléga nem nulláról mér fel, hanem "
        "ellenőriz.",
        "Elküldhetem a videót?",
        "Ha nem érdekli, elég egy szó, és többé nem keresem.",
        "Üdvözlettel:\nPirint Milán",
    ])


# ── Follow-ups: steps 1-3 of hu_followup.SCHEDULE (+3 / +7 / +14 days) ────

def followup_body(step: int) -> str:
    if step == 1:
        blocks = [
            random.choice([
                "Csak felhozom az előző levelemet, hátha elsüllyedt.",
                "Csak visszajelzek a múltkori levelemre, hátha lecsúszott.",
            ]),
            random.choice([
                "A 2 perces videó, amin a program egy valódi ház tervein "
                "dolgozik, megvan. Elküldhetem?",
                "A videó megvan, 2 perc az egész. Elküldhetem?",
            ]),
        ]
    elif step == 2:
        blocks = [
            random.choice([
                "Egy kérdést hadd tegyek fel: egy ház terveinek felmérése "
                "most mennyi időt visz el Önöknél?",
                "Egy kérdés csak: mennyi idő most Önöknél, mire egy ház "
                "terveiből kész a mennyiségkimutatás?",
            ]),
            "A videóban látszik, mennyi marad ebből, ha a program végzi az "
            "első kört, és a kolléga csak ellenőriz.",
            "Elküldhetem?",
        ]
    else:
        blocks = [
            random.choice([
                "Nem szeretnék tolakodni, ez az utolsó levelem ebben az ügyben.",
                "Nem akarok tolakodni, ezért ez az utolsó levelem emiatt.",
            ]),
            random.choice([
                "Ha most nem aktuális, az teljesen rendben van.",
                "Ha most nem téma, azt is teljesen megértem.",
            ]),
            "Ha később mégis jól jönne, elég egy szó, és küldöm a videót.",
        ]
    return "\n\n".join(["Jó napot kívánok!", *blocks, "Üdvözlettel:\nPirint Milán"])


# ── Reply types and what happens next ─────────────────────────────────────

REPLY_TYPES = {
    "interested": ("Érdekli",      "Videó elküldése"),
    "not_now":    ("Most nem",     "Később újra, kézzel"),
    "no":         ("Nem",          "Lezárva"),
    "optout":     ("Leiratkozott", "Lezárva, tiltólistán"),
}


# ── Contacted / do-not-contact (DNC list shared with márka mode) ──────────

dnc = marka.dnc
add_dnc = marka.add_dnc


def contacted() -> dict:
    try:
        raw = json.loads(CONTACTED_PATH.read_text(encoding="utf-8"))
    except (FileNotFoundError, ValueError):
        return {}
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
    """Nothing is blocked by older HU lead files in this mode (user,
    2026-11-11): builders who got the quoting-assistant email may get this
    one, it is a different offer. Only the do-not-contact list and
    terv_contacted.json stop a first email here."""
    return {}


def contact_block(email: str, old: dict) -> str:
    e = (email or "").strip().lower()
    if e in dnc():
        return "Tiltólistán van (leiratkozott). Soha többé nem keressük."
    if e in contacted():
        return "Ezt a címet már megkerestük egy korábbi tervolvasós körben."
    if e in old:
        return f"Ez a cím szerepel egy régebbi magyar leadlistában ({old[e]}), nem írunk neki újra."
    return ""


# ── Per-lead pipeline ─────────────────────────────────────────────────────

def process_lead(lead: dict) -> dict:
    tag = lead.get("name") or lead.get("email") or f"row {lead['row']}"
    print(f"  [terv {lead['row']}] {tag}")
    lead["niche_resolved"] = lead.get("niche", "")
    lead["hero_image"] = None
    lead["error"] = ""
    lead["email_on_site"] = "unknown"
    lead["blocked"] = lead.get("blocked", "")
    if lead["blocked"]:
        lead["subject"] = lead["body"] = lead["opener"] = ""
        return lead
    opener = ((lead.get("opener_custom") or "").strip()
              or openers().get((lead.get("email") or "").strip().lower(), "")
              or fallback_opener(lead.get("niche", "")))
    lead["opener"] = opener
    lead["subject"] = subject()
    lead["body"] = body(opener, lead.get("building", ""))
    print(f"    -> {opener[:90]}")
    return lead

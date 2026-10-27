"""
Ajánlatkészítő (job-ad) mode for the HU email sender (hu_email_app.py).

The leads are companies advertising a job for someone who writes quotes
(ajanlatkeszito_leadek.xlsx). Same shape as hu_email_mech.py: an opener plus
fixed body copy, no AI writing anything at send time.

THE OPENER
The one fact we really know is the job ad. Each row can carry a hand-written
"Nyitómondat" that says, in plain words, what the ad asks for; the user
reads it in Excel before anything goes out. A row without one falls back to
"Láttam a hirdetésüket, hogy [Pozíció] munkatársat keresnek."

ONE BODY FOR EVERY TRADE
Same text for all leads, only the trade word changes ("Nyomdáknak",
"Generálkivitelezőknek"), said the way the trade names itself. Separate
bodies per made-up category read like a consultant wrote them (user's call,
2026-10-27).

"ALREADY BUILT" IS THE POINT
Every lead's demo is built before their email goes out; the email only asks
permission to send it, since a first cold email carries no link (user's
call, 2026-10-27). So the line sits right after the opener, for everyone.
That makes it a rule for sending: build the lead's demo first.

Copy rules (same as mech mode):
  * magázás, no em dash, no link, no time figure
  * never pitch it as simple: built on their own calculation, the program
    makes a first draft, a colleague checks it
  * two follow-ups only (the shared schedule has three; hu_email_app stops
    this mode after step 2)
"""

import random
import re

import hu_email_gen as heg

MAX_STEP = 2

# (keyword in the sheet's Niche cell, how the email names the trade).
# First match wins, so the more specific keyword goes first.
_TRADE_DATIVE = [
    ("nyomd",              "Nyomdáknak"),
    ("nyílászáró",         "Nyílászárógyártóknak"),
    ("kaputechn",          "Kaputechnikai cégeknek"),
    ("bútor",              "Bútorgyártóknak"),
    ("textil",             "Ponyva- és kárpitkészítőknek"),
    ("elektronik",         "Elektronikai gyártóknak"),
    ("acélszerkezet",      "Acélszerkezet-gyártóknak"),
    ("fémszerkezet",       "Fémszerkezet-gyártóknak"),
    ("lemezmegmunk",       "Lemezmegmunkáló cégeknek"),
    ("automatiz",          "Automatizálással foglalkozó cégeknek"),
    ("generálkivitel",     "Generálkivitelezőknek"),
    ("épületgépész",       "Épületgépészeti cégeknek"),
    ("csőszerel",          "Épületgépészeti cégeknek"),
    ("villanyszerel",      "Villanyszerelő cégeknek"),
    ("homlokzat",          "Homlokzatburkoló cégeknek"),
    ("házépít",            "Házépítő cégeknek"),
    ("irodaépít",          "Irodaépítéssel foglalkozó cégeknek"),
    ("magasépít",          "Építőipari cégeknek"),
    ("építőip",            "Építőipari cégeknek"),
]


def trade_dative(niche: str) -> str:
    n = (niche or "").lower()
    return next((d for k, d in _TRADE_DATIVE if k in n), "")


_VOWELS = "aáeéiíoóöőuúüű"


def _article(word: str) -> str:
    """'az' before a vowel sound, 'a' otherwise. Numbers and acronyms are
    rare in a job title, so first letter is good enough here."""
    w = (word or "").strip().lower()
    return "az" if w[:1] in _VOWELS else "a"


_QUOTE_WORDS = ("ajánlat", "kalkul")


def clean_position(pos: str) -> str:
    """One readable title out of a job-board title.

    Real ads stack several roles in one title ('Kalkulációs mérnök / Műszaki
    ajánlatkészítő', 'Ajánlatkészítő - Rendeléskoordinátor - Heti 3-5 nap').
    Pasting all of it into a sentence reads like a scraper wrote it, so this
    keeps the one part about quoting, else the first part. Then it drops a
    bracketed tail and a trailing munkatárs/kolléga (no 'munkatárs
    munkatársat') and lowercases, since an ALL CAPS title is shouting, not
    an acronym."""
    p = re.sub(r"\s*[\(\[].*?[\)\]]\s*", " ", pos or "")
    parts = [s.strip() for s in re.split(r"\s+[/–-]\s+|\s*/\s*|,\s*", p) if s.strip()]
    if not parts:
        return ""
    p = next((s for w in _QUOTE_WORDS for s in parts if w in s.lower()), parts[0])
    p = re.sub(r"\s+(munkatárs|munkatársat|kolléga|kollégát)$", "", p,
               flags=re.IGNORECASE)
    p = p.strip(" ,.;:")
    # Lowercase all but real acronyms ('CAD tervező' keeps CAD).
    return " ".join(w if (w.isupper() and len(w) <= 4) else w.lower()
                    for w in p.split())


def build_opener(position: str, custom: str = "") -> str:
    """The row's own Nyitómondat if it has one, else the plain fallback."""
    custom = (custom or "").strip()
    if custom:
        return custom
    return f"Láttam a hirdetésüket, hogy {clean_position(position)} munkatársat keresnek."


def ajanlat_subject(position: str) -> str:
    pos = clean_position(position)
    art = _article(pos)
    return f"{art[0].upper()}{art[1:]} {pos} pozícióról"


def ajanlat_body(opener: str, first: str, niche: str = "") -> str:
    trade = trade_dative(niche)
    intro = (f"{trade} építek ilyen programokat." if trade
             else "Ilyen programokat építek.")
    blocks = [
        f"Kedves {first}!" if first else "Jó napot kívánok!",

        opener + " Összeraktam Önöknek egy ajánlatkészítő programot, ami "
        "ebben a munkában segít.",

        # Sample prices said up front keeps "built for you" true: the shape
        # is theirs, the numbers come once they share them.
        intro + " Minden cég máshogy számol, ezért az éles verziót mindig az "
        "adott cég saját kalkulációjára építem, abból kiindulva, ahogy a "
        "kollégák most dolgoznak. Egyelőre mintaárakkal működik.",

        "A program a beérkező ajánlatkérésből elkészíti az ajánlat első "
        "változatát, a kolléga pedig ellenőrzi és kiküldi. Ahol bizonytalan, "
        "ott szól.",

        # "kevesebbet gépel" undersold the engineers on the list.
        "Így akit most felvesznek, kevesebb időt tölt ismétlődő munkával, "
        "több ajánlatot ad ki, és a bonyolultabb munkákra marad ideje.",

        "Csak megmutatnám, mást nem kérek. Elküldhetem?",
        "Ha nem érdekli, elég egy szó, és többé nem keresem.",
        # No name here: the Gmail signature below already carries it.
        "Üdvözlettel,",
    ]
    return "\n\n".join(blocks)


def process_lead(lead: dict) -> dict:
    """Ajánlat version of hu_email_gen.process_lead. Same output keys, plus
    'blocked', set only when the row has no position (no fact to open on)."""
    tag = lead.get("name") or lead.get("email") or f"row {lead['row']}"
    print(f"  [ajanlat {lead['row']}] {tag}")

    first = heg._first_name_for(lead)
    lead["niche_resolved"] = "ajánlatkészítő"
    lead["hero_image"] = None
    lead["error"] = ""
    lead["email_on_site"] = "unknown"
    lead["blocked"] = ""

    position = lead.get("position", "")
    if not clean_position(position):
        lead["blocked"] = "Nincs pozíció a sorban, nincs mire hivatkozni."
        lead["subject"] = lead["body"] = lead["opener"] = ""
        return lead

    opener = build_opener(position, lead.get("opener_custom", ""))
    lead["opener"] = opener
    lead["subject"] = ajanlat_subject(position)
    lead["body"] = ajanlat_body(opener, first, lead.get("niche", ""))
    print(f"    -> {opener[:90]}")
    return lead


# ── Follow-ups: steps 1 and 2 of hu_followup.SCHEDULE, nothing after ──────

def followup_body(step: int, first: str, niche: str = "") -> str:
    g = f"Kedves {first}!" if first else "Jó napot kívánok!"
    if step == 1:
        blocks = [
            g,
            random.choice([
                "Csak felhozom a korábbi levelemet, hátha elsüllyedt.",
                "Csak visszajelzek a múltkori levelemre, hátha lecsúszott.",
            ]),
            random.choice([
                "A program, amit Önöknek raktam össze, megvan, csak szólnia "
                "kell. Elküldhetem?",
                "Az Önöknek összerakott program megvan. Elküldhetem?",
            ]),
            "Üdvözlettel,",
        ]
    else:
        blocks = [
            g,
            random.choice([
                "Nem szeretnék tolakodni, ez az utolsó levelem ebben az ügyben.",
                "Nem akarok tolakodni, ezért ez az utolsó levelem emiatt.",
            ]),
            random.choice([
                "Ha a pozíciót már betöltötték, vagy most nem aktuális, az "
                "teljesen rendben van.",
                "Ha közben megtalálták az embert, vagy most nem téma, azt is "
                "teljesen megértem.",
            ]),
            "Ha később mégis jól jönne, elég egy szó, és megmutatom.",
            "Üdvözlettel,",
        ]
    return "\n\n".join(blocks)

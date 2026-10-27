"""
Ajánlatkészítő (job-ad) mode for the HU email sender (hu_email_app.py).

The leads are companies advertising a job for someone who writes quotes
(ajanlatkeszito_leadek.xlsx). Same shape as hu_email_mech.py: a deterministic
opener plus fixed body copy, no AI writing anything.

WHY THE OPENER IS DETERMINISTIC
The one fact we really know is the job ad, so the opener says that and
nothing else: "Láttam, hogy [Pozíció] munkatársat keresnek." If the sheet's
"Mit árajánlatoz" cell is filled, it goes in quoted, in their own words,
never rephrased. No site scraping, no years in business, no city.

Copy rules (same as mech mode):
  * magázás, no em dash, no link, no time figure
  * the benefit: the person they are hiring makes many more quotes, or they
    don't need a second one
  * close stays small: "Megmutassam egy saját, korábbi árajánlatukon? Mást
    nem kérek."
  * two follow-ups only (the shared schedule has three; hu_email_app stops
    this mode after step 2)
"""

import random
import re

import hu_email_gen as heg

MAX_STEP = 2

# Trades whose quotes come out of a price table (format, material, quantity),
# which is what the body promises. Engineering roles that price from drawings
# ("műszaki előkészítő", quantity takeoff) are left out on purpose: the email
# would promise something the assistant can't do yet (user's call, 2026-10-27).
FIT_NICHES = ("nyomda", "textil", "nyílászáró", "nyilaszaro")


def is_fit(niche: str) -> bool:
    n = (niche or "").lower()
    return any(w in n for w in FIT_NICHES)

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


def build_opener(position: str, quote_what: str) -> str:
    pos = clean_position(position)
    # quote_what is no longer quoted back: a pasted „ad sentence” read like a
    # scraper wrote the email (user's call, 2026-10-27). Kept as a parameter
    # so the column can come back without touching callers.
    return f"Láttam a hirdetésüket, hogy {pos} munkatársat keresnek."


def ajanlat_subject(position: str) -> str:
    pos = clean_position(position)
    art = _article(pos)
    return f"{art[0].upper()}{art[1:]} {pos} pozícióról"


def _is_print(niche: str) -> bool:
    return "nyomd" in (niche or "").lower()


def ajanlat_body(opener: str, first: str, niche: str = "") -> str:
    """The close depends on what exists. Print shops: the nyomda demo is
    built, so it offers that. Other trades: no demo yet (it gets built after
    a reply, user's call 2026-10-27), so it only offers to show how it would
    look. Never promises work on their own old quote: that needs their
    numbers, which we don't have."""
    printer = _is_print(niche)
    blocks = [f"Kedves {first}!" if first else "Jó napot kívánok!"]
    blocks.append(opener)
    # The how-it-works paragraph is what makes it believable: the price
    # comes from THEIR table, the program only reads the email, and it says
    # when something is missing instead of guessing.
    blocks += [
        ("Nyomdáknak építek egy programot" if printer
         else "Olyan programot építek")
        + ", ami az ajánlatkérésekből árajánlatot készít. Bemásolják az "
        "ügyfél levelét, a program kigyűjti belőle az adatokat, és az Önök "
        "kalkulációs szabályai és árlistája alapján összerakja a tételes "
        "ajánlatot.",

        # Don't list the inputs: a calculator knows a price depends on more
        # than format and paper, and a short list reads as naive (user's
        # call, 2026-10-27). Admit it and say it follows THEIR rules.
        ("Tudom, hogy egy nyomdai ár sok mindentől függ" if printer
         else "Tudom, hogy egy ár sok mindentől függ")
        + ", és nem minden munka sablonos. Az árat ezért nem a gép "
        "találja ki, hanem úgy számol, ahogy Önök. Ha valami hiányzik, vagy a "
        "munka túl egyedi, nem találgat, hanem jelzi.",

        "Így a kollégának nem kell gépelnie és számolnia, csak átnézi és "
        "elküldi. Akit most felvesznek, sokkal több ajánlatot tud kiadni, és "
        "nem kell mellé még egy ember.",

        ("Összeraktam egy rövid bemutatót. Elküldhetem?" if printer else
         "Összerakhatok Önöknek egy rövid bemutatót?"),
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

    opener = build_opener(position, lead.get("quote_what", ""))
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
                "A nyomdáknak készült bemutató megvan, csak szólnia kell. "
                "Elküldhetem?",
                "A bemutató, amit nyomdáknak raktam össze, megvan. Küldhetem?",
            ] if _is_print(niche) else [
                "Szívesen megmutatom, hogyan nézne ki ez Önöknél. Megmutassam?",
                "Meg tudom mutatni, hogyan működne Önöknél. Megmutassam?",
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

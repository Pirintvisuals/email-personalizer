"""
DM Generator — Flask web app
Open the profile in Chrome → come back → click Generate.
Scrolls through the page, takes screenshots, sends to Groq Vision.
"""

import os, re, base64, random
from pathlib import Path
from flask import Flask, request, jsonify, render_template
from groq import Groq
from google import genai as _genai
from en_niche import (
    detect_niche as _en_detect_niche,
    detect_niche_from_name as _en_detect_niche_from_name,
    opener_matches_niche as _en_opener_matches_niche,
    NICHE_OPENER_INSTRUCTIONS as _EN_NICHE_OPENER_INSTRUCTIONS,
)
try:
    import pyperclip as _pyperclip
    _CLIPBOARD_OK = True
except ImportError:
    _CLIPBOARD_OK = False

# ── Load .env ──────────────────────────────────────────────────────────────────
_env = Path(__file__).parent / ".env"
if _env.exists():
    for line in _env.read_text().splitlines():
        if "=" in line and not line.startswith("#"):
            k, v = line.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip())

API_KEY          = os.environ.get("GROQ_API_KEY", "")
GEMINI_API_KEY   = os.environ.get("GEMINI_API_KEY", "")
CDP_PORT         = int(os.environ.get("CDP_PORT", "9222"))
DEFAULT_EXCEL    = os.environ.get("DEFAULT_EXCEL", r"C:\Users\pirin\OneDrive\Asztali gép\USA_Leads_Master.xlsx")
GRAMMAR_CHECK    = False  # set to True in dm_app_hu.py
GREETING_PREFIX  = "Hi"
STRICT_NAME_DETECTION = True    # mirror HU: only greet with a confidently detected name
ALWAYS_GREET     = True          # mirror HU: always greet ("Hi. " when no name found)
GREETING_NAME_FMT = "Hi, {name}. "   # (legacy, unused) kept for reference
GREETING_SOLO_FMT = "Hi. "            # (legacy, unused) kept for reference


def _apply_greeting(opener: str, name_hint: str, market: str = "US") -> str:
    """Prepend a rotating casual greeting. 'Hey' / 'Hi' read like a real
    person rather than a flat 'Hi.', and it's the first thing shown in the
    Facebook message preview, so it earns the open. A greeting ending in . ! ?
    keeps the opener capitalised; a comma greeting flows into a lowercase opener
    ('Hey Wade, that boiler came up mint.'). market="UK" swaps in 'Alright'/
    'Cheers' — the only thing that should differ between US and UK is
    accent/word-choice, not substance, and the greeting is the very first
    word read, so it's worth getting right per market."""
    import random as _r
    if market == "UK":
        if name_hint:
            g = _r.choice([
                f"Alright {name_hint}? ",
                f"Alright {name_hint}, ",
                f"Cheers {name_hint}, ",
            ])
        else:
            g = _r.choice(["Alright? ", "Alright, ", "Cheers, "])
    else:
        if name_hint:
            g = _r.choice([
                f"Hey {name_hint}, ",
                f"Hi {name_hint}, ",
                f"Hey {name_hint}! ",
            ])
        else:
            g = _r.choice(["Hey, ", "Hi, ", "Hey there, "])
    keep_cap = g.rstrip()[-1:] in ".!?"
    return g + (opener if keep_cap else opener[:1].lower() + opener[1:])


LOWERCASE_DM = True   # casual all-lowercase DMs (proper nouns kept). False = normal case.


def _casual_lowercase(text: str) -> str:
    """Casual lowercasing that KEEPS proper nouns. Only lowercases the pronoun 'I'
    (and its contractions) and the first letter of each sentence; mid-sentence
    words the model/template capitalised (Worcester Bosch, WhatsApp, names, AI)
    are never touched, so brands and names stay correct."""
    if not text:
        return text
    # pronoun I / I'm / I've / I'll / I'd -> lowercase
    text = re.sub(r"\bI(?=['’ ]|$)", "i", text)
    # first letter of each sentence -> lowercase, unless it's a camel-case brand
    def _lower_start(m):
        lead, word = m.group(1), m.group(2)
        if len(word) > 1 and any(c.isupper() for c in word[1:]):
            return m.group(0)                       # e.g. WhatsApp -> keep
        return lead + word[0].lower() + word[1:]
    return re.sub(r"(^|[.!?]\s+)([A-Za-z][^\s.!?]*)", _lower_start, text)


def _dm_case(text: str) -> str:
    """Apply casual lowercase to a finished DM when LOWERCASE_DM is on."""
    return _casual_lowercase(text) if LOWERCASE_DM else text
# Your real, current social proof. EDIT THIS as your case studies grow. Specific
# trades/towns beat "a couple of trades"; a real result beats usage. No em dashes.
SOCIAL_PROOF = "got a couple of trades using one already"
CUSTOM_NAME_EXTRACTOR = None               # if set, overrides extract_first_name (callable: text -> str)
CUSTOM_PROMPT_BUILDER = None               # if set, called as (full_prompt, text_data) -> str before vision call
CUSTOM_OPENER_POSTPROCESS = None           # if set, called as (opener, text_data) -> str after generation
# If set, called as (text_data) -> bool BEFORE any screenshot is taken. True
# means this lead's opener can't come from a photo, so skip the post
# screenshot, the vision call and the profile screenshot entirely and let
# CUSTOM_OPENER_POSTPROCESS build the opener from the page text instead.
# Exists because those three steps are the whole per-lead wait (the vision
# call alone is 12s per attempt, up to 26s with the retry, and has been
# measured at 66s), and for a niche with no finished job to photograph the
# result was being discarded anyway.
SKIP_VISION_FOR = None
VISION_MODELS  = ["meta-llama/llama-4-scout-17b-16e-instruct", "meta-llama/llama-4-maverick-17b-128e-instruct"]
TEXT_MODELS    = ["llama-3.3-70b-versatile", "llama-3.1-8b-instant"]
# Facts extraction is trivial bullet-pulling from explicit text — the fast 8B
# does it identically to the 70B, and it's the per-lead bottleneck. Quality of
# the OPENER comes from the vision call, which is untouched. 70B kept as fallback.
FACTS_MODELS   = ["llama-3.1-8b-instant", "llama-3.3-70b-versatile"]

# ── Message variant (A/B test toggle) ────────────────────────────────────────
# "full"     = current opener + full pitch body (default, unchanged behaviour).
# "stripped" = one-line direct variant. No personalised opener, so NO vision call
#              and NO screenshot — faster too. Reads:
#              "Hi [name]. I make AI quote chatbots for [trade] and I've already
#               built a sample for [company]. Can I send it over?"
# Persisted to settings.json so the toggle survives app restarts.
_SETTINGS_PATH = Path(__file__).parent / "settings.json"

def _load_variant() -> str:
    try:
        import json
        v = str(json.loads(_SETTINGS_PATH.read_text()).get("variant", "full")).strip().lower()
        return v if v in ("full", "stripped", "british", "funny") else "full"
    except Exception:
        return "full"

def _save_variant(v: str) -> None:
    try:
        import json
        _SETTINGS_PATH.write_text(json.dumps({"variant": v}))
    except Exception as e:
        print(f"  Could not save variant setting: {e}")

MESSAGE_VARIANT = _load_variant()

PROMPT = """\
You write the opening line of a cold Facebook DM. Return ONLY the opening line — nothing else. No explanation, no emoji, no extra text.

You're looking at a screenshot of a UK trade business Facebook page, scrolled to their freshest, most recent post. Look at the photo carefully. React to a finished job — like texting a mate who actually knows the trade.

This is their newest post, not a random one — so react like you just saw it go up, and be as specific as the photo actually lets you be. A vague reaction ("that bathroom came up lovely") that could describe almost any job in this trade is a FAILURE — every line must include at least one detail (a material, a brand, a colour, a fixture, an install detail) that proves this is THIS specific post, not a generic trade compliment.

PRIORITY — PROFESSIONAL DETAIL:
If you can see any of these, react to it — they're far more convincing than a generic "looks good":
- Brand name (see the list below)
- Material type (porcelain, Indian sandstone, EPDM, resin, composite, slate, lead, block paving, etc.)
- Placement / quality of the install (joint width, fall/slope, pipework run, edging course, pattern match)

PAIRING — TWO DETAILS THAT GO TOGETHER:
If two elements in the photo obviously complement each other, mention both. This is the detail only someone who actually looked would catch. Examples of strong pairings:
- Tile colour/finish + brassware colour ("that copper rusteel tiling with the matte black taps")
- Frame colour + opening type ("anthracite tilt-and-turn, deeply set into the reveal")
- Deck board colour + fascia/edging ("composite boards with the aluminium trim")
- Boiler brand + a specific fitting detail ("Worcester Bosch combi, magnetic filter sat neat on the return")
When a pairing is obvious, structure the line as: [focal element], [second complementary element] [verb]. Keep it under 16 words total.
- Size / scale of the project (how many panels, how big an area, a large system)
- A detail only a tradesperson would notice (ventilation, drainage fall, substructure, flashing detail, ridge line, gutter fall, lead soaker around a chimney, board spacing, membrane layer showing)

WINDOW / DOOR PHOTOS — look specifically for these:
- Frame colour: anthracite, dark brown, golden oak, white — how it suits the facade
- Opening type: tilt and turn, sliding / lift-and-slide, bi-fold — instantly professional
- Cill slope: is there a visible fall for drainage (a flat cill is a fault)
- Reveal depth: deeply set frame vs flush fit — only a tradesperson notices
- Frame-to-wall seal: taped or just rendered — taped is the premium detail
- On a front door: threshold meeting the floor finish, the seal compressing
- On a roof window (Velux/Fakro): flashing laid in, tiles returned around the frame
- On an integrated roller shutter: is the box built into the wall or sticking out

BRAND NAMES BY NICHE — if you see them, use them:
- Air con / split / heat pump: Daikin, Mitsubishi, LG, Samsung, Panasonic, Fujitsu, Midea, Gree, Toshiba
- Boiler / heating: Worcester Bosch, Vaillant, Ideal, Baxi, Viessmann, Glow-worm, Potterton
- Solar: LONGi, JA Solar, Jinko, Canadian Solar, Trina, SolarEdge (inverter), SMA, Fronius, GivEnergy
- Roofing: Marley, Redland, Sandtoft, Velux, Fakro, GRP/fibreglass, EPDM, lead
- Windows / doors: Schüco, Rehau, Veka, Liniar, Origin, Residence 9
- Paving / driveway: Marshalls, Bradstone, Brett, Tobermore, resin bound
- Decking: composite, Trex, Millboard, Cladco, hardwood

PROFESSIONAL EXAMPLES — this is what it looks like when someone knows the trade:
Air con / split (vary the focal point — the outdoor/condenser unit, the indoor head or cassette, the lagged line set, the wall bracket, the trunking, the condensate drain, the grille, where it's sited — NOT always "pipework run". Only what's actually visible):
"That Daikin condenser in Leeds came up mint, line set lagged dead neat."
"That Mitsubishi split in Bristol is sat lovely, indoor head level on the wall."
"That ceiling cassette in Leeds looks the part, grille sat flush in the ceiling."
"That outdoor unit in York is bang on, trunking clipped tight down the wall."
"That multi-split in Manchester looks proper tidy, two heads matched up neat."
"That condensate run in Sheffield is spot on, drain kept tidy off the unit."
Boiler / heating (vary the focal point — flue, condensate pipe, gas pipe, isolation valves, system/magnetic filter, the boiler model & position, cupboard mount, cylinder, manifold — NOT always "pipework on the wall". Only what's actually visible):
"That Worcester Bosch in Manchester looks proper tidy, flue routed clean through the wall."
"That combi swap in Leeds came up mint, isolation valves all lined up neat."
"That underfloor heating manifold in Leeds is bang on, neat loops off the ports."
"That Vaillant in Bristol looks the part, magnetic filter sat neat on the return."
"That boiler in York is sat lovely, condensate run kept tight to the wall."
"That cylinder install in Sheffield came up a treat, tidy tank and pump set."
"That gas run in Derby is proper tidy, clipped dead straight along the wall."
Solar:
"That LONGi array in York looks spot on, even pitch across the panels."
"That solar roof in Sheffield looks serious, proper symmetry on the layout."
Roofing:
"That Marley tiled roof in Leeds looks spot on, even coursing across the lot."
"That standing seam roof in York came up mint, clean verge line."
"That EPDM flat roof in Bolton looks proper tidy, neat edge trim."
"That lead flashing in Oxford is bang on, tight detail round the chimney."
Paving / driveway:
"That block paving in Coventry looks spot on, even joints right across."
"That resin driveway in Cardiff came up mint, proper clean finish."
Tiling / bathroom (vary the focal point — basin, suite, screen, niche, bath panel, tile pattern, brassware — instead of always grout. These show TONE and STRUCTURE only: only mention a detail you can ACTUALLY SEE, never borrow these specifics if they're not in the photo):
"That walk-in shower in Leeds came up mint, tiled niche sat dead level."
"That bathroom in Manchester is bang on, tidy floor-to-wall transition."
"That freestanding bath in Bristol looks the part, neat panel reveal round it."
"That metro tiling in York is proper tidy, crisp brick-bond lines."
"That wet room in Sheffield came up a treat, fall to the drain dead on."
"That vanity unit in Derby looks cracking, basin and brassware lined up neat."
"That herringbone floor in Leeds is bang on, tight joints right across."
Windows / doors:
"That Schüco anthracite window in Leeds looks spot on, deeply set frame."
"That Rehau tilt and turn in Bristol looks tidy, proper fall on the cill."
"That bi-fold in York came up mint, slim frame profile right across."
"That front door in Manchester looks bang on, tidy threshold to the floor."
"That Velux in Sheffield is spot on, flashing returned proper neat."
Garden / decking:
"That composite decking in Leeds looks spot on, even board spacing."
"That garden in York looks serious, proper line on the retaining levels."
Render / facade:
"That K Rend facade in Sheffield looks spot on, even coverage across the lot."
"That render in Manchester is proper tidy, sharp corner beads."
Painting / decorating:
"That exterior in Leeds came up spot on, even coverage right across."
"That interior in Manchester looks tidy, clean cutting-in round the frames."

TONE: Casual, British, working class. Personality. Like a text from someone who genuinely rates good work. Use words like: spot on, mint, cracking, proper, came up a treat, looks the part, bang on, tidy.

WHEN TO USE THE FALLBACK (compulsory, no exceptions):
- You can't see a finished trade job in the photo
- You only see text, a logo, a document, a certificate or an AI image
- The photo is a branded VAN / vehicle, a sign / banner, a business card, or a team / headshot photo — none of these are a finished job. Do NOT describe the van or the branding ("that silver car looks spot on" is wrong). → fallback
- You can't work out the location
- You CAN'T name a specific material, brand or professional detail — the location alone is NOT enough
- Generic words like "wallpaper", "painting", "tiling", "roof" ON THEIR OWN aren't enough — they need a colour, technique, brand or specific detail alongside. If there's none → fallback.
Fallback: "Came across your page in [location]." — town or city only.

STRICT RULES:
- Return ONLY the opening line. No other text.
- No emoji, no exclamation marks, no em dashes.
- Don't copy the examples word for word.
- 10-16 words max.
- EVERY sentence starts with "That" — the only exception is the fallback.
- BANNED headline style: "Nice job on the...", "Great work on the...", "Lovely finish on..." — banned.
- BANNED passive endings: "nicely finished", "well executed", "well laid", "nicely done" — banned. Just describe what you can actually see in a concrete second clause.
- VARY THE VERB: do NOT default to "looks spot on" every time — rotate naturally across "came up mint", "proper tidy", "bang on", "looks the part", "came up a treat", "is sat lovely", "spot on". Reusing the same phrase on every job reads like a template and kills it.
- NAME THE SPECIFIC THING you can see, and pick whatever is genuinely the STANDOUT in THIS photo. Do NOT let any one detail become the default across a batch — name the actual focal point of THIS particular image.
  - Bathroom menu (wide): the bath or freestanding tub, the shower or enclosure/screen, the WC, the flooring, the tiling layout (metro, herringbone, brick-bond), the vanity unit, the recessed niche, the towel rail/radiator, the mirror, the basin, the brassware/taps. OVERUSED, avoid as default: "grout lines", "basin / brassware".
  - Boiler / heating menu (wide): the flue, the condensate pipe, the gas pipe, the isolation valves, the system/magnetic filter, the boiler model & where it's sat, the cupboard mount, the cylinder, the manifold. OVERUSED, avoid as default: "pipework on the wall", "clean pipework".
  - Air con / HVAC menu (wide): the outdoor/condenser unit, the indoor head or cassette, the lagged line set, the wall bracket, the trunking/conduit, the condensate drain, the grille, where it's sited, the brand. OVERUSED, avoid as default: "tidy pipework run", "Daikin unit ... pipework run".
- BUT only describe what's genuinely visible — accuracy beats variety. If the only thing you can clearly see is plain tiling and grout, say that honestly, or use the fallback. NEVER invent a fixture, pattern, brand or colour just to sound specific or to avoid repeating yourself.
- BANNED generic endings: "looks well-finished", "looks well-built", "looks like a lot of work", "nice work", "good finish", "grout lines", "tight grout lines", "square grout lines", "dead straight grout lines", "crisp grout lines", "connections to the wall", "pipework to the wall", "against the wall", "to the wall", "across the wall", "to the system", "connections to the system", "crisp connections", "clean connections", "sharp connections", "square connections", "clean lines", "crisp lines" — banned. The verb ALWAYS needs a concrete named element (brand, component, fitting, material) — not a direction or generic surface.
- COMPULSORY: every opening line needs a COMMA-separated second clause with a concrete observation. e.g. "...came up mint, tidy detail round the frame." — the sentence never ends just on the location or the name of the object. The second clause must be a PHYSICAL detail of the work (a fit, a run, a level, a finish, a join) — NEVER a location or a phrase like "[town] installation" / "[town] job". If a location is all you've got, that's a fallback, not a second clause.
- No "roof job" — say the type: flat roof, pitched roof, lead roof, fibreglass roof, re-slate.
- No street names, road names, house numbers or postcodes — town or city only.
- No "in the photo" or "in the post". No "apparently" — never reference the caption. No commenting on their feed or post count.
- NEVER put the business / company name into the line. Describe the work, the brand, or the location — not the company's name. WRONG: "That DAJ Gas Services radiator...", "That Janie McK Heating Services installation...". RIGHT: "That radiator looks proper tidy, neat connections to the wall.", "That boiler swap came up mint, cylinder set sat neat." If you have no brand and no location, describe the work plainly or use the fallback.
- No inventing specs — only describe what you can ACTUALLY SEE in the photo. This rule OVERRIDES the "vary it" and "be specific" rules above: if being specific would mean guessing, don't — describe only what's really there, or use the fallback. A plain accurate line always beats a specific made-up one.
- For windows NEVER write "plastic window" — write "uPVC window" or the brand name.
- If you see a brand name in the photo, you MUST work it in — it's not optional.
- BRAND-NICHE MATCH: Air con / cooling brands (Daikin, Mitsubishi, Fujitsu, Midea, Gree, Toshiba) are only for air con businesses. Boiler / heating brands (Worcester Bosch, Vaillant, Baxi, Ideal, Viessmann) are only for heating businesses. Don't mix them.
- NEVER write "nice finish" — always name EXACTLY what's good: the joint width, the fit, the fall, the pipework run, the board spacing, the ridge line, etc.
- If a colour or material colour is visible (anthracite, graphite grey, natural, white, brown), name it too."""

# ── DM body copy — one per trade/sub-trade, no em dashes ─────────────────────

DM_BODY_ROOFING     = "heard a few roofers saying they lose jobs just cos someone else got back quicker. you find that?"
DM_BODY_DRIVEWAY    = "heard a few driveway lads saying enquiries come in then go dead after the quote. same with you?"
DM_BODY_PATIO       = "heard a few patio guys saying enquiries come in then go dead after the quote. same with you?"
DM_BODY_LANDSCAPERS = "heard a few landscapers saying they get enquiries then nothing after the quote. same with you?"
DM_BODY_BUILDERS    = "heard a few builders saying half their enquiries just want a ballpark then disappear. same with you?"
DM_BODY_RENOVATION  = "heard a few renovation lads saying they go price a job then just hear nothing back. same with you?"
DM_BODY_QUOTES      = "heard a few trades saying enquiries come in then go dead after the quote. same with you?"

# ── Follow-up sequences — 3 touches per trade ────────────────────────────────
# Touch 1 (Day 3):  new angle / peer curiosity — sounds like genuine interest
# Touch 2 (Day 10): social proof — concrete result from a similar business
# Touch 3 (Day 21): break-up / loss aversion — walk away, highest reply rate

FOLLOWUPS = {
    "roofers": [
        "Spoke to a roofer in Sheffield last week who said he loses half his enquiries just because he couldn't get back same day, does that happen much your end?",
        "Helped a roofing lad in Leeds go from missing calls to booking an extra 3-4 jobs a week, just by having enquiries handled while he was on the roof. Thought it might be relevant.",
        "Seems like the timing's not right. No worries, I'll leave it here. If you ever want to chat about getting more out of your enquiries, you know where I am.",
    ],
    "driveways": [
        "Spoke to a driveway firm in Birmingham who reckoned he was losing 2 jobs a week to lads who just quoted cheaper, do you find price shopping's got worse lately?",
        "Helped a driveway company in Manchester stop chasing dead quotes and start closing the ones that matter, 6 extra jobs booked in the first month. Might be worth a look.",
        "Looks like the timing's off. I'll leave this here. If it ever becomes relevant, feel free to drop me a message.",
    ],
    "patio": [
        "Spoke to a patio contractor in Bristol who said he spends half his time quoting jobs that go nowhere, do you get a lot of that round your way?",
        "Helped a landscaping firm in Leeds filter out the time wasters and only spend time on serious enquiries, their close rate went from 1 in 5 to 1 in 2. Thought that might be useful.",
        "Looks like it's not the right time. No worries, I'll leave this here and you can always come back to it.",
    ],
    "landscapers": [
        "Spoke to a landscaper in Nottingham who said most of his best jobs still come through word of mouth but he can't make it consistent, is that the same for you?",
        "Helped a garden design firm in the Midlands turn their Facebook page into a steady stream of enquiries, went from sporadic word of mouth to 8-10 leads a week. Worth a chat?",
        "Seems like now isn't the time. I'll leave it there. If that changes, you know where I am.",
    ],
    "builders": [
        "Spoke to a builder in Leeds last week who said 3 out of 5 enquiries are people who just want a rough idea and disappear, do you find that's got worse recently?",
        "Helped a building firm in Manchester filter out the time-wasters and focus on serious enquiries, they went from 10 enquiries a week to 6 decent ones that actually go somewhere. Might be relevant.",
        "Looks like the timing's not right. No worries, I'll leave it here. Feel free to reach out if anything changes.",
    ],
    "renovation contractors": [
        "Spoke to a renovation contractor in London who said he spends days pricing jobs that go quiet after he sends the quote, does that happen much your end?",
        "Helped a renovation company in the South East cut their quote-to-job time in half by changing how they follow up, stopped losing jobs to silence. Thought it might be relevant.",
        "Seems like now isn't the right time. I'll leave it there. If you ever want to talk about it, feel free.",
    ],
    "quotes": [
        "Spoke to a contractor last week who said his quotes go cold about 60% of the time, no reply, no reason. Do you find that with yours?",
        "Helped a trades business in the Midlands go from chasing cold quotes to having clients chase them, changed one thing in how they follow up. Happy to share if it's useful.",
        "Looks like the timing's off. I'll leave this here. You know where I am if anything changes.",
    ],
}

# Follow-up sequence: 3 touches, widening gaps (≈ Day 3 / Day 10 / Day 21).
FOLLOWUP_TOUCHES = 3

def _clean_company_name(name: str) -> str:
    """Tidy a raw business name so it reads natural mid-sentence.
        'Pierce - Your Local Gas/Heating Engineer' -> 'Pierce'
        'Beagle Heating, Plumbing, Gas'            -> 'Beagle Heating'
        'Ava Bathrooms Ltd'                        -> 'Ava Bathrooms'
    Returns '' when nothing usable is left so the caller can fall back to
    'your business' instead of dumping a messy string into the message."""
    n = (name or "").strip()
    if not n:
        return ""
    # Cut at the first tagline separator ("Name - tagline", "Name | tagline"...)
    for sep in (" - ", " – ", " — ", " | ", " · ", " • ", ": ", " / "):
        i = n.find(sep)
        if i > 0:
            n = n[:i]
            break
    # Cut a comma list down to the lead name ("Beagle Heating, Plumbing, Gas")
    if "," in n:
        n = n.split(",", 1)[0]
    n = n.strip(" .-–—|/")
    # Drop a trailing parenthetical ("Air Con Services (UK)" -> "Air Con Services")
    n = re.sub(r"\s*\([^)]*\)\s*$", "", n).strip()
    # Drop trailing legal suffixes
    n = re.sub(r"\s*\b(ltd\.?|limited|llp|llc|plc|inc\.?|co\.?|company)\b\.?$", "", n, flags=re.I).strip()
    if not re.search(r"[A-Za-z]", n):
        return ""
    words = n.split()
    if len(words) > 6:           # still descriptor-laden — cap so it stays readable
        n = " ".join(words[:4])
    n = n.strip()
    # Facebook Page data is often ALL CAPS ("KGAS BRADFORD"). Dropping that raw
    # into a DM reads like a scraped field, not something a person typed. Rewrite
    # word-by-word, but leave short all-letter tokens (<=3 chars) alone since
    # those are usually a real acronym (AC, GT, UK) — except common short
    # connector words, which are never acronyms even when short.
    _CONNECTOR_WORDS = {"and", "or", "of", "the", "for", "in", "on", "at", "to", "by"}
    if n.isupper():
        n = " ".join(
            w if w.isalpha() and len(w) <= 3 and w.lower() not in _CONNECTOR_WORDS
            else w[:1] + w[1:].lower()
            for w in n.split(" ")
        )
    return n


# Words that make a "name" actually a generic descriptor / location, not a brand.
_PITCH_GENERIC = {
    "the", "and", "your", "local", "new", "quality", "professional", "pro", "group",
    "bathroom", "bathrooms", "bath", "baths", "kitchen", "kitchens", "wetroom", "wetrooms",
    "fitting", "fittings", "fitter", "fitters", "fitted", "fit", "renovation", "renovations",
    "plumbing", "plumber", "plumbers", "heating", "gas", "tiling", "tiler", "tilers",
    "tile", "tiles", "shower", "showers", "installation", "installations", "electrical",
    "electric", "roofing", "roofer", "roofers", "landscaping", "landscaper", "landscapers",
    "building", "builder", "builders", "construction", "services", "service", "solutions",
    "design", "designs", "interiors", "studio", "studios", "company", "centre", "center",
    "co", "ltd", "uk", "aircon", "air", "conditioning", "transformations",
}
_PITCH_PLACES = {
    "glasgow", "edinburgh", "manchester", "london", "leeds", "bristol", "cardiff",
    "liverpool", "birmingham", "sheffield", "newcastle", "worcester", "carlisle",
    "nottingham", "leicester", "coventry", "hull", "preston", "wales", "scotland",
    "yorkshire", "kent", "essex", "surrey", "midlands", "lancashire", "cheshire",
}


def _who_for_pitch(company_name: str) -> str:
    """How to address them in 'I've set one up for ___ too'. Most trade names read
    badly there ('for Bathroom Fitting too', 'for New Bathrooms Glasgow too'), so
    default to 'you' and only use the real name when it's a short, genuinely
    distinctive brand ('Ava Bathrooms', 'Downs Brothers')."""
    clean = _clean_company_name(company_name)
    if not clean:
        return "you"
    # Domain / handle ("7bathrooms.uk", "acme.co.uk") reads like a URL, not a name.
    if re.search(r"\.(uk|com|co|net|org|io|biz)\b", clean.lower()) or "www." in clean.lower() or "@" in clean:
        return "you"
    words = [w for w in clean.split() if w != "&"]
    if len(words) > 2:                      # long names sound like a descriptor here
        return "you"
    distinctive = [w for w in words
                   if w.lower().strip(".&") not in _PITCH_GENERIC
                   and w.lower().strip(".&") not in _PITCH_PLACES]
    return clean if distinctive else "you"  # all-generic/location -> "you"

def _followup_text(touch: int, company_name: str = "", niche: str = "") -> str:
    """Niche-specialised nudge — references the actual trade/job and the sample,
    in the same voice as the opener, so it reads tailored rather than generic.
    Two wordings per touch for spam-safety."""
    trade_noun, job, examples = _pitch_for(niche)
    who = _clean_company_name(company_name) or "your business"
    t = max(1, min(touch, FOLLOWUP_TOUCHES))
    if t == 1:
        return random.choice([
            f"Hi again, no rush at all. That sample I put together for {who} gives someone a rough price for {job} in about 30 seconds, then passes you their details. Happy to send it over whenever, just say the word.",
            f"Hey, no pressure on this. The sample I built for {who} shows someone a ballpark for {job} in seconds and sends you their info. Happy to drop it over whenever suits.",
        ])
    if t == 2:
        return random.choice([
            f"Just circling back once. It asks a couple of quick questions ({examples}), gives them a ballpark, and only passes you the serious ones. Want me to send the sample over?",
            f"Coming back to this one more time. The whole point is it sorts the time wasters from the real ones before they reach you, so you only quote the jobs worth quoting. Worth a quick look?",
        ])
    return random.choice([
        "All good if it's not for you, I'll leave it here. If you ever fancy a quick look at the sample, just give me a shout. Cheers.",
        "No worries if it's not your thing, I'll leave it here. The offer's open if you ever want a look. Cheers.",
    ])

VOICE_NOTE_PROMPT = """\
You write a short WhatsApp voice note script for someone reaching out cold to a UK landscaping business owner. A screenshot is being sent in the same message as this voice note. Output ONLY the script — no labels, no explanation, no quotes. Write it as natural spoken words.

This is the FIRST message. The recipient can already see the sender's business name on WhatsApp so there is no need to introduce yourself.

TONE: Casual, British, like a real person. Not a salesperson. Use natural British expressions. Sound like a mate who spotted something interesting, not an agency pitching a product.

STRUCTURE — 5 parts, under 90 words total (must be speakable in under 45 seconds):

1. GREETING + HOOK
Start with "Cheers [name]," — find the owner's first name from the About section ("Owner at", "Founder of", "Director"), the page name if it includes a person's name, or post captions. Never use the company name as a name.
Then reference ONE specific post — name the actual material, job type, and a visual detail or location from that post. Go as specific as possible.
GOOD: "just saw that resin bound driveway with the charcoal border you finished in Bristol, looks tidy" / "noticed that split-level garden with the porcelain and sleepers you posted"
BAD: "noticed you do landscaping in Devon" / "saw you work across Yorkshire" — these read the bio, not the posts.
End the hook naturally: "looks spot on" or similar. Keep it brief.

2. SCREENSHOT REFERENCE
Direct them to look at the screenshot. One sentence.
Example: "Have a quick look at that screenshot I sent — it is basically how an enquiry looks when it hits your phone."

3. WHAT THEY GET
Explain both sides in one sentence — what the customer gets and what the owner gets.
The customer gets a rough price instantly. The owner gets all their details — budget, postcode, job type — before picking up the phone.

4. WHAT IT DOES
One sentence. The point of it is to filter out time wasters and save them grief.
Use natural British expressions: "time wasters", "save you some grief", "takes the hassle out of it".
Do NOT use: leads, solution, system, software, AI, agent, scale, grow, qualified.

5. DISQUALIFY + CTA
"Might not be for you, but if you are interested in how it works, just reply and we will have a quick chat. Cheers."

RULES:
- Under 90 words total
- The hook must be specific to THIS page — not something generic you could send to any landscaper
- British vocabulary throughout — "mate", "cheers", "time wasters", "the lot", "reckon", "grief" are all fair game
- No exclamation marks. No emojis. No over-the-top compliments.
- Sound like a real person, not a script

GOLD STANDARD EXAMPLE — match this tone exactly:
"Cheers Haydan. Just saw that garden transformation you posted over in West Yorkshire — looks spot on, mate. Have a quick look at that screenshot I sent — it is basically how an enquiry looks when it hits your phone. The customer gets a rough price instantly, and you get all their details — budget, postcode, the lot — before you even pick up the phone. I am building these for landscapers to help filter out the time wasters and save you some grief. Might not be for you, but if you are interested in how it works, just reply and we will have a quick chat. Cheers."

Bad examples — never write these:
- "I have an AI solution that generates qualified leads" (jargon, wrong)
- "Would you be open to a quick chat?" (wrong CTA structure)
- "Saw you do landscaping in Manchester..." (no specific detail, too generic)
- "It asks questions and calculates a price..." (too specific, kills curiosity)"""

app = Flask(__name__)

def _strip_overlays(page) -> None:
    """Remove/hide whatever overlay is dimming the page (login wall, cookie
    dialog, a full-screen scrim) so it doesn't sit on top of screenshots.
    Best-effort, never raises.

    Call this fresh right before EVERY screenshot that matters, not just
    once early in the flow — the previous version installed a MutationObserver
    that only kept re-cleaning for 6 seconds before disconnecting itself. The
    full pipeline (page load + text extraction + vision AI call) routinely
    takes longer than that, so by the time the profile screenshot was taken,
    the watcher had already given up — if Facebook re-showed a login/cookie
    overlay in the meantime, it just sat there un-removed, which is why the
    profile picture came out washed-out/white on EVERY capture, not just
    occasionally (a real timing race would be intermittent; a dead watcher
    is consistent). Re-running this immediately before each screenshot fixes
    that regardless of how much time has elapsed since page load."""
    try:
        page.evaluate(r"""() => {
            // 1) Persistent CSS so the login/cookie dialog stays hidden even
            //    when React re-adds it (a one-time removal alone loses that race).
            if (!document.getElementById('dm-overlay-killer')) {
                const st = document.createElement('style');
                st.id = 'dm-overlay-killer';
                st.textContent =
                    '[role="dialog"]{display:none!important;}' +
                    'html,body{overflow:auto!important;}';
                (document.head || document.documentElement).appendChild(st);
            }
            // 2) Strip full-viewport dark/semi-transparent backdrop scrims —
            //    a one-shot pass, run fresh each time this is called.
            document.querySelectorAll('[role="dialog"]').forEach(el => el.remove());
            document.querySelectorAll('body *').forEach(el => {
                const s = getComputedStyle(el);
                if (s.position !== 'fixed' && s.position !== 'absolute') return;
                const r = el.getBoundingClientRect();
                if (r.width < innerWidth * 0.8 || r.height < innerHeight * 0.6) return;
                const bg = s.backgroundColor || '';
                const op = parseFloat(s.opacity || '1');
                const m  = bg.match(/rgba?\(([^)]+)\)/);
                let a = 1; if (m) { const p = m[1].split(','); a = p.length > 3 ? parseFloat(p[3]) : 1; }
                if ((a > 0 && a < 1) || (op > 0 && op < 1) || bg === 'rgb(0, 0, 0)') el.remove();
            });
            document.documentElement.style.overflow = 'auto';
            document.body.style.overflow = 'auto';
        }""")
    except Exception:
        pass


def _scroll_to_freshest_post(page) -> bool:
    """Scroll to the actual first post element (Facebook posts are role="article")
    instead of guessing a fixed pixel offset. Cover-photo height, pinned posts,
    and About cards vary by page, so a blind scrollTo(0, 850) often lands on the
    wrong section — the vision opener then reacts to whatever's there instead of
    the freshest post, which reads as generic rather than specific. Falls back to
    the old fixed offset if no post element is found. Returns True if a real post
    element was found and targeted."""
    try:
        return bool(page.evaluate("""() => {
            const post = document.querySelector('[role="article"]');
            if (post) { post.scrollIntoView({block: 'center'}); return true; }
            window.scrollTo(0, 850);
            return false;
        }"""))
    except Exception:
        try:
            page.evaluate("window.scrollTo(0, 850)")
        except Exception:
            pass
        return False


# ── Capture existing Chrome tab ────────────────────────────────────────────────
def capture_profile_tab(hint_url: str = "") -> tuple:
    """
    Find the already-open Facebook/Instagram profile tab in Chrome.
    Scroll through it and take screenshots. Do NOT open a new tab.
    Returns (list of (bytes, mime_type), text_string)
    """
    import socket
    from playwright.sync_api import sync_playwright

    print("\nChecking Chrome debug port...")
    try:
        s = socket.create_connection(("127.0.0.1", CDP_PORT), timeout=3)
        s.close()
        print(f"Port {CDP_PORT} OPEN")
    except Exception:
        print(f"Port {CDP_PORT} CLOSED — Chrome is not running with debug port")
        return [], ""

    with sync_playwright() as p:
        try:
            browser = p.chromium.connect_over_cdp(f"http://localhost:{CDP_PORT}")
            context = browser.contexts[0]
            print(f"Connected. Open tabs: {len(context.pages)}")
            for pg in context.pages:
                print(f"  Tab: {pg.url[:80]}")
        except Exception as e:
            print(f"CDP connect error: {e}")
            return [], ""

        # Find the right tab — NEVER create a new one
        target_page = None

        # 1. Try to match hint URL if provided
        if hint_url:
            hint_clean = hint_url.rstrip("/").split("?")[0].lower()
            for pg in context.pages:
                if hint_clean in pg.url.lower():
                    target_page = pg
                    print(f"Matched hint URL: {pg.url}")
                    break

        # 2. Find any Facebook profile tab (not messenger/messages)
        if not target_page:
            for pg in context.pages:
                u = pg.url.lower()
                if "facebook.com/" in u and "messenger.com" not in u and "/messages" not in u:
                    target_page = pg
                    print(f"Found Facebook tab: {pg.url}")
                    break

        # 3. Find any Instagram tab (not DMs)
        if not target_page:
            for pg in context.pages:
                u = pg.url.lower()
                if "instagram.com/" in u and "/direct" not in u:
                    target_page = pg
                    print(f"Found Instagram tab: {pg.url}")
                    break

        if not target_page:
            print("No suitable Facebook/Instagram tab found")
            return [], ""

        try:
            target_page.bring_to_front()
            target_page.wait_for_timeout(600)

            # ── Extract text FIRST — fast, no scrolling needed ────────────────
            text = ""
            try:
                text = target_page.evaluate("""() => {
                    return document.body ? document.body.innerText : '';
                }""")
                text = re.sub(r"\n{3,}", "\n\n", text).strip()
                text = re.sub(r"(?im)^\s*[\d,\.]+\s*(followers?|likes?|reviews?|ratings?|check.?ins?|people follow|people like).*$", "", text)
                text = re.sub(r"\n{3,}", "\n\n", text).strip()[:12000]
                print(f"Text extracted: {len(text)} chars")
            except Exception as e:
                print(f"Text extraction skipped: {e}")

            # ── Always screenshot the first post area — vision sees what text misses ──
            screenshots = []
            try:
                _scroll_to_freshest_post(target_page)
                target_page.wait_for_timeout(1400)   # let images render
                shot = target_page.screenshot(type="jpeg", quality=72)
                screenshots.append((shot, "image/jpeg"))
                print(f"First-post screenshot: {len(shot)} bytes")
                target_page.evaluate("window.scrollTo(0, 0)")
            except Exception as _se:
                print(f"Screenshot failed: {_se}")

            return screenshots, text

        except Exception as e:
            print(f"Capture error: {e}")
            return [], ""

def _capture_and_keep_open(hint_url: str = "") -> tuple:
    """Same scan as capture_profile_tab, but keeps the Playwright/CDP connection
    alive afterwards (via .start()/.stop() instead of a `with` block) so the
    caller can reuse the SAME page later to click Message, instead of opening a
    second connection. /generate used to call capture_profile_tab() to scan and
    then _open_current_messenger() to click, i.e. two full Playwright driver
    launches per Generate click — that doubled the browser-automation overhead
    and was the actual cause of "it's slow again" once the auto-click feature
    was added, separate from the earlier AI-retry-loop slowdowns.
    Returns (pw, page, screenshots, text). pw and page are None if no tab was
    found or Chrome isn't reachable — caller should skip the later click in
    that case. Caller is responsible for calling pw.stop() when done."""
    import socket
    from playwright.sync_api import sync_playwright

    try:
        s = socket.create_connection(("127.0.0.1", CDP_PORT), timeout=3)
        s.close()
    except Exception:
        return None, None, [], ""

    pw = sync_playwright().start()
    try:
        browser = pw.chromium.connect_over_cdp(f"http://localhost:{CDP_PORT}")
        context = browser.contexts[0]

        target_page = None
        if hint_url:
            hint_clean = hint_url.rstrip("/").split("?")[0].lower()
            for pg in context.pages:
                if hint_clean in pg.url.lower():
                    target_page = pg
                    break
        if not target_page:
            for pg in context.pages:
                u = pg.url.lower()
                if "facebook.com/" in u and "messenger.com" not in u and "/messages" not in u:
                    target_page = pg
                    break
        if not target_page:
            for pg in context.pages:
                u = pg.url.lower()
                if "instagram.com/" in u and "/direct" not in u:
                    target_page = pg
                    break

        if not target_page:
            pw.stop()
            return None, None, [], ""

        target_page.bring_to_front()
        target_page.wait_for_timeout(600)

        text = ""
        try:
            text = target_page.evaluate("""() => {
                return document.body ? document.body.innerText : '';
            }""")
            text = re.sub(r"\n{3,}", "\n\n", text).strip()
            text = re.sub(r"(?im)^\s*[\d,\.]+\s*(followers?|likes?|reviews?|ratings?|check.?ins?|people follow|people like).*$", "", text)
            text = re.sub(r"\n{3,}", "\n\n", text).strip()[:12000]
        except Exception as e:
            print(f"Text extraction skipped: {e}")

        screenshots = []
        try:
            _scroll_to_freshest_post(target_page)
            target_page.wait_for_timeout(1400)
            shot = target_page.screenshot(type="jpeg", quality=72)
            screenshots.append((shot, "image/jpeg"))
            target_page.evaluate("window.scrollTo(0, 0)")
        except Exception as _se:
            print(f"Screenshot failed: {_se}")

        return pw, target_page, screenshots, text

    except Exception as e:
        print(f"CDP connect error: {e}")
        try:
            pw.stop()
        except Exception:
            pass
        return None, None, [], ""

# ── Name extractor ────────────────────────────────────────────────────────────
# Common UK / English first names. A detected name is only used if it appears
# here — so we greet "Hi, <name>" only when a real first name is genuinely in
# the company name or bio, and never guess a non-name word ("Designer",
# "Newcastle", "Bathroom", etc.). Lowercase for case-insensitive matching.
FIRST_NAMES = {
    # ── male ──────────────────────────────────────────────────────────────────
    "aaron","adam","adrian","aidan","alan","albert","alex","alexander","alfie","alfred",
    "ali","allan","andrew","andy","angus","anthony","antony","archie","arthur","ashley",
    "barry","ben","benjamin","bernard","bill","billy","bobby","brad","bradley","brandon",
    "brendan","brett","brian","bruce","bryan","callum","cameron","carl","charles","charlie",
    "chris","christian","christopher","clive","colin","connor","conor","craig","curtis","dale",
    "damian","damien","daniel","danny","darren","darryl","dave","david","dean","declan",
    "dennis","derek","dom","dominic","donald","douglas","duncan","dylan","eddie","edward",
    "elliot","elliott","eric","ethan","ewan","felix","finlay","finn","francis","frank",
    "frankie","fred","freddie","frederick","gareth","garry","gary","gavin","geoff","geoffrey",
    "george","gerald","gerard","glen","glenn","gordon","graeme","graham","grant","greg",
    "gregory","guy","harold","harry","harvey","henry","howard","hugh","hugo","ian",
    "isaac","ivan","jack","jacob","jake","james","jamie","jared","jason","jay",
    "jeff","jeffrey","jeremy","jim","jimmy","joe","joel","john","johnny","jon",
    "jonathan","jonny","jordan","joseph","josh","joshua","julian","justin","karl","keith",
    "ken","kenneth","kevin","kieran","kyle","lance","larry","lee","leon","leonard",
    "leslie","lewis","liam","lloyd","logan","louis","luca","lucas","luke","malcolm",
    "marc","marcus","mark","martin","mason","matt","matthew","maurice","max","maxwell",
    "micah","michael","mick","mike","miles","mitchell","morgan","nathan","nathaniel","neil",
    "nick","nicholas","nigel","noah","norman","oliver","ollie","oscar","owen","pablo",
    "patrick","paul","percy","perry","pete","peter","philip","phillip","ralph","ray",
    "raymond","reece","reg","reginald","rhys","richard","rick","ricky","rob","robbie",
    "robert","robin","rod","rodney","roger","ronald","ronnie","rory","ross","roy",
    "russell","ryan","sam","samuel","scott","sean","seb","sebastian","shane","shaun",
    "sidney","simon","stanley","stephen","steve","steven","stewart","stuart","ted","terence",
    "terry","theo","thomas","tim","timothy","toby","todd","tom","tommy","tony",
    "travis","trevor","tristan","troy","tyler","vernon","victor","vincent","wade","walter",
    "warren","wayne","wesley","will","william","willie","zac","zach",
    # ── female ────────────────────────────────────────────────────────────────
    "abbie","abby","abigail","adele","aimee","alana","alexandra","alexis","alice","alicia",
    "alison","amanda","amber","amelia","amy","andrea","angela","angie","anita","ann",
    "anna","anne","annette","annie","april","audrey","barbara","becky","bella","beth",
    "bethan","bethany","betty","beverley","bonnie","brenda","bridget","brooke","carla","carly",
    "carol","caroline","carolyn","carrie","casey","catherine","cathy","charlene","charlotte","chelsea",
    "cheryl","chloe","christina","christine","claire","clara","clare","colleen","connie","courtney",
    "daisy","dana","danielle","dawn","deborah","debbie","debra","denise","diana","diane",
    "donna","doreen","doris","dorothy","ebony","eden","eileen","elaine","eleanor","elena",
    "eliza","elizabeth","ella","ellen","ellie","eloise","elsie","emily","emma","erica",
    "erin","esme","esther","eva","eve","evelyn","faith","faye","felicity","fern",
    "fiona","florence","frances","francesca","freya","gabrielle","gail","gemma","georgia","georgina",
    "gillian","gina","gloria","grace","gwen","hannah","harriet","hayley","hazel","heather",
    "heidi","helen","helena","holly","hope","imogen","irene","iris","isabel","isabella",
    "isabelle","isla","ivy","jacqueline","jade","jane","janet","janice","jasmine","jean",
    "jenna","jennifer","jenny","jessica","jill","joan","joanna","joanne","jodie","josephine",
    "joy","joyce","judith","judy","julia","julie","june","karen","kate","katherine",
    "kathleen","kathryn","kathy","katie","katrina","kay","kayleigh","keeley","kelly","kerry",
    "kim","kimberley","kirsten","kirsty","lacey","lara","laura","lauren","leah","leanne",
    "lesley","libby","lillian","lily","linda","lindsay","lindsey","lisa","lois","lola",
    "lorna","lorraine","louise","lucy","lydia","lynda","lynn","lynne","mabel","madeleine",
    "maggie","mandy","margaret","maria","marie","marilyn","marion","marjorie","marlene","martha",
    "mary","matilda","maureen","maya","megan","melanie","melissa","mia","michelle","millie",
    "miranda","miriam","molly","monica","nadia","nancy","naomi","natalie","natasha","nicola",
    "nicole","nina","nora","norma","olive","olivia","paige","pamela","patricia","paula",
    "pauline","pearl","peggy","penelope","penny","phoebe","phyllis","polly","poppy","rachael",
    "rachel","rebecca","renee","rhona","rita","roberta","robyn","rosa","rose","rosemary",
    "rosie","roxanne","ruby","ruth","sabrina","sadie","sally","samantha","sandra","sara",
    "sarah","sasha","scarlett","shannon","sharon","sheila","shelley","shirley","sian","sienna",
    "simone","sofia","sonia","sophia","sophie","stacey","stella","stephanie","sue","susan",
    "suzanne","sylvia","tamara","tania","tanya","tara","teresa","tessa","theresa","tiffany",
    "tina","tracey","tracy","valerie","vanessa","vera","veronica","vicky","victoria","violet",
    "virginia","vivienne","wendy","whitney","willow","yvonne","zara","zoe",
}


def _name_from_company(company_name: str) -> str:
    """A person's first name inside a company name, or ''.

    Whitelist-only (FIRST_NAMES): we use a token from the company name as the
    greeting name only if it's a known first name. Structural 'Firstname
    Surname' guessing was tried and removed — it mistook brand words like
    'Plumbergy', 'Westcountry' and 'Queen' for names. If a real name is missed,
    add it to FIRST_NAMES.
    """
    for t in re.findall(r"[A-Za-z]+", company_name or ""):
        if t.lower() in FIRST_NAMES:
            return t.capitalize()
    return ""


def extract_first_name(text: str, company_name: str = "") -> str:
    """
    Return the owner's first name ONLY if it's genuinely present in the company
    name or the bio. Never guesses a trade/place/generic word. Returns "" when
    no real first name is found → caller falls back to a plain "Hi.".
    """
    if CUSTOM_NAME_EXTRACTOR:
        return CUSTOM_NAME_EXTRACTOR(text)

    # 1) A first name inside the COMPANY NAME (whitelist, then Firstname-Surname)
    name = _name_from_company(company_name)
    if name:
        return name

    # 2) An explicit owner line / self-intro in the BIO — captured name must be
    #    a known first name, otherwise we ignore it (no guessing).
    bio_patterns = [
        r"\b(?:Owner|Founder|Co[\s\-]?founder|Director|Proprietor|Manager|Run by|MD)\b"
        r"[^\n]{0,15}?\b([A-Z][a-z]{1,14})\b",
        r"\b(?:I['’]?m|I am|my name is|name['’]s)\s+([A-Z][a-z]{1,14})\b",
    ]
    for pat in bio_patterns:
        m = re.search(pat, text or "")
        if m and m.group(1).lower() in FIRST_NAMES:
            return m.group(1).capitalize()

    return ""

# ── Location extractor ───────────────────────────────────────────────────────
UK_PLACES = [
    # Major cities
    "London","Manchester","Birmingham","Leeds","Liverpool","Sheffield","Bristol",
    "Edinburgh","Glasgow","Cardiff","Belfast","Newcastle","Nottingham","Leicester",
    "Coventry","Bradford","Stoke","Wolverhampton","Derby","Southampton","Portsmouth",
    "Brighton","Oxford","Cambridge","Norwich","Exeter","Plymouth","Swindon",
    "Gloucester","Cheltenham","Bournemouth","Peterborough","Luton","Milton Keynes",
    "Reading","Slough","Ipswich","Middlesbrough","Sunderland","Bolton","Blackpool",
    "Blackburn","Oldham","Rochdale","Stockport","Huddersfield","York","Lincoln",
    "Swansea","Newport","Dundee","Aberdeen","Inverness","Stirling","Perth",
    # Counties / regions
    "Devon","Cornwall","Somerset","Dorset","Wiltshire","Hampshire","Kent","Surrey",
    "Sussex","Essex","Suffolk","Norfolk","Hertfordshire","Bedfordshire","Berkshire",
    "Buckinghamshire","Oxfordshire","Gloucestershire","Worcestershire","Warwickshire",
    "Northamptonshire","Cambridgeshire","Leicestershire","Nottinghamshire","Derbyshire",
    "Staffordshire","Shropshire","Cheshire","Lancashire","Yorkshire","Cumbria",
    "Durham","Northumberland","Herefordshire","Lincolnshire","Rutland",
    "East Yorkshire","West Yorkshire","South Yorkshire","North Yorkshire",
    "West Midlands","East Midlands","East Anglia",
    # Common towns
    "Guildford","Woking","Crawley","Horsham","Tunbridge Wells","Maidstone","Canterbury",
    "Rochester","Basildon","Chelmsford","Colchester","Southend","St Albans","Watford",
    "Hemel Hempstead","Stevenage","Harlow","Aylesbury","High Wycombe","Bracknell",
    "Basingstoke","Winchester","Eastbourne","Worthing","Hastings","Folkestone",
    "Bath","Taunton","Yeovil","Poole","Weymouth","Salisbury","Chippenham","Swindon",
    "Torquay","Paignton","Barnstaple","Truro","Newquay","Penzance","Falmouth",
    "Hereford","Shrewsbury","Telford","Worcester","Redditch","Kidderminster",
    "Leamington Spa","Rugby","Nuneaton","Tamworth","Lichfield","Burton upon Trent",
    "Stafford","Walsall","West Bromwich","Dudley","Solihull","Sutton Coldfield",
    "Crewe","Chester","Warrington","Wigan","Preston","Burnley","Lancaster","Carlisle",
    "Harrogate","Wakefield","Barnsley","Doncaster","Rotherham","Chesterfield",
    "Mansfield","Newark","Grimsby","Scunthorpe","Hull","Scarborough","Whitby",
    "Darlington","Durham","Stockton","Gateshead","South Shields","Hartlepool",
    "Berwick","Hexham","Morpeth",
]

UK_NEIGHBOURS = {
    # Counties
    "Essex":             ["Suffolk", "Kent", "Hertfordshire", "Cambridgeshire"],
    "Kent":              ["Surrey", "East Sussex", "West Sussex", "Essex"],
    "Surrey":            ["Kent", "East Sussex", "West Sussex", "Hampshire", "Berkshire"],
    "East Sussex":       ["Kent", "Surrey", "West Sussex"],
    "West Sussex":       ["Hampshire", "Surrey", "East Sussex"],
    "Sussex":            ["Kent", "Surrey", "Hampshire"],
    "Devon":             ["Cornwall", "Somerset", "Dorset"],
    "Cornwall":          ["Devon", "Somerset"],
    "Somerset":          ["Devon", "Dorset", "Wiltshire", "Gloucestershire"],
    "Dorset":            ["Somerset", "Wiltshire", "Hampshire", "Devon"],
    "Hampshire":         ["Dorset", "Wiltshire", "Berkshire", "Surrey", "West Sussex"],
    "Wiltshire":         ["Somerset", "Dorset", "Hampshire", "Berkshire", "Gloucestershire", "Oxfordshire"],
    "Berkshire":         ["Wiltshire", "Hampshire", "Surrey", "Buckinghamshire", "Oxfordshire"],
    "Oxfordshire":       ["Berkshire", "Gloucestershire", "Warwickshire", "Northamptonshire", "Buckinghamshire"],
    "Buckinghamshire":   ["Oxfordshire", "Berkshire", "Hertfordshire", "Northamptonshire"],
    "Hertfordshire":     ["Essex", "Buckinghamshire", "Bedfordshire", "Cambridgeshire"],
    "Bedfordshire":      ["Hertfordshire", "Buckinghamshire", "Northamptonshire", "Cambridgeshire"],
    "Cambridgeshire":    ["Hertfordshire", "Bedfordshire", "Northamptonshire", "Lincolnshire", "Norfolk", "Suffolk"],
    "Suffolk":           ["Norfolk", "Cambridgeshire", "Essex"],
    "Norfolk":           ["Suffolk", "Cambridgeshire", "Lincolnshire"],
    "Lincolnshire":      ["Norfolk", "Cambridgeshire", "Leicestershire", "Nottinghamshire", "Yorkshire"],
    "Northamptonshire":  ["Oxfordshire", "Buckinghamshire", "Bedfordshire", "Cambridgeshire", "Leicestershire", "Warwickshire"],
    "Leicestershire":    ["Northamptonshire", "Lincolnshire", "Nottinghamshire", "Derbyshire", "Warwickshire"],
    "Nottinghamshire":   ["Lincolnshire", "Leicestershire", "Derbyshire", "Yorkshire"],
    "Derbyshire":        ["Nottinghamshire", "Leicestershire", "Staffordshire", "Cheshire", "Yorkshire"],
    "Staffordshire":     ["Derbyshire", "Leicestershire", "Warwickshire", "Worcestershire", "Shropshire", "Cheshire"],
    "Warwickshire":      ["Northamptonshire", "Leicestershire", "Staffordshire", "Worcestershire", "Oxfordshire", "Gloucestershire"],
    "Worcestershire":    ["Warwickshire", "Staffordshire", "Shropshire", "Herefordshire", "Gloucestershire"],
    "Herefordshire":     ["Worcestershire", "Shropshire", "Gloucestershire"],
    "Shropshire":        ["Staffordshire", "Worcestershire", "Herefordshire", "Cheshire"],
    "Gloucestershire":   ["Herefordshire", "Worcestershire", "Warwickshire", "Oxfordshire", "Wiltshire", "Somerset"],
    "Cheshire":          ["Staffordshire", "Derbyshire", "Lancashire", "Shropshire"],
    "Lancashire":        ["Cheshire", "Yorkshire", "Cumbria"],
    "North Yorkshire":   ["Lancashire", "Durham", "Cumbria", "East Yorkshire", "West Yorkshire"],
    "West Yorkshire":    ["Lancashire", "Derbyshire", "South Yorkshire", "North Yorkshire"],
    "South Yorkshire":   ["Nottinghamshire", "Derbyshire", "West Yorkshire", "Lincolnshire"],
    "East Yorkshire":    ["Lincolnshire", "North Yorkshire", "South Yorkshire"],
    "Yorkshire":         ["Lancashire", "Lincolnshire", "Nottinghamshire", "Derbyshire", "Durham"],
    "Cumbria":           ["Lancashire", "Yorkshire", "Durham", "Northumberland"],
    "Durham":            ["Cumbria", "Yorkshire", "Northumberland"],
    "Northumberland":    ["Durham", "Cumbria"],
    # Cities / large towns
    "London":            ["Essex", "Kent", "Surrey", "Hertfordshire"],
    "Manchester":        ["Cheshire", "Lancashire", "Derbyshire"],
    "Birmingham":        ["Warwickshire", "Staffordshire", "Worcestershire"],
    "Leeds":             ["Lancashire", "Harrogate", "Bradford", "Wakefield"],
    "Liverpool":         ["Lancashire", "Cheshire"],
    "Sheffield":         ["Derbyshire", "Nottinghamshire", "Barnsley"],
    "Bristol":           ["Somerset", "Gloucestershire", "Wiltshire"],
    "Newcastle":         ["Durham", "Northumberland", "Sunderland"],
    "Sunderland":        ["Durham", "Northumberland", "Newcastle"],
    "Norwich":           ["Norfolk", "Suffolk", "Cambridgeshire"],
    "Exeter":            ["Somerset", "Cornwall", "Dorset"],
    "Plymouth":          ["Cornwall", "Somerset"],
    "Brighton":          ["East Sussex", "West Sussex", "Surrey"],
    "Southampton":       ["Hampshire", "Dorset", "Wiltshire"],
    "Portsmouth":        ["Hampshire", "West Sussex", "Surrey"],
    "York":              ["North Yorkshire", "East Yorkshire", "West Yorkshire"],
    "Hull":              ["East Yorkshire", "Lincolnshire"],
    "Nottingham":        ["Derbyshire", "Leicestershire", "Lincolnshire"],
    "Leicester":         ["Northamptonshire", "Warwickshire", "Lincolnshire"],
    "Derby":             ["Nottinghamshire", "Staffordshire", "Leicestershire"],
    "Oxford":            ["Berkshire", "Buckinghamshire", "Gloucestershire"],
    "Cambridge":         ["Suffolk", "Hertfordshire", "Bedfordshire"],
    "Coventry":          ["Warwickshire", "Northamptonshire", "Staffordshire"],
    "Stoke":             ["Cheshire", "Derbyshire", "Shropshire"],
    "Wolverhampton":     ["Staffordshire", "Worcestershire", "Shropshire"],
    "Reading":           ["Berkshire", "Oxfordshire", "Hampshire"],
    "Gloucester":        ["Worcestershire", "Herefordshire", "Wiltshire"],
    "Cheltenham":        ["Worcestershire", "Oxfordshire", "Wiltshire"],
    "Bournemouth":       ["Dorset", "Somerset", "Hampshire"],
    "Poole":             ["Dorset", "Somerset", "Hampshire"],
    "Bath":              ["Wiltshire", "Gloucestershire", "Somerset"],
    "Taunton":           ["Devon", "Dorset", "Wiltshire"],
    "Torquay":           ["Cornwall", "Somerset"],
    "Barnstaple":        ["Cornwall", "Somerset"],
    "Preston":           ["Lancashire", "Cheshire", "Yorkshire"],
    "Blackpool":         ["Lancashire", "Cumbria"],
    "Blackburn":         ["Lancashire", "Yorkshire", "Cheshire"],
    "Bolton":            ["Lancashire", "Cheshire"],
    "Chester":           ["Cheshire", "Shropshire", "Lancashire"],
    "Warrington":        ["Cheshire", "Lancashire"],
    "Shrewsbury":        ["Staffordshire", "Cheshire", "Herefordshire"],
    "Telford":           ["Staffordshire", "Worcestershire", "Cheshire"],
    "Worcester":         ["Herefordshire", "Warwickshire", "Gloucestershire"],
    "Hereford":          ["Worcestershire", "Shropshire", "Gloucestershire"],
    "Harrogate":         ["West Yorkshire", "Lancashire", "Durham"],
    "Wakefield":         ["West Yorkshire", "South Yorkshire", "Lancashire"],
    "Huddersfield":      ["West Yorkshire", "Lancashire", "Derbyshire"],
    "Doncaster":         ["Lincolnshire", "Nottinghamshire", "West Yorkshire"],
    "Barnsley":          ["West Yorkshire", "Derbyshire", "Nottinghamshire"],
    "Scarborough":       ["North Yorkshire", "East Yorkshire"],
    "Middlesbrough":     ["Durham", "North Yorkshire"],
    "Darlington":        ["Durham", "North Yorkshire", "Cumbria"],
    "Gateshead":         ["Durham", "Northumberland"],
    "Carlisle":          ["Cumbria", "Northumberland", "Durham"],
    "Colchester":        ["Suffolk", "Hertfordshire", "Cambridgeshire"],
    "Chelmsford":        ["Suffolk", "Hertfordshire", "Kent"],
    "Ipswich":           ["Norfolk", "Cambridgeshire", "Essex"],
    "Guildford":         ["Kent", "Hampshire", "Berkshire"],
    "Maidstone":         ["Surrey", "East Sussex", "Essex"],
    "Canterbury":        ["East Sussex", "Surrey", "Essex"],
    "Hastings":          ["Kent", "Surrey"],
    "Eastbourne":        ["Kent", "Surrey", "West Sussex"],
    "Tunbridge Wells":   ["East Sussex", "Surrey", "Essex"],
    "Truro":             ["Devon", "Somerset"],
    "Newquay":           ["Devon", "Somerset"],
    "Penzance":          ["Devon"],
    "Falmouth":          ["Devon", "Somerset"],
    "Basingstoke":       ["Surrey", "Berkshire", "Wiltshire"],
    "Winchester":        ["Hampshire", "Berkshire", "Wiltshire"],
    # Scotland
    "Edinburgh":         ["Glasgow", "Stirling", "Fife", "Livingston"],
    "Glasgow":           ["Edinburgh", "Stirling", "Paisley", "Hamilton"],
    "Dundee":            ["Perth", "Fife", "St Andrews", "Angus"],
    "Aberdeen":          ["Dundee", "Inverness", "Elgin"],
    "Inverness":         ["Aberdeen", "Elgin", "Aviemore"],
    "Stirling":          ["Edinburgh", "Glasgow", "Perth", "Falkirk"],
    "Perth":             ["Dundee", "Stirling", "Fife", "Crieff"],
    "Falkirk":           ["Edinburgh", "Glasgow", "Stirling"],
    "Livingston":        ["Edinburgh", "Falkirk", "Stirling"],
    "Paisley":           ["Glasgow", "Ayr", "Kilmarnock"],
    "Hamilton":          ["Glasgow", "Motherwell", "Lanark"],
    "Fife":              ["Edinburgh", "Dundee", "Perth", "St Andrews"],
    "St Andrews":        ["Dundee", "Fife", "Perth"],
    "Ayr":               ["Glasgow", "Paisley", "Kilmarnock"],
    "Kilmarnock":        ["Ayr", "Glasgow", "Paisley"],
    "Scotland":          ["Glasgow", "Stirling", "Perth", "Fife"],
    # Wales
    "Cardiff":           ["Newport", "Swansea", "Bristol"],
    "Swansea":           ["Cardiff", "Newport", "Carmarthen"],
    "Newport":           ["Cardiff", "Bristol", "Herefordshire"],
    "Wrexham":           ["Chester", "Shrewsbury", "Flintshire"],
    "Bangor":            ["Caernarfon", "Conwy", "Anglesey"],
    "Carmarthen":        ["Swansea", "Pembroke", "Ceredigion"],
    "Wales":             ["Cardiff", "Swansea", "Newport"],
    # Northern Ireland
    "Belfast":           ["Lisburn", "Bangor", "Newtownabbey"],
    "Derry":             ["Coleraine", "Limavady", "Strabane"],
    "Northern Ireland":  ["Belfast", "Lisburn", "Bangor"],
}

# Map specific towns to their county for neighbour lookup
TOWN_TO_COUNTY = {
    "Poulton le Fylde": "Lancashire", "Poulton-le-Fylde": "Lancashire",
    "Lytham": "Lancashire", "Lytham St Annes": "Lancashire",
    "Fleetwood": "Lancashire", "Cleveleys": "Lancashire",
    "Southend": "Essex", "Basildon": "Essex", "Harlow": "Essex",
    "Woking": "Surrey", "Crawley": "West Sussex", "Horsham": "West Sussex",
    "Worthing": "West Sussex",
    "Rochester": "Kent", "Folkestone": "Kent",
    "Paignton": "Devon", "Bideford": "Devon",
    "Yeovil": "Somerset",
    "Weymouth": "Dorset",
    "Salisbury": "Wiltshire", "Chippenham": "Wiltshire",
    "Aylesbury": "Buckinghamshire", "High Wycombe": "Buckinghamshire",
    "Bracknell": "Berkshire", "Slough": "Berkshire",
    "Luton": "Bedfordshire", "St Albans": "Hertfordshire",
    "Watford": "Hertfordshire", "Hemel Hempstead": "Hertfordshire",
    "Stevenage": "Hertfordshire",
    "Peterborough": "Cambridgeshire",
    "Grimsby": "Lincolnshire", "Scunthorpe": "Lincolnshire", "Lincoln": "Lincolnshire",
    "Wigan": "Lancashire", "Burnley": "Lancashire", "Lancaster": "Lancashire",
    "Crewe": "Cheshire", "Stockport": "Cheshire",
    "Oldham": "Lancashire", "Rochdale": "Lancashire",
    "Bradford": "West Yorkshire", "Halifax": "West Yorkshire",
    "Rotherham": "South Yorkshire", "Chesterfield": "Derbyshire",
    "Mansfield": "Nottinghamshire", "Newark": "Nottinghamshire",
    "Tamworth": "Staffordshire", "Stafford": "Staffordshire",
    "Lichfield": "Staffordshire", "Burton upon Trent": "Staffordshire",
    "Walsall": "Staffordshire", "West Bromwich": "West Midlands",
    "Dudley": "West Midlands", "Solihull": "Warwickshire",
    "Sutton Coldfield": "Warwickshire",
    "Redditch": "Worcestershire", "Kidderminster": "Worcestershire",
    "Leamington Spa": "Warwickshire", "Rugby": "Warwickshire", "Nuneaton": "Warwickshire",
    "Hexham": "Northumberland", "Berwick": "Northumberland", "Morpeth": "Northumberland",
    "Hartlepool": "Durham", "Stockton": "Durham", "South Shields": "Durham",
    "Whitby": "North Yorkshire",
    "Milton Keynes": "Buckinghamshire",
    "Swindon": "Wiltshire",
    "Gloucester": "Gloucestershire",
    "Swansea": "Swansea", "Newport": "Newport", "Cardiff": "Cardiff",
    "Dundee": "Dundee", "Aberdeen": "Aberdeen", "Inverness": "Inverness",
    "Stirling": "Stirling", "Perth": "Perth", "Edinburgh": "Edinburgh",
    "Glasgow": "Glasgow", "Belfast": "Belfast",
    "Falkirk": "Falkirk", "Livingston": "Livingston", "Paisley": "Paisley",
    "Hamilton": "Hamilton", "Ayr": "Ayr", "Kilmarnock": "Kilmarnock",
    "St Andrews": "St Andrews", "Wrexham": "Wrexham", "Carmarthen": "Carmarthen",
}


def extract_prospect_location(opener: str, text: str) -> str:
    """
    Extract the prospect's own location from the opener or page text.
    """
    # 0. The Excel "LISTED CITY" column, when present, is authoritative — it's
    # real data for this specific lead, not a guess. Trust it before any
    # regex/gazetteer scan. This also sidesteps the UK_PLACES substring bug
    # below (e.g. "Bath" false-matching inside "Bathroom Renovation").
    listed_city = _td_section(text, "LISTED CITY")
    if listed_city:
        return listed_city

    # 1. From the opener — "in Bristol", "up in Leeds", "over in Devon"
    m = re.search(
        r'\b(?:in|up in|down in|over in)\s+([A-Z][a-zA-Z ]{2,30}?)(?:\s*[,.]|\s*$|\s+last\b|\s+the\s+other)',
        opener
    )
    if m:
        loc = m.group(1).strip().rstrip(",. ")
        words = loc.split()
        if 1 <= len(words) <= 3 and words[0][0].isupper():
            return loc

    # 2. Page text — "based in X", "serving X", "covering X"
    m = re.search(
        r'(?:based in|serving|located in|covering|working in|work across|operating in)'
        r'\s+([A-Z][a-zA-Z ,&]{2,40}?)(?:\s*[\n,]|\.|\band\b)',
        text, re.I
    )
    if m:
        loc = m.group(1).strip().split(",")[0].strip().title()
        # Marketing copy often puts filler words before the real place
        # ("Serving Homeowners In Chelmsford") — only trust this match if its
        # trailing words actually name a known UK place, and return just that,
        # not the whole captured phrase. Otherwise fall through to the
        # full-text scan below, which only matches real place names.
        words = loc.split()
        for n in (3, 2, 1):
            if len(words) >= n:
                candidate = " ".join(words[-n:])
                for place in UK_PLACES:
                    if candidate.lower() == place.lower():
                        return place

    # 3. Scan full text for any known UK place name
    text_lower = text.lower()
    for place in UK_PLACES:
        if place.lower() in text_lower:
            return place

    return ""


def get_nearby_location(prospect_loc: str) -> str:
    """
    Given the prospect's location, return a different nearby UK place.
    """
    import random

    if not prospect_loc:
        return "the next county"

    # Direct lookup
    neighbours = UK_NEIGHBOURS.get(prospect_loc)
    if neighbours:
        return random.choice(neighbours)

    # Try town → county mapping
    county = TOWN_TO_COUNTY.get(prospect_loc)
    if county:
        neighbours = UK_NEIGHBOURS.get(county)
        if neighbours:
            return random.choice(neighbours)

    # Partial key match (e.g. "North Yorkshire" matches "Yorkshire")
    for key, vals in UK_NEIGHBOURS.items():
        if prospect_loc.lower() in key.lower() or key.lower() in prospect_loc.lower():
            return random.choice(vals)

    # Nothing found — return a placeholder that still reads naturally
    return "the next county"


def extract_location_for_body(opener: str, text: str) -> str:
    """Returns a NEARBY location (not the prospect's own) to use in the DM body."""
    prospect_loc = extract_prospect_location(opener, text)
    nearby = get_nearby_location(prospect_loc)
    print(f"Prospect location: {prospect_loc!r} -> body location: {nearby!r}")
    return nearby


def detect_trade(text: str) -> str:
    """
    Detect the trade type from the page text.
    Returns a plural noun for use in the DM body (e.g. "roofers", "landscapers").
    """
    text_lower = text.lower()

    # ORDER MATTERS on a tie: max() returns the first key it saw at that score,
    # and dicts keep insertion order, so SPECIFIC trades must come before the
    # broad "builders"/"renovation contractors" catch-alls. Otherwise a bathroom
    # firm that says "extension" once scores level and gets called a builder.
    trade_signals = [
        (["bathroom", "wet room", "wetroom", "en-suite", "ensuite", "shower room",
          "bathroom fitting", "bathroom install", "bathroom suite", "walk-in shower",
          "disabled bathroom", "mobility bathroom"], "bathroom fitters"),
        (["kitchen", "kitchen fitting", "kitchen install", "shaker kitchen",
          "fitted kitchen", "worktop", "kitchen refurb", "handleless",
          "quartz worktop", "granite worktop"], "kitchen fitters"),
        (["boiler", "gas safe", "gas engineer", "combi", "worcester", "vaillant",
          "baxi", "ideal boiler", "glow worm", "viessmann", "boiler service",
          "boiler install", "powerflush", "gas cert", "landlord certificate"], "gas engineers"),
        (["air con", "aircon", "air conditioning", "hvac", "daikin", "mitsubishi",
          "fujitsu", "split system", "multi split", "refrigeration", "chiller",
          "f-gas", "fgas", "climate control"], "air con firms"),
        (["heat pump", "air source", "ground source", "ashp", "gshp",
          "underfloor heating", "central heating", "heating engineer",
          "radiator", "heating system"], "heating engineers"),
        (["plumbing", "plumber", "leak repair", "burst pipe", "bathroom plumbing",
          "drainage", "blocked drain", "soil pipe", "water tank"], "plumbers"),
        (["electrician", "electrical", "niceic", "napit", "sparky", "fuse board",
          "consumer unit", "rewire", "pat test", "eicr", "downlight"], "sparkies"),
        (["solar", "photovoltaic", "solar panel", "pv system", "battery storage",
          "givenergy", "solaredge", "inverter", "renewable"], "solar installers"),
        (["roofing", "roofer", "flat roof", "pitched roof", "epdm", "felt roof",
          "fascia", "soffit", "guttering", "lead flashing", "chimney", "velux"], "roofers"),
        (["window", "glazing", "upvc", "double glaz", "conservatory", "bi-fold",
          "bifold", "composite door", "french door", "patio door", "rehau",
          "veka", "aluk"], "window fitters"),
        (["tiling", "tiler", "wall and floor til", "porcelain til", "ceramic til",
          "mosaic", "grout", "tile install"], "tilers"),
        (["plastering", "plasterer", "skim", "rendering", "render", "dry lining",
          "drylining", "artex", "coving"], "plasterers"),
        (["joinery", "joiner", "carpenter", "carpentry", "bespoke furniture",
          "fitted wardrobe", "fitted bedroom", "staircase", "skirting"], "joiners"),
        (["painting", "decorator", "decorating", "wallpaper", "spray painting",
          "emulsion"], "decorators"),
        (["flooring", "luxury vinyl", "lvt", "laminate floor", "engineered wood",
          "floor fitting", "karndean", "amtico"], "flooring firms"),
        (["drone", "aerial", "uav", "uas", "thermal imaging", "aerial survey",
          "aerial photography", "aerial footage", "roof inspection drone"], "drone operators"),
        (["renovation", "refurb", "kitchen reno", "bathroom reno", "full house",
          "house renovation", "property renovation", "interior renovation",
          "renovations", "refurbishment", "home improvement", "full refurb",
          "period property", "barn conversion", "restoration"], "renovation contractors"),
        (["builder", "building contractor", "extension", "loft conversion",
          "groundwork", "brickwork", "blockwork", "underpinning", "footings",
          "new build", "steel frame", "construction", "bricklaying",
          "house extension", "garage conversion"], "builders"),
        (["landscap", "garden", "paving", "driveway", "patio", "turf", "decking",
          "fencing", "artificial grass", "tree work", "hedging", "planting",
          "block pav", "resin", "sandstone"], "landscapers"),
    ]

    scores = {}
    for keywords, trade in trade_signals:
        score = sum(1 for kw in keywords if kw in text_lower)
        if score > 0:
            scores[trade] = score

    if scores:
        best = max(scores, key=scores.get)
        print(f"Trade detected: {best} (scores: {scores})")
        return best

    return "tradespeople"


def detect_landscaper_type(text: str) -> str:
    """
    Sub-detect within landscapers: driveway, patio, or general garden.
    Used to pick the right 'my mate' body copy.
    """
    t = text.lower()
    driveway_kw = ["driveway", "block pav", "resin bound", "resin driveway",
                   "tarmac", "imprinted concrete", "pattern imprint", "tegula", "sett",
                   "monoblock", "cobble"]
    patio_kw    = ["patio", "porcelain", "indian sandstone", "limestone", "paving slab",
                   "decking", "composite deck", "sleeper", "sandstone", "flags", "flagstone"]

    driveway_score = sum(1 for kw in driveway_kw if kw in t)
    patio_score    = sum(1 for kw in patio_kw    if kw in t)

    if driveway_score > patio_score:
        sub = "driveway"
    elif patio_score > 0:
        sub = "patio"
    else:
        sub = "garden"
    print(f"Landscaper sub-type: {sub} (driveway={driveway_score}, patio={patio_score})")
    return sub


def extract_company_name(text: str) -> str:
    """
    Extract company name from scraped Facebook page text.
    The first meaningful line of page content is usually the page/company name.
    """
    if not text:
        return ""
    working = text
    if "--- PAGE TEXT ---" in working:
        working = working.split("--- PAGE TEXT ---", 1)[1]
    elif "--- TRADE TYPE ---" in working:
        after = working.split("--- TRADE TYPE ---", 1)[1].strip()
        lines = after.split("\n")
        working = "\n".join(lines[1:]) if len(lines) > 1 else ""
    for h in ("--- LISTED CITY ---", "--- USER NOTES ---", "--- DETECTED FIRST NAME ---"):
        if h in working:
            working = working.split(h, 1)[0]
    for line in working.strip().split("\n"):
        line = line.strip()
        if line and len(line) >= 3 and not line.startswith("---") and not line.startswith("http"):
            return line
    return ""


def _td_section(text_data: str, header: str) -> str:
    """Pull the first non-empty line after a --- HEADER --- marker."""
    marker = f"--- {header} ---"
    for line in text_data.split("\n"):
        if line.startswith(marker):
            idx = text_data.find(line) + len(line)
            rest = text_data[idx:].strip().split("\n")
            return rest[0].strip() if rest else ""
    return ""


def _pitch_for(niche: str):
    """(trade_noun, job_noun, example_questions) for the pitch body, tailored
    per niche so it reads bespoke. Matches on KEYWORDS contained in the value so
    it's robust to column phrasing ("HVAC / Air Con", "Gas Boiler", etc.).
    Unknown ("") stays NEUTRAL — never assumes a trade that could be wrong."""
    n = (niche or "").strip().lower()
    if not n:
        return ("trades", "a job", "size of the job, what's involved")
    has = lambda *words: any(w in n for w in words)
    # Order matters: bathroom first (a "Bathroom Renovation" value also contains
    # "renovation"); air con before heating; etc.
    if has("bathroom", "wet room", "ensuite"):
        return ("bathroom fitters", "a new bathroom", "full refit, wet room, retile")
    if has("hvac", "air con", "aircon", "air conditioning", "refrigeration", "cooling"):
        # NOT "air con engineers" — a real prospect replied pointing out not
        # everyone in the trade holds/uses that title. "Expert" doesn't claim
        # a specific qualification/registration the way "engineer" does.
        return ("aircon experts", "a new system", "one room, multi-split, or the whole place")
    if has("boiler", "gas", "heating", "central heat"):
        return ("heating engineers", "a new boiler", "combi swap, system upgrade, or just a repair")
    if has("plumb"):
        return ("plumbers", "a job", "a repair, a new install, or a full job")
    if has("renovation", "refurb", "general build", "builder", "extension", "addition", "remodel"):
        return ("renovation specialists", "a renovation", "how big a job, what's involved")
    if has("kitchen"):
        return ("kitchen fitters", "a new kitchen", "full refit, bespoke units, worktops")
    if has("joinery", "joiner", "carpenter"):
        return ("joiners", "a bespoke piece", "fitted wardrobes, custom furniture, a full fit-out")
    # any other recognised trade (roofing, tiling, electrical, etc.)
    return (f"{n} businesses", "a job", "size of the job, what's involved")


_TRADE_SINGULAR = {
    "bathroom fitters": "bathroom fitter",
    "aircon experts": "aircon expert",
    "heating engineers": "heating engineer",
    "plumbers": "plumber",
    "renovation specialists": "renovation specialist",
    "kitchen fitters": "kitchen fitter",
    "joiners": "joiner",
    "trades": "client",   # unknown-niche fallback — "trade" reads oddly as a person
}


def _singular_trade(trade_noun: str) -> str:
    """Singular form of a _pitch_for() trade_noun, for 'my last ___ got...'
    lines. Known plurals map directly; the '{n} businesses' fallback just
    drops the trailing 's' off 'businesses'."""
    if trade_noun in _TRADE_SINGULAR:
        return _TRADE_SINGULAR[trade_noun]
    if trade_noun.endswith("businesses"):
        return trade_noun[:-len("businesses")] + "business"
    return trade_noun


def pick_body(trade: str, text_data: str, city_hint: str = "", company_name: str = "") -> str:
    """
    English demo-pitch body. Niche-tailored so each message reads bespoke.
    The `trade` arg is kept for signature compatibility but the niche is
    detected here via the name-first + text-scoring classifier.
    """
    # Always prefer the clean Excel name injected into text_data
    # over whatever extract_company_name() scraped from the page
    excel_name = _td_section(text_data, "COMPANY NAME")
    if excel_name:
        company_name = excel_name
    # Niche: prefer the Excel "type / niche" column (injected as TRADE TYPE);
    # fall back to the name/text classifier when that column is empty.
    excel_niche = _td_section(text_data, "TRADE TYPE")
    niche = excel_niche or _en_detect_niche(text_data, company_name)
    trade_noun, job, examples = _pitch_for(niche)

    who = _who_for_pitch(company_name)
    # Capitalise the social-proof line since it starts its own sentence in the
    # intro, so you can type the SOCIAL_PROOF constant in any case.
    sp = (SOCIAL_PROOF[:1].upper() + SOCIAL_PROOF[1:]) if SOCIAL_PROOF else ""

    # Rotate the wording so two leads in the same niche never get a textually
    # identical body (kills Facebook's duplicate-content flag and the "template"
    # read if two owners ever compare notes). The meaning and the honest "student"
    # anchor stay constant, only phrasing moves. Same approach as
    # pick_stripped_message().
    # Intro leads with SOCIAL PROOF (others already trust it) + built-for-you-free.
    # No em dashes anywhere (tidy_message also strips them as a backstop).
    intro = random.choice([
        f"I'm a student studying AI, I build quoting chatbots for {trade_noun}. {sp}, "
        f"and I've set one up for {who} too, free.",
        f"I'm a student getting into AI, I build quote bots for {trade_noun}. {sp}, "
        f"and I've put one together for {who} as well, free.",
        f"I'm a student learning AI, I make quoting chatbots for {trade_noun}. {sp}, "
        f"and I've made one for {who} too, no charge.",
    ])
    # Disarm: your real differentiator an agency can't fake. Answers "why free?".
    disarm = random.choice([
        "No catch, I'm a student not an agency, I just need the real examples more than the money right now.",
        "No catch, I'm a student not an agency, I need the real-world examples more than anything right now.",
    ])
    # Sell the BENEFIT, not the mechanics. Chill, spoken-out-loud tone (contractions,
    # "basically", "gonna"): what's in it for them, not how the bot works.
    # ONE self-contained middle (not a how×benefit cross-product) so the two
    # sentences never echo each other ("every site visit ... every site visit",
    # "wasted trip ... wasted trips"). Lead-QUALITY angle: they already have leads.
    middle = random.choice([
        "It's not about more leads, you've already got those. It just makes sure the ones that reach you are actually worth your time. Less messing about, serious work booked.",
        f"People get an estimate for {job} up front, so the ones who message you are already serious, budget and details already in. You end up with customers who actually fit, not just more messages.",
        f"It filters for fit before anyone reaches you, so the people who get in touch about {job} are ones actually worth taking on, not just more enquiries.",
        f"They get a rough estimate for {job} first, so the only ones who message are keen, informed, and a proper fit. Quality over quantity, basically.",
    ])
    # Take-away close: removing the pressure lifts replies. Easy channel choice.
    close = random.choice([
        "Might not even be your thing, but if you fancy a look I'll send it over, here or WhatsApp?",
        "Might not be for you, but if you want a look I'll send it across, here or WhatsApp?",
    ])
    pitch = f"{intro} {disarm}\n\n{middle}\n\n{close}\n\nCheers"

    return "\n" + pitch


# ── Owner email composer (for leads whose contact is an email) ─────────────────
EMAIL_SIGNOFF = "Milan"

def _email_first_name(owner_name: str) -> str:
    """First name for the email greeting, taken from the 'Owner Name' column.
    Returns '' when the value doesn't look like a person's name, so we fall back
    to a plain 'Hi,' rather than greeting a company by its trading name."""
    o = (owner_name or "").strip()
    if not o:
        return ""
    first = o.split()[0]
    bad = {"the", "a", "gas", "air", "ltd", "uk"}
    if first.isalpha() and len(first) >= 2 and first.lower() not in bad:
        return first[:1].upper() + first[1:]
    return ""

def _email_subject(company_name: str, niche: str = "", first: str = "") -> str:
    """Short, lowercase, personalised subject — the format the open-rate data backs
    (Gong's 85M-email study + Belkins' 5.5M-email 2025 study): all-lowercase reads
    like a colleague not marketing, 1-4 words wins (≤40-50 chars), a first name is
    the single biggest lever (~+22%), question framing adds ~+10%, 1-2 word pattern
    interrupts perform, and ANY selling/hype/AI language drags opens ~18%. So the
    set mixes British-casual ('quick one'), a question ('worth a look?'), a pattern
    interrupt ('for {name}'), an honest curiosity hook tied to the email ('built you
    something'), and a niche-relevant line ('your {thing} quotes'). Rotated so a
    batch isn't identical; falls back to the business name with no first name."""
    who = _clean_company_name(company_name) or "your business"
    # Short niche word for a relevant variant: "a new bathroom" -> "bathroom".
    _, job, _ex = _pitch_for(niche)
    thing = (job or "").replace("a new ", "").replace("a ", "").strip() or "job"
    if first:
        f = first.lower()   # all-lowercase reads like a colleague, not marketing
        return random.choice([
            f"{f}, quick one",            # British-casual, 3 words, no hype
            f"built you something, {f}",  # honest curiosity, ties to the email's hook
            f"for {f}",                   # 2-word pattern interrupt
            f"{f}, worth a look?",        # question framing (~+10% opens)
            f"{f}, your {thing} quotes",  # niche-relevant, still curiosity not pitch
        ])
    return random.choice([
        f"built you something",
        f"quick one for {who}",
        f"{who}, worth a look?",
        f"your {thing} quotes",
    ])

def _clean_fb_title(title: str) -> str:
    """Pull a usable business name out of a Facebook page <title>. FB titles look
    like 'Steve Jones Plumbing | Facebook'. Crucially this must survive a Chrome
    that is logged into Facebook in ANOTHER language: the body text then contains
    UI chrome like 'Az olvasatlan értesítések száma' (Hungarian for 'unread
    notifications'), which the old body-text extractor wrongly used as the company
    name. We take the part before the separator and reject any FB/notification
    UI noise."""
    t = (title or "").strip()
    if not t:
        return ""
    for sep in ("|", "–", "—", " - ", "/", "·", "•"):
        if sep in t:
            t = t.split(sep)[0].strip()
            break
    low = t.lower()
    noise = ("facebook", "messenger", "értesít", "olvasatlan", "notification",
             "unread", "home", "kezdőlap", "log in", "bejelentkez")
    if not t or len(t) > 60 or any(n in low for n in noise):
        return ""
    return t

def _build_email(opener: str, page_text: str, lead: dict, page_title: str = "") -> tuple:
    """Compose the owner email as (subject, body).

    This is a COLD email to a busy tradesman who has plenty of competition, so it
    earns its place: one genuine line proving we looked at their page, then it
    leads with what's in it for THEM (instant prices filter time-wasters; serious
    leads arrive with budget + details; they reply before a competitor does). It
    does NOT explain how the tool works step by step — owners don't care. The
    "I already built one on your branding, want the link?" close makes it
    believable that a real thing exists, and the student framing explains the
    'free'. Segments are rotated so a batch isn't textually identical.

    Business name comes from the page <title> (clean), else 'your business' — we
    never scrape it from body text, which on a non-English FB session is noise."""
    owner = lead.get("name", "")
    city  = lead.get("city", "")
    niche = lead.get("niche", "")

    td = page_text or ""
    if niche:
        td = f"--- TRADE TYPE ---\n{niche}\n\n" + td
    if city:
        td += f"\n\n--- LISTED CITY ---\n{city}"

    company = _clean_fb_title(page_title) or "your business"
    niche_resolved = niche or _en_detect_niche(td, company)
    _, job, _examples = _pitch_for(niche_resolved)   # e.g. "a new bathroom"

    first = _email_first_name(owner)
    greet = f"Hey {first}," if first else "Hey,"

    # One genuine personal line (from their page), kept to a single sentence.
    op = (opener or "").strip()
    if op:
        op = op[:1].upper() + op[1:]
        if not op.endswith((".", "!", "?")):
            op += "."

    # ── TAGOLT layout ───────────────────────────────────────────────────────
    # NOT a paragraph email. Every line below is its own short thought with a
    # blank line around it, so a busy tradesman SCANS it in a couple of seconds
    # and gets it — Charlie-Morgan style. So each piece is a list of one-line
    # blocks, all flattened into `blocks` and joined with blank lines. Keep each
    # line short; never let two clauses share a line.

    # The honest student / case-study framing (disarms "what's the catch", and
    # never the word "free" — a spam-filter trigger). Two short lines.
    # Quiet social proof + mild scarcity: imply others are already on board (true)
    # and we just need one more. Raises credibility without FOMO pressure, which
    # would clash with the genuine student tone. Stays plain and warm.
    frame = random.choice([
        ["I'm a student building instant-quote tools for trades.",
         "Not chasing anything, a couple of trades are using one already, I just want one more solid reference of my work."],
        ["I'm a student making instant-quote tools for trades.",
         "Not after anything, I've got a few businesses using one already, just need one last good reference of my work."],
    ])

    # The "what it is" — ONE short line that names the mechanism concretely (an
    # estimate on THEIR site), niche-aware via {job}. No second clause. "Estimate"
    # not "price" — it's an honest ballpark range, and over-promising a firm
    # "price" then showing a range reads worse.
    what_it_is = random.choice([
        f"It puts an instant estimate for {job} right on your site.",
        f"It gives people an instant estimate for {job} on your page.",
        f"It prices up {job} for people straight off your site.",
    ])

    # The payoff — a short filter line, then the punchline broken across two lines
    # for a dramatic beat (reads as one sentence, lands harder). The build-up line
    # above ("a price on your site") earns the punchline so it isn't vague.
    payoff = random.choice([
        ["The wrong enquiries drop off before they even reach you.",
         "What's left",
         "are the right customers."],
        ["No more pricing up jobs that never land.",
         "Just the customers",
         "actually worth taking on."],
        ["The tyre kickers filter themselves out first.",
         "What gets through",
         "are proper customers."],
    ])

    # Believable proof (a real thing exists, on their branding) + the ask, with a
    # channel choice. Two short lines.
    close = random.choice([
        ["I've built one on your branding already.",
         "Want it here or on WhatsApp?"],
        ["Built one on your branding already.",
         "Shall I send it here or over WhatsApp?"],
    ])

    blocks = [greet]
    if op:
        blocks.append(op)
    blocks += frame
    blocks.append(what_it_is)
    blocks += payoff
    blocks += close
    blocks.append(f"Cheers,\n{EMAIL_SIGNOFF}")
    body = "\n\n".join(blocks)
    return _email_subject(company, niche, first), tidy_message(body)


def _email_clipboard(to: str, subject: str, body: str) -> str:
    """Full email on the clipboard with NO field labels — just the address, then
    the subject, then the body. So one paste gives: address line, subject line,
    blank line, body (not 'To:'/'Subject:' prefixes)."""
    return f"{to}\n{subject}\n\n{body}"


def pick_stripped_message(text_data: str, name_hint: str = "", company_name: str = "") -> str:
    """Direct, honest A/B variant — no personalised opener, no vision call.
    'Hi [name], I'm a student building AI quote bots for [trade] and I've put a
    sample one together for [company]. No pressure, thought I'd just show you
    what it does rather than explain it. Can I send it over?'
    Leads with the honest "I'm a student" framing (disarming, not vendor-y).
    Closer never rotates. Currently the channel-neutral "Can I send it over?"
    (deliver the demo link on Messenger). Switch to "Can I send it here or would
    WhatsApp be better?" once WhatsApp is confirmed working.
    Niche comes from the Excel TRADE TYPE column first, else the name/text
    classifier — the SAME source the full body uses, so it stays niche-accurate."""
    excel_name = _td_section(text_data, "COMPANY NAME")
    if excel_name:
        company_name = excel_name
    excel_niche = _td_section(text_data, "TRADE TYPE")
    niche = excel_niche or _en_detect_niche(text_data, company_name)
    trade_noun, _job, _ex = _pitch_for(niche)

    # Straight in — no hollow pleasantries ("hope you're well"). They don't care
    # about that; brevity and getting to the point is the warmth here.
    greet = f"Hi {name_hint}, " if name_hint else "Hi, "
    who = _clean_company_name(company_name) or "your business"
    if name_hint and who.lower() == name_hint.lower():
        who = "you"   # sole trader: name == company, avoid "Hi Pierce ... for Pierce"

    # Rotate natural wordings of the middle so no two messages are textually
    # identical (kills Facebook's duplicate-content spam flag) while keeping the
    # exact meaning and tone. The "I'm a student" anchor and the "Can I send it
    # over?" closer never rotate out.
    build = random.choice([
        f"I'm a student building AI quote bots for {trade_noun}",
        f"I'm a student getting into AI, building quote bots for {trade_noun}",
        f"I'm a student learning AI, I build quote bots for {trade_noun}",
    ])
    sample = random.choice([
        f"and I've put a sample one together for {who}",
        f"and I've built a sample for {who}",
        f"and I've made a sample one for {who}",
    ])
    show = random.choice([
        "No pressure, thought I'd just show you what it does rather than explain it.",
        "No pressure, figured it's easier to show you than explain it.",
        "No pressure, thought it'd be easier to just show you what it does.",
    ])
    return f"{greet}{build} {sample}. {show} Can I send it over?"


def _clean_dm_location(loc: str) -> str:
    """Tidy a scraped location so it reads like a human typed it mid-sentence.
        'Liverpool & Merseyside Areas' -> 'Liverpool'
        'Leeds and surrounding areas'  -> 'Leeds'
        'Greater Manchester'           -> 'Manchester'
        'Bristol, Bath & Somerset'     -> 'Bristol'
    A real person names ONE town. Returns '' if nothing usable is left."""
    n = (loc or "").strip()
    if not n:
        return ""
    # Cut at the first connector — keep only the lead place.
    for sep in (" & ", " and ", ",", "/", " - "):
        i = n.lower().find(sep.lower())
        if i > 0:
            n = n[:i]
            break
    # Drop scraped filler words.
    n = re.sub(r"\b(areas?|surrounding|greater|the|region|county|all of|covering)\b", "", n, flags=re.I)
    n = re.sub(r"\s{2,}", " ", n).strip(" .,-&")
    return n


def pick_british_message(opener: str, text_data: str, name_hint: str = "", company_name: str = "",
                          brand_color: str = "", market: str = "US") -> str:
    """Student variant — reintroduces "I'm a student learning AI" framing that
    was previously pulled (a real prospect read it as the exact pretext
    overseas outreach agencies use). Brought back per explicit instruction
    after real data showed it drove more replies than the proof-first
    version, on the theory that what actually read as a pretext before was
    STACKING "I'm just a student" with a confident proven-track-record stat
    in the same message, not the student framing itself. This version drops
    the stat/case-study entirely and replaces it with a different, more
    personal proof: OFFERING to build a working sample with their actual
    logo/colour on it (not claiming one is already built — no per-lead demo
    actually gets built automatically at send time, so "already built"
    would be an unverifiable claim, same rule as the stat and the colour).
    Ends on "or are you just too busy for that", a callback to the time-back
    theme as a light dare instead of a flat ask, since "can I send the link
    over" was repeated in every previous version.

    brand_color: the lead's actual extracted brand color, if known. Only used
    when non-empty — never claims a specific color that wasn't actually pulled
    for this lead, since an unverifiable claim is what we're avoiding here.

    market: "US" (default) or "UK" — swaps word-choice only (sorted/minute
    vs set up/right now, and "too busy"/"too slammed" in the close), never
    the substance.
    """
    excel_name = _td_section(text_data, "COMPANY NAME")
    if excel_name:
        company_name = excel_name

    # Name the business directly when it reads naturally as a name ("for AC
    # Breeze"); falls back to "you" for generic/descriptive/URL-like names, in
    # which case we drop the "for ___" clause entirely rather than say "for you".
    who = _who_for_pitch(company_name)
    has_name = who.lower() != "you"
    for_clause = f" for {who}" if has_name else ""

    # Same niche resolution pick_body() uses, so the pitch always matches the
    # trade actually detected for this lead (e.g. "bathroom fitters",
    # "heating engineers"; unrecognised trades fall back to "{niche} businesses").
    niche = _resolve_niche(text_data)
    trade_noun, _job, _ex = _pitch_for(niche)

    op = (opener or "").strip().rstrip(".!,")
    op += "."

    # Rotate the surface wording so 150 messages aren't byte-identical (Facebook's
    # duplicate-content spam flag, same reason pick_body/pick_stripped rotate).
    # Meaning and voice stay fixed per block; only phrasing moves. The user's
    # locked wording is the first option in each block.
    # Student + time-back framing, per explicit instruction. No em dashes
    # anywhere (reads as an AI/spam tell). "Leads" is avoided throughout (a
    # known trigger word for "this is a marketing pitch").
    is_uk = market == "UK"

    # Block 1: restored per explicit instruction, "I'm a student." alone
    # read as too bare once seen live, the AI-studying + time-back line is
    # the actual second paragraph.
    if is_uk:
        student = random.choice([
            f"I'm a student studying AI, and the whole point of what I build is getting busy {trade_noun} their time back.",
            f"I'm a student, studying AI at the minute, and what I build is meant to get busy {trade_noun} their time back.",
            f"Studying AI as a student at the minute, and what I build is meant to get busy {trade_noun} their time back.",
        ])
    else:
        student = random.choice([
            f"I'm a student studying AI, and the whole point of what I build is getting busy {trade_noun} their time back.",
            f"I'm a student, studying AI right now, and what I build is meant to get busy {trade_noun} their time back.",
            f"Studying AI as a student right now, and what I build is meant to get busy {trade_noun} their time back.",
        ])

    # Block 2: the mechanism, unchanged in substance from the proof-first
    # version (quote -> answer -> follow up -> booked in), just without the
    # "not just a one off reply" defensive clause, since "follows up until
    # booked in" already implies more than a one-off reply on its own.
    mechanism = random.choice([
        f"It's a quoting assistant{for_clause}, gives an estimate on the spot, answers people's questions, and follows up until they're booked in for you.",
        f"It's a quoting assistant{for_clause}, gives them an estimate on the spot, answers what they ask, and chases it up until it's booked in for you.",
        f"It's a quoting assistant{for_clause}, gives them an estimate straight away, answers their questions, and follows it through until it's booked in for you.",
    ])

    # Block 3: claims the sample is ALREADY built, per explicit instruction
    # (overrides the earlier "offer to build" default this was written with).
    # brand_color only appears when actually detected off their logo; falls
    # back to plain "logo and branding" otherwise.
    if is_uk:
        if brand_color:
            # Names "{brand_color}" explicitly as their branding colour
            # rather than dropping the bare colour word in ("with your red
            # and branding sorted"), which reads ambiguous, red WHAT.
            sample = random.choice([
                f"I know it's hard to picture, so I've already put together a sample for you with your branding colour, {brand_color}, sorted.",
                f"Hard to picture without seeing it, so I've already built a sample with your branding colour, {brand_color}, sorted.",
                f"Know it's hard to imagine, so I've already knocked up a sample for you with your branding colour, {brand_color}, sorted.",
            ])
        else:
            sample = random.choice([
                "I know it's hard to picture, so I've already put together a sample for you with your logo and branding sorted.",
                "Hard to picture without seeing it, so I've already built a sample with your logo and branding sorted.",
                "Know it's hard to imagine, so I've already knocked up a sample for you with your logo and branding sorted.",
            ])
    else:
        if brand_color:
            sample = random.choice([
                f"I know it's hard to picture, so I've already put together a sample for you with your branding color, {brand_color}, set up.",
                f"Hard to picture without seeing it, so I've already built a sample with your branding color, {brand_color}, set up.",
                f"Know it's hard to imagine, so I've already knocked up a sample for you with your branding color, {brand_color}, set up.",
            ])
        else:
            sample = random.choice([
                "I know it's hard to picture, so I've already put together a sample for you with your logo and branding set up.",
                "Hard to picture without seeing it, so I've already built a sample with your logo and branding set up.",
                "Know it's hard to imagine, so I've already knocked up a sample for you with your logo and branding set up.",
            ])

    # Block 4: lets THEM pick the delivery channel instead of guessing.
    # First contact always stays on Messenger regardless of their answer, so
    # this never counts as cold WhatsApp outreach, if they pick WhatsApp
    # that's now an invited, opted-in switch, not unsolicited contact.
    close = random.choice([
        "Should I send it here or on WhatsApp, whichever's easier for you?",
        "Want it here or on WhatsApp, your call?",
        "Here or WhatsApp, whichever suits, just say the word?",
    ])

    return f"{op}\n\n{student}\n\n{mechanism}\n\n{sample}\n\n{close}"


EXTRACT_PROMPT = """\
Look at this trade business page text carefully. List ONLY details you can actually read. Do not guess.

Output a short bullet list with ONLY these if present:
- Trade type (landscaper / roofer / builder / renovation contractor / drone operator)
- MOST RECENT POST: job type, material, colour/finish, any location in the caption (full place name, no abbreviations or postcodes)
- Other post locations: any town/city names in post captions (full names only)
- Profile location: the About/bio location (for body message use)
- Owner first name (ONLY if in the company name, email, or About — not guessed)
- Anything unusual: massive job, niche material, striking before/after, unusual service

Do not include anything not clearly written in the text. Bullet list only, nothing else."""


_FALLBACK_TEMPLATES = {
    "roofers": [
        "Came across your page over in {loc}, fair bit of roofing work on there.",
        "Clocked your page over in {loc}, looks like you keep busy with the roofing.",
    ],
    "landscapers": [
        "Came across your page over in {loc}, fair bit of landscaping work on there.",
        "Clocked your page over in {loc}, looks like you keep busy with the landscaping.",
    ],
    "builders": [
        "Came across your page over in {loc}, fair bit of building work on there.",
        "Clocked your page over in {loc}, some solid building work on there.",
    ],
    "renovation contractors": [
        "Came across your page over in {loc}, decent renovation work on there.",
    ],
    "drone operators": [
        "Came across your page over in {loc}, decent aerial work on there.",
    ],
}
_FALLBACK_LOC_ONLY = [
    "Came across your page over in {loc}, solid work on there.",
    "Clocked your page over in {loc}, looks like you keep busy.",
]
_FALLBACK_NOTHING = "Came across your page, solid work on there."


def copy_to_clipboard(text: str) -> bool:
    """Copy text to Windows clipboard. Returns True if successful."""
    if not _CLIPBOARD_OK:
        return False
    try:
        _pyperclip.copy(text)
        print("  DM copied to clipboard.")
        return True
    except Exception as e:
        print(f"  Clipboard copy failed: {e}")
        return False


def copy_image_to_clipboard(image_bytes: bytes) -> bool:
    """Copy an image (any format Pillow can read, e.g. the JPEG screenshots we
    already take) to the Windows clipboard as a DIB, so it can just be Ctrl+V'd
    straight into the Messenger chat as an attachment. Separate call from
    copy_to_clipboard() on purpose — text and image can't usefully share the
    clipboard at once in a manual-paste workflow, so this is only called when
    the user explicitly asks for the picture (via /auto_copy_screenshot),
    after they've already pasted the text."""
    try:
        import io
        from PIL import Image
        import win32clipboard
        image = Image.open(io.BytesIO(image_bytes)).convert("RGB")
        output = io.BytesIO()
        image.save(output, "BMP")
        dib = output.getvalue()[14:]   # strip the 14-byte BMP file header — CF_DIB wants just the info header + pixel data
        output.close()
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, dib)
        win32clipboard.CloseClipboard()
        print("  Screenshot copied to clipboard.")
        return True
    except Exception as e:
        print(f"  Image clipboard copy failed: {e}")
        try:
            import win32clipboard
            win32clipboard.CloseClipboard()
        except Exception:
            pass
        return False


def patch_fallback_opener(opener: str, text_data: str, city_hint: str = "") -> str:
    """
    If the model used the generic fallback ("came across your page"), swap it
    for an honest line about the ACTUAL reason: the vision model genuinely
    couldn't find a job photo on the freshest post to react to. This is
    specific and true, unlike a generic "I know you get tons of these"
    acknowledgment — that phrasing was tried and cut earlier for being about
    cold DMs in general rather than this actual interaction, and firing it on
    ~18% of sends would just become its own recognizable fixed phrase (the
    same pattern-match problem this whole rewrite is fighting). Runs AFTER
    call_gemini() — Python safety net.
    """
    # Only touch it if it's the generic fallback
    if "came across your page" not in opener.lower() and "clocked your page" not in opener.lower():
        return opener

    loc = (_td_section(text_data, "LISTED CITY")
           or extract_prospect_location("", text_data)
           or city_hint)
    if loc:
        result = random.choice([
            f"Honestly couldn't find a specific job on your {loc} page to point to, but didn't want that to stop me reaching out.",
            f"Had a look through your {loc} page but couldn't find one specific job to react to — reaching out anyway.",
            f"Couldn't spot a specific post to point to on your {loc} page, but wanted to reach out regardless.",
        ])
    else:
        result = random.choice([
            "Honestly couldn't find a specific job on your page to point to, but didn't want that to stop me reaching out.",
            "Had a look through your page but couldn't find one specific job to react to — reaching out anyway.",
            "Couldn't spot a specific post to point to on your page, but wanted to reach out regardless.",
        ])
    print(f"Fallback normalised -> {result!r}")
    return result


# ── Niche resolution + prompt builder + opener post-validation ───────────────
def _canon_niche(value: str) -> str:
    """Map a free-form niche/column value to the canonical niche key used by the
    opener instructions and the contradiction guard. Keyword-based so it copes
    with 'Bathroom Renovation', 'Gas Boiler', 'HVAC / Air Con', etc."""
    n = (value or "").strip().lower()
    if not n:
        return ""
    has = lambda *w: any(x in n for x in w)
    if has("bathroom", "wet room", "ensuite"):            return "bathroom renovation"
    if has("hvac", "air con", "aircon", "air conditioning", "refrigeration", "cooling"):
        return "air conditioning"
    if has("boiler", "gas", "heating", "central heat"):   return "heating"
    if has("plumb"):                                      return "plumbing"
    if has("window", "glazing", "conservatory", "upvc"):  return "windows and doors"
    if has("roof"):                                       return "roofing"
    if has("solar", "photovolt"):                         return "solar"
    if has("paving", "driveway", "block pav", "resin"):   return "paving"
    if has("tiling", "tiler"):                            return "tiling"
    if has("render", "ewi", "k rend"):                    return "external wall insulation"
    if has("landscap", "garden"):                         return "landscaping"
    if has("decking"):                                    return "decking"
    # Kitchen split out from joinery: a business fitting kitchens identifies
    # as a kitchen fitter, not a joiner, even though en_niche.py's detection
    # vocab still scores them together (kitchen work overlaps heavily with
    # bespoke joinery). Checked first so "Kitchen Fitting" / "Bespoke
    # Kitchens" never falls into the joinery bucket downstream.
    if has("kitchen"):                                    return "kitchen fitting"
    if has("joinery", "joiner", "carpenter"):              return "joinery"
    if has("electric"):                                   return "electrical"
    if has("renovation", "refurb", "general build", "builder", "extension", "addition"):
        return "general building"
    return n


_GENERAL_REMODEL_NAME_SIGNALS = (
    "remodeling", "remodelling", "renovation", "renovations", "home improvement",
    "home improvements", "general contractor", "general contracting", "handyman",
)
# Strong, specific bathroom signal — a name with one of these is trusted as
# bathroom-specific no matter what else it mentions.
_BATHROOM_NAME_SIGNALS = ("bathroom", "wet room", "ensuite", "shower room")
# Other-trade words that, alongside a bare "bath" (not "bathroom"), signal a
# genuinely multi-trade name ("Kitchen-Bath-Flooring") rather than a
# bathroom-specific franchise that just happens to use "Bath" in its name
# ("Re-Bath", "Bath Fitter", "Bath Planet").
_OTHER_TRADE_NAME_SIGNALS = ("kitchen", "flooring") + _GENERAL_REMODEL_NAME_SIGNALS


def _resolve_niche(text_data: str) -> str:
    """
    Resolve the niche the SAME way the body does: the Excel TRADE TYPE column is
    authoritative (so the opener and the pitch never disagree) — EXCEPT when the
    column says "Bathroom Renovation" but the business's own name reads as
    multi-trade or general remodeling rather than bathroom-specific. Two cases:
    (1) no bath-related word in the name at all, but a general-remodel word is
    ("Tim Ormerod Remodeling & Home Improvements"); (2) a bare "bath" (not
    "bathroom") alongside another trade word ("Attleboro Kitchen-Bath-Flooring")
    — a genuine bathroom franchise ("Re-Bath", "Bath Fitter") only ever has
    "bath" alone, never paired with "kitchen"/"flooring"/etc. Lead lists
    compiled from a "bathroom renovation" search sweep in businesses like these,
    so the business's own name is trusted over the list-compiler's category tag.
    A name with "bathroom" spelled out always keeps the column's bathroom tag
    regardless of what else it mentions ("ABC Bathroom Remodeling" stays
    bathroom). Then name-first, then full text-scoring. Result is a canonical
    key.
    """
    excel_niche = _td_section(text_data, "TRADE TYPE")
    company_name = _td_section(text_data, "COMPANY NAME")

    if excel_niche:
        canon = _canon_niche(excel_niche)
        if canon == "bathroom renovation" and company_name:
            name_lower = company_name.lower()
            has_strong_bathroom = any(s in name_lower for s in _BATHROOM_NAME_SIGNALS)
            if not has_strong_bathroom:
                has_general_signal = any(s in name_lower for s in _GENERAL_REMODEL_NAME_SIGNALS)
                has_bath_word = bool(re.search(r"\bbath\b", name_lower))
                has_other_trade = any(s in name_lower for s in _OTHER_TRADE_NAME_SIGNALS)
                if has_general_signal or (has_bath_word and has_other_trade):
                    return "general building"
        return canon

    if company_name:
        niche = _en_detect_niche_from_name(company_name)
        if niche:
            return niche
    return _en_detect_niche(text_data, company_name)


def _en_prompt_builder(full_prompt: str, text_data: str) -> str:
    """Inject per-niche context before the vision call."""
    niche = _resolve_niche(text_data)
    # Unknown ("") still gets the general-builder guidance ("react to whatever
    # trade work you can see") so the vision opener stays useful.
    instr = (_EN_NICHE_OPENER_INSTRUCTIONS.get(niche)
             or _EN_NICHE_OPENER_INSTRUCTIONS.get("general building"))
    if instr:
        print(f"  Prompt builder: niche='{niche or '(unknown)'}' — injecting instructions")
        return full_prompt + f"\n\n{instr}"
    return full_prompt


# Generic trade / legal words — a pair made of ONLY these isn't "the company
# name" (e.g. "bathroom renovation"), so it shouldn't trigger the name guard.
_NAME_GUARD_GENERIC = {
    "the", "and", "co", "ltd", "limited", "llp", "uk", "group", "services",
    "service", "bathroom", "bathrooms", "renovation", "renovations",
    "refurbishment", "refurbishments", "fitter", "fitters", "fitting",
    "installations", "installation", "kitchen", "kitchens", "plumbing",
    "heating", "gas", "boiler", "tiling", "design", "designs", "builders",
    "building", "construction", "company", "air", "con", "conditioning",
    "hvac", "interiors", "interior", "property", "maintenance", "solutions",
    "contractors", "specialist", "specialists",
}


def _opener_repeats_company_name(opener: str, company: str) -> bool:
    """True if the opener embeds 2+ consecutive words of the company name where
    at least one is distinctive (e.g. 'That Sussex Bathrooms installation...').
    A single shared word like 'bathroom' is fine; a generic-only pair like
    'bathroom renovation' doesn't count."""
    if not opener or not company:
        return False
    clean = lambda s: re.sub(r"[^a-z0-9 ]", " ", s.lower()).split()
    o, c = clean(opener), clean(company)
    for i in range(len(c) - 1):
        pair = c[i:i + 2]
        if all(t in _NAME_GUARD_GENERIC for t in pair):
            continue  # generic-only pair, not a real name match
        for j in range(len(o) - 1):
            if o[j:j + 2] == pair:
                return True
    return False


# Mirrors the PROMPT's own "BANNED generic endings" list (dm_app.py ~L243) —
# the vision model doesn't always obey it, so this is a hard code-side backstop.
_BANNED_OPENER_PHRASES = [
    "grout lines", "to the wall", "against the wall", "across the wall",
    "on the wall", "through the wall", "to the system", "crisp connections",
    "clean connections", "sharp connections", "square connections",
    "pipework connections", "clean lines", "crisp lines",
    "well-finished", "well-built", "lot of work", "nice work", "good finish",
    "nicely finished", "well executed", "well laid", "nicely done",
    "pipework run", "clean pipework", "condensate run",
]


def _opener_banned_phrase(opener: str) -> str:
    """Return the banned phrase found in the opener, or '' if none."""
    low = (opener or "").lower()
    for phrase in _BANNED_OPENER_PHRASES:
        if phrase in low:
            return phrase
    return ""


def _en_opener_postprocess(opener: str, text_data: str) -> str:
    """Reject opener if it contradicts the detected niche, embeds the company
    name, or uses a banned generic phrase the prompt already forbids; fall
    back to the city line."""
    niche = _resolve_niche(text_data)
    company = _td_section(text_data, "COMPANY NAME")
    reason = ""
    banned = _opener_banned_phrase(opener)
    if not _en_opener_matches_niche(opener, niche):
        reason = f"contradicts niche '{niche}'"
    elif _opener_repeats_company_name(opener, company):
        reason = "embeds the company name"
    elif banned:
        reason = f"uses banned phrase '{banned}'"
    if reason:
        city = _td_section(text_data, "LISTED CITY")
        fallback = (f"Came across your page in {city}."
                    if city else "Came across your page.")
        print(f"  Opener REJECTED ({reason}): '{opener[:60]}' -> '{fallback}'")
        return fallback
    return opener


CUSTOM_PROMPT_BUILDER     = _en_prompt_builder
CUSTOM_OPENER_POSTPROCESS = _en_opener_postprocess


# ── Groq call ────────────────────────────────────────────────────────────────
def _run_groq(client, model: str, content, max_tokens: int = 500, temperature: float = 0.7) -> str:
    import time
    resp = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": content}],
        max_tokens=max_tokens,
        temperature=temperature,
    )
    return resp.choices[0].message.content.strip()


def _try_models(client, attempts, max_tokens=500, temperature=0.7) -> str:
    import time
    last_error = None
    for model, content in attempts:
        for attempt in range(3):
            try:
                result = _run_groq(client, model, content, max_tokens, temperature)
                return result
            except Exception as e:
                msg = str(e)
                last_error = msg
                skip_keywords = ("404", "not found", "decommissioned", "model_not_found", "no longer supported")
                if any(k in msg.lower() for k in skip_keywords):
                    print(f"{model} not available, skipping...")
                    break
                if attempt < 1:
                    time.sleep(2)
                    continue
                print(f"{model} failed ({msg[:80]}), trying next...")
                break
    raise RuntimeError(f"All Groq models failed. Last error: {last_error}")


# Praise phrases for the opener. We pin ONE per call (below) so a batch doesn't
# all read "came up mint" — each Gemini call is stateless and would otherwise
# default to the same favourite. Each fits "That [work] {phrase}, [detail]."
_OPENER_LEADS = [
    "came up mint", "looks proper tidy", "is sat lovely", "came up a treat",
    "looks the part", "looks bang on", "looks cracking", "is spot on",
    "came up lovely", "looks the business", "came up cracking", "is proper tidy",
]
# Descriptors for the SECOND clause, so it isn't always "neat ___". These must NOT
# overlap with _OPENER_LEADS words (tidy/spot on/sat lovely/bang on) or the line
# doubles up ("looks bang on, bang on ..."). Kept to physical-observation words.
_OPENER_DETAILS = [
    "clean", "crisp", "dead straight", "dead level", "lined up",
    "sharp", "spaced even", "square", "flush", "tight", "even",
]


def call_gemini(text_data: str = "", images: list = None, prompt: str = None, retry_on_reject: bool = True) -> str:
    client = Groq(api_key=API_KEY)
    if prompt is None:
        prompt = PROMPT

    # ── Step 1: extract clean facts from page text (fast, text-only) ──────────
    facts = text_data
    if text_data and len(text_data.strip()) > 200:
        try:
            extract_prompt = EXTRACT_PROMPT + f"\n\nPAGE TEXT:\n{text_data[:4000]}"
            facts = _try_models(client,
                [(m, extract_prompt) for m in FACTS_MODELS],
                max_tokens=250, temperature=0.2)
            print(f"Extracted facts:\n{facts}\n---")
        except Exception:
            facts = text_data  # fall back to raw text

    # ── Step 2: write opener from clean facts ─────────────────────────────────
    context_str = f"\n\nPAGE FACTS:\n{facts}" if facts else ""
    full_prompt = prompt + context_str
    if CUSTOM_PROMPT_BUILDER:
        full_prompt = CUSTOM_PROMPT_BUILDER(full_prompt, text_data) or full_prompt
    # Pin one praise phrase AND one detail descriptor for THIS opener so a batch
    # reads varied, not 10x "came up mint, neat ___". Each call is stateless.
    _lead_phrase = random.choice(_OPENER_LEADS)
    _detail_word = random.choice(_OPENER_DETAILS)
    full_prompt += (
        f"\n\nFOR THIS ONE: use '{_lead_phrase}' as the praise phrase and no other "
        f"(do not reach for 'mint' unless that is the given phrase). For the second "
        f"clause lean on '{_detail_word}' rather than 'neat' (avoid the word 'neat', "
        f"and do NOT repeat the praise word in the detail). "
        f"Shape it as: That [the actual work] {_lead_phrase}, {_detail_word} [detail].")

    write_attempts = []
    if images:
        vision_content = []
        for img_bytes, mime in images:
            b64 = base64.b64encode(img_bytes).decode()
            vision_content.append({"type": "image_url", "image_url": {"url": f"data:{mime};base64,{b64}"}})
        vision_content.append({"type": "text", "text": full_prompt})
        write_attempts += [(m, vision_content) for m in VISION_MODELS]
    write_attempts += [(m, full_prompt) for m in TEXT_MODELS]

    last_error = None
    last_fallback = None  # generic-but-safe opener, kept as a last resort only
    # Cap TOTAL rejections across all models/attempts, not just per-model.
    # Without this, a single niche misclassification or a banned-phrase habit
    # in one model burns through all 4 models x 2 sub-attempts = up to 8 full
    # API calls before giving up. First cut of this cap used MAX_REJECTIONS=2,
    # which measurably overcorrected — real fallback-rate data went from ~18%
    # (unbounded retries) to ~60% (2-attempt cap) on a real batch, because the
    # model has a strong, specific habit of reaching for a small set of banned
    # generic phrases ("pipework run", "grout lines", "on the wall" — all
    # already banned in the PROMPT and in _BANNED_OPENER_PHRASES) FIRST on
    # boiler/heating and bathroom content, and typically needs 3-4 tries to
    # escape that habit, not 1-2. Two changes: raised the cap to 4 (still
    # bounds the old up-to-8 worst case), AND each retry now tells the model
    # exactly which phrase it just got rejected for and to avoid it — turns
    # blind resampling into a targeted retry, so each attempt is more likely
    # to succeed within the budget instead of repeating the same mistake.
    rejection_count = 0
    MAX_REJECTIONS = 4
    avoid_note = ""
    for model, content in write_attempts:
        if rejection_count >= MAX_REJECTIONS:
            break
        for attempt in range(2):
            if rejection_count >= MAX_REJECTIONS:
                break
            try:
                import time
                this_content = content
                if avoid_note:
                    if isinstance(content, list):   # vision_content: [..., {"type": "text", "text": ...}]
                        this_content = content[:-1] + [{"type": "text", "text": content[-1]["text"] + avoid_note}]
                    else:                            # text-only full_prompt string
                        this_content = content + avoid_note
                result = _run_groq(client, model, this_content, max_tokens=400, temperature=0.92)
                if re.search(r"\b\d[\d,]*\s*(followers?|likes?|reviews?|ratings?|fans?)", result, re.I):
                    print(f"Output contained forbidden metric, retrying...")
                    continue
                if CUSTOM_OPENER_POSTPROCESS:
                    processed = CUSTOM_OPENER_POSTPROCESS(result, text_data) or result
                    if processed != result and retry_on_reject:
                        # Rejected (banned phrase / niche clash / embeds company
                        # name) — spend a BOUNDED retry budget trying for a real
                        # personalised line instead of settling immediately, but
                        # not the whole budget. Skipped for "extra" candidates
                        # (retry_on_reject=False) so a single Generate click
                        # doesn't burn 3x the budget.
                        last_fallback = processed
                        rejection_count += 1
                        banned = _opener_banned_phrase(result)
                        avoid_note = (
                            f"\n\nFOR THIS RETRY: your last attempt used the banned phrase '{banned}' — "
                            f"do not use it or anything close to it. Pick a different focal detail and verb entirely."
                        ) if banned else (
                            f"\n\nFOR THIS RETRY: your last attempt ('{result[:80]}') didn't work — "
                            f"try a genuinely different detail or angle, not a small rewording of the same line."
                        )
                        print(f"  Opener rejected ({rejection_count}/{MAX_REJECTIONS}), retrying for a personalised line...")
                        continue
                    result = processed
                return result
            except Exception as e:
                msg = str(e)
                last_error = msg
                skip_keywords = ("404", "not found", "decommissioned", "model_not_found", "no longer supported")
                if any(k in msg.lower() for k in skip_keywords):
                    print(f"{model} not available, skipping...")
                    break
                if attempt < 1:
                    time.sleep(2)
                    continue
                print(f"{model} failed ({msg[:80]}), trying next...")
                break
    if last_fallback is not None:
        print("  All attempts rejected, using generic fallback.")
        return last_fallback
    raise RuntimeError(f"All Groq models failed. Last error: {last_error}")


_BRAND_COLOR_PROMPT = """\
Look at this screenshot, cropped to the TOP of a trade business's Facebook page (cover photo and, below or overlapping it, the small circular profile picture).

Find the circular profile picture specifically, and ignore everything else on the page, the cover banner, any buttons, any job photos, any other colours visible anywhere else. Only reply with a colour if you are genuinely confident it's a clean, single-colour logo, one plain colour word (e.g. "navy blue", "forest green", "red").

Reply with exactly the word none, and nothing else, if ANY of these apply:
- The profile picture is a photo (a job, a van, a person, a team), not a clean logo/icon/wordmark
- The logo has two or more colours in it, even if one looks more prominent than the other, never guess which one is "the" brand colour
- You are not certain exactly what the profile picture shows, or which part of the image it even is
- Anything about this is unclear, a wrong colour is worse than no colour, so when in doubt, always answer none

Reply with ONLY the colour, or ONLY the word none. No explanation, no punctuation."""


def detect_brand_color(image_bytes: bytes, mime: str = "image/jpeg") -> str:
    """Best-effort read of a business's actual logo colour, cropped by the
    caller to the top of the page (cover + avatar) before this is called —
    sending the WHOLE page screenshot made the small circular avatar a tiny,
    hard-to-read fraction of the image, which is why obvious logos were
    still coming back "none" in practice. Only ONE model call (no retries)
    since this is a nice-to-have, not worth spending the retry budget on.
    Returns "" (never claims a colour) on anything except a short, clean
    colour name, so a hallucinated or hedged answer never ends up in the
    message.

    Confidence bar was raised back up after a real miss: a yellow/black
    logo came back "red", a genuinely wrong colour, not just a missed one.
    The earlier "trust your first read, don't overthink" wording (added to
    fix an over-cautious "none" rate) pushed it too far the other way. A
    wrong claim to a real prospect is worse than staying silent, so this
    now explicitly tells the model to bail to "none" on any multi-colour
    logo or any uncertainty at all, accepting more misses in exchange."""
    try:
        client = Groq(api_key=API_KEY)
        b64 = base64.b64encode(image_bytes).decode()
        content = [
            {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{b64}"}},
            {"type": "text", "text": _BRAND_COLOR_PROMPT},
        ]
        result = _run_groq(client, VISION_MODELS[0], content, max_tokens=16, temperature=0.1)
        result = result.strip().strip(".").lower()
        # Strip filler the model sometimes prepends despite the instruction
        # ("the colour is navy blue", "colour: red") instead of rejecting
        # the whole answer for not being in the exact bare format.
        result = re.sub(r"^(the )?(main )?(colou?r|logo)( is| looks like)?[:\-]?\s*", "", result).strip()
        if not result or result == "none":
            return ""
        # Guard against the model ignoring the format (a full sentence, a
        # hedge, "I think navy blue" etc.) — only accept a short, plain
        # colour name.
        if len(result.split()) > 3 or not re.fullmatch(r"[a-z\- ]+", result):
            return ""
        return result
    except Exception as e:
        print(f"Brand color detection skipped: {e}")
        return ""


def fix_grammar(text: str) -> str:
    """Run the final DM through Groq to fix any grammatical errors. Hungarian-aware."""
    if not GRAMMAR_CHECK or not text.strip():
        return text
    try:
        client = Groq(api_key=API_KEY)
        prompt = (
            "Te egy magyar nyelvhelyességi ellenőrző vagy. "
            "Javítsd ki a következő szöveg összes nyelvtani hibáját. "
            "Fontos szabályok: "
            "1. Ha egy kérdés 'Ti,' szóval kezdődik vesszővel, töröld a 'Ti,' részt. "
            "2. Ne változtasd meg a jelentést, a stílust vagy a hangnemet. "
            "3. Ne adj hozzá semmit, ne magyarázz. Csak a javított szöveget add vissza.\n\n"
            f"{text}"
        )
        result = _run_groq(client, "llama-3.3-70b-versatile", prompt, max_tokens=300, temperature=0.1)
        print(f"Grammar fix: {text!r} -> {result!r}")
        return result
    except Exception as e:
        print(f"Grammar check failed, using original: {e}")
        return text


def tidy_message(text: str) -> str:
    """Instant, local cleanup of the assembled DM. Replaces the old LLM
    grammar pass (which was a HUNGARIAN checker being run on ENGLISH copy —
    wasted a round-trip per lead and could mangle good lines). Cosmetic only:
    collapse repeated spaces, drop space-before-punctuation, trim line ends.
    Never changes wording. Preserves the paragraph breaks (blank lines)."""
    if not text:
        return text
    out = []
    for line in text.split("\n"):
        line = line.replace("—", ", ").replace("–", ", ")  # no em/en dashes (an AI tell)
        line = re.sub(r"[ \t]{2,}", " ", line)        # collapse runs of spaces/tabs
        line = re.sub(r"[ \t]+([,.!?;:])", r"\1", line)  # no space before punctuation
        line = re.sub(r",\s*,", ",", line)            # tidy any doubled comma from the swap
        out.append(line.rstrip())
    return "\n".join(out).strip()


# ── Routes ─────────────────────────────────────────────────────────────────────
# Set True by dm_app_hu.py (via hu_mech.register) to show the Autószerelő tab.
MECHANICS_SECTION = False
# Set by hu_mech.register() too: the mechanics sheet prefilled in that tab.
MECH_DEFAULT_EXCEL = ""

@app.route("/")
def index():
    return render_template("dm.html", default_excel=DEFAULT_EXCEL, variant=MESSAGE_VARIANT,
                           mech_enabled=MECHANICS_SECTION,
                           mech_excel=MECH_DEFAULT_EXCEL)

@app.route("/toggle_variant", methods=["POST"])
def toggle_variant():
    """Cycle the message variant: full -> stripped -> british -> funny -> full."""
    global MESSAGE_VARIANT
    cycle = {"full": "stripped", "stripped": "british",
             "british": "funny", "funny": "full"}
    MESSAGE_VARIANT = cycle.get(MESSAGE_VARIANT, "full")
    _save_variant(MESSAGE_VARIANT)
    print(f"Message variant -> {MESSAGE_VARIANT}")
    return jsonify({"variant": MESSAGE_VARIANT})

@app.route("/generate", methods=["POST"])
def generate():
    pw = None            # Playwright connection kept open scan->click; see below
    page_for_click = None
    try:
        if not API_KEY:
            return jsonify({"error": "GROQ_API_KEY not set in .env"}), 500

        images    = []
        text_data = ""
        warning   = ""
        name_hint = ""

        # Manual image upload (screenshot pasted in the app)
        if "image" in request.files and request.files["image"].filename:
            f = request.files["image"]
            images.append((f.read(), f.mimetype or "image/jpeg"))
            print(f"Manual image uploaded: {f.mimetype}, {len(images[0][0])} bytes")

        # Manual notes
        notes = request.form.get("notes", "").strip()
        if notes:
            text_data += f"\n--- USER NOTES ---\n{notes}"

        # Capture from Chrome (primary method — only if no manual image). Keeps
        # the same Playwright connection open so we can click Message with it
        # later, instead of connecting to Chrome a second time at the end.
        if not images:
            hint_url = request.form.get("url", "").strip()
            pw, page_for_click, captured_images, captured_text = _capture_and_keep_open(hint_url)
            if captured_images or captured_text:
                images = captured_images  # may be empty if text was rich — that's fine
                if captured_text:
                    name_hint = extract_first_name(captured_text)
                    if name_hint:
                        print(f"Name detected: {name_hint}")
                        text_data += f"\n--- DETECTED FIRST NAME ---\n{name_hint}"
                    text_data += f"\n--- PAGE TEXT ---\n{captured_text}"
                print(f"Captured {len(images)} screenshots, {len(captured_text)} chars text from Chrome")
            else:
                warning = (
                    "No Facebook/Instagram profile tab found in Chrome. "
                    "Open the profile page in Chrome first, then click Generate."
                )
                if not text_data.strip():
                    text_data = "--- NO PAGE DATA --- Do NOT invent any details."

        if not images and not text_data.strip() and not notes:
            return jsonify({
                "error": "Open the profile in Chrome first, then click Generate. "
                         "Or paste a screenshot / type profile info."
            }), 400

        if MESSAGE_VARIANT == "stripped":
            # Direct one-line variant — no opener, so skip the vision call.
            dm = _dm_case(tidy_message(pick_stripped_message(
                text_data, name_hint, extract_company_name(text_data))))
            dms = [dm]
        elif MESSAGE_VARIANT == "british":
            # British casual variant — uses vision opener, short punchy body.
            company = extract_company_name(text_data)

            def _assemble_british(opener: str) -> str:
                opener = patch_fallback_opener(opener, text_data)
                opener = _apply_greeting(opener, name_hint)
                return _dm_case(tidy_message(pick_british_message(
                    opener, text_data, name_hint, company)))

            dms, seen = [], set()
            first = call_gemini(text_data=text_data, images=images if images else None)
            seen.add(re.sub(r"\s+", " ", first.strip().lower()))
            dms.append(_assemble_british(first))
            for _ in range(2):
                try:
                    raw = call_gemini(text_data=text_data, images=images if images else None, retry_on_reject=False)
                except Exception:
                    break
                key = re.sub(r"\s+", " ", raw.strip().lower())
                if key in seen:
                    continue
                seen.add(key)
                dms.append(_assemble_british(raw))
            dm = dms[0]
        else:
            company = extract_company_name(text_data)
            trade   = detect_trade(text_data)

            def _assemble(opener: str) -> str:
                opener = opener.rstrip(".!,") + "."
                opener = patch_fallback_opener(opener, text_data)
                opener = _apply_greeting(opener, name_hint)
                body = pick_body(trade, text_data, company_name=company)
                return _dm_case(tidy_message((opener + " " + body).strip()))

            # Generate up to 3 candidates so you pick the best read instead of
            # regenerating by hand. Vision temp 0.92 + rotated body keep them
            # distinct; we dedup on the opener so identical fallbacks collapse.
            dms, seen = [], set()
            # First candidate — let a hard failure bubble up with its real error.
            first = call_gemini(text_data=text_data, images=images if images else None)
            seen.add(re.sub(r"\s+", " ", first.strip().lower()))
            dms.append(_assemble(first))
            # Two more best-effort candidates; ignore transient failures.
            for _ in range(2):
                try:
                    raw = call_gemini(text_data=text_data, images=images if images else None, retry_on_reject=False)
                except Exception:
                    break
                key = re.sub(r"\s+", " ", raw.strip().lower())
                if key in seen:
                    continue
                seen.add(key)
                dms.append(_assemble(raw))
            dm = dms[0]
        clipboard = copy_to_clipboard(dm)
        # Open the chat and focus the composer so the user just Ctrl+V's. Reuses
        # the page found during the scan above (one connection, not two) when
        # available; a manual image upload skips the scan, so falls back to its
        # own (rarer) connect-and-click.
        if page_for_click:
            msg_opened = _click_message_btn(page_for_click)
        else:
            msg_opened = _open_current_messenger(request.form.get("url", "").strip())
        return jsonify({"dm": dm, "dms": dms, "warning": warning, "clipboard": clipboard, "msg_opened": msg_opened})

    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500
    finally:
        if pw:
            try:
                pw.stop()
            except Exception:
                pass

@app.route("/followup", methods=["POST"])
def followup():
    data  = request.get_json(force=True)
    trade = (data.get("trade") or "quotes").lower().strip()
    touch = int(data.get("touch", 1)) - 1  # 0-indexed

    sequence = FOLLOWUPS.get(trade) or FOLLOWUPS["quotes"]
    touch = max(0, min(touch, len(sequence) - 1))
    msg   = sequence[touch]

    copy_to_clipboard(msg)
    return jsonify({"message": msg})


@app.route("/generate_voicenote", methods=["POST"])
def generate_voicenote():
    try:
        if not API_KEY:
            return jsonify({"error": "GROQ_API_KEY not set in .env"}), 500

        images    = []
        text_data = ""
        warning   = ""

        if "image" in request.files and request.files["image"].filename:
            f = request.files["image"]
            images.append((f.read(), f.mimetype or "image/jpeg"))

        notes = request.form.get("notes", "").strip()
        if notes:
            text_data += f"\n--- USER NOTES ---\n{notes}"

        if not images:
            hint_url = request.form.get("url", "").strip()
            captured_images, captured_text = capture_profile_tab(hint_url)
            if captured_images:
                images = captured_images
                if captured_text:
                    name_hint = extract_first_name(captured_text)
                    if name_hint:
                        text_data += f"\n--- DETECTED FIRST NAME ---\n{name_hint}"
                    text_data += f"\n--- PAGE TEXT ---\n{captured_text}"
            else:
                warning = (
                    "No Facebook/Instagram profile tab found in Chrome. "
                    "Open the profile page in Chrome first, then click Generate."
                )
                if not text_data.strip():
                    text_data = "--- NO PAGE DATA --- Do NOT invent any details."

        if not images and not notes:
            return jsonify({
                "error": "Open the profile in Chrome first, then click Generate. "
                         "Or paste a screenshot / type profile info."
            }), 400

        script = call_gemini(text_data=text_data, images=images if images else None, prompt=VOICE_NOTE_PROMPT)
        return jsonify({"script": script, "warning": warning})

    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500

# ── Auto Sender state ──────────────────────────────────────────────────────────
import openpyxl as _xl
from pathlib import Path as _Path

# ── "Funny" message style ─────────────────────────────────────────────────────
# A whole message set, picked with the "Funny" button next to Message style.
# Each one is a saying taken literally, with a photo of the literal thing, so it
# reads like nothing else in their inbox. The joke buys attention; the middle
# paragraph still says plainly what the thing does and what the ask is.
#
# Which of the five a lead gets is decided by ROW NUMBER, so it stays stable
# across restarts and doesn't drift when leads get skipped. The greeting varies
# the same way — the same text going out 100 times is exactly what spam filters
# look for, so no two consecutive rows open identically.

FUNNY_IMAGE_DIR = _Path(__file__).parent

# Trade label (from the sheet's Trade Type column, or the niche detector) mapped
# to how you'd actually SAY it: "i build these for bathroom fitters", never
# "for bathroom renovation". Matched loosely — substring, lowercased.
NICHE_WORDS = [
    ("bathroom",         "bathroom fitters"),
    ("wet room",         "bathroom fitters"),
    ("kitchen",          "kitchen fitters"),
    ("gas",              "gas engineers"),
    ("boiler",           "gas engineers"),
    ("heat pump",        "heat pump installers"),
    ("heating",          "heating engineers"),
    ("air con",          "air con firms"),
    ("aircon",           "air con firms"),
    ("hvac",             "air con firms"),
    ("air conditioning", "air con firms"),
    ("refrigerat",       "air con firms"),
    ("plumb",            "plumbers"),
    ("roof",             "roofers"),
    ("electric",         "sparkies"),
    ("sparky",           "sparkies"),
    ("sparkies",         "sparkies"),
    ("drone",            "drone operators"),
    ("aerial",           "drone operators"),
    ("solar",            "solar installers"),
    ("tiling",           "tilers"),
    ("tiler",            "tilers"),
    ("window",           "window fitters"),
    ("glazing",          "window fitters"),
    ("door",             "door fitters"),
    ("landscap",         "landscapers"),
    ("paving",           "driveway firms"),
    ("driveway",         "driveway firms"),
    ("decking",          "decking firms"),
    ("fencing",          "fencing firms"),
    ("plaster",          "plasterers"),
    ("render",           "rendering firms"),
    ("joiner",           "joiners"),
    ("joinery",          "joiners"),
    ("carpent",          "carpenters"),
    ("floor",            "flooring firms"),
    ("decorat",          "decorators"),
    ("painting",         "decorators"),
    ("drainage",         "drainage firms"),
    ("groundwork",       "groundworkers"),
    ("brick",            "bricklayers"),
    ("extension",        "extension builders"),
    ("renovation",       "renovation firms"),
    ("construction",     "construction firms"),
    ("building",         "builders"),
    ("builder",          "builders"),
]

NICHE_DEFAULT = "trade firms"

# The same trade said as the WORK, not the worker: you'd say "saw you're doing
# bathrooms", never "saw you're doing bathroom fitters". Keyed off the output of
# niche_words() so there's only one place to add a trade.
NICHE_WORK = {
    "bathroom fitters":      "bathrooms",
    "kitchen fitters":       "kitchens",
    "gas engineers":         "boilers",
    "heating engineers":     "heating",
    "heat pump installers":  "heat pumps",
    "air con firms":         "air con",
    "plumbers":              "plumbing",
    "roofers":               "roofs",
    "sparkies":              "electrical work",
    "solar installers":      "solar",
    "tilers":                "tiling",
    "window fitters":        "windows",
    "door fitters":          "doors",
    "landscapers":           "gardens",
    "driveway firms":        "driveways",
    "decking firms":         "decking",
    "fencing firms":         "fencing",
    "plasterers":            "plastering",
    "rendering firms":       "rendering",
    "joiners":               "joinery",
    "carpenters":            "carpentry",
    "flooring firms":        "flooring",
    "decorators":            "decorating",
    "drainage firms":        "drainage",
    "groundworkers":         "groundwork",
    "bricklayers":           "brickwork",
    "extension builders":    "extensions",
    "renovation firms":      "renovations",
    "construction firms":    "building work",
    "builders":              "building work",
    "drone operators":       "aerial work",
}


def funny_personal_line(niche: str = "", city: str = "") -> str:
    """A light personal touch built from the SPREADSHEET (Trade Type + City),
    so it costs no page scan and no AI call at all.

    It's weaker than "that resin driveway you posted last week", but it's free
    and instant, and it still says "I picked you" rather than "I picked 400 of
    you". Returns "" when the sheet hasn't got enough to say anything true —
    a wrong guess is worse than no line.
    """
    city = (city or "").strip()
    work = NICHE_WORK.get(niche_words(niche), "")
    if city and work:
        return f" saw you're doing {work} around {city}, so thought i'd give you a go."
    if city:
        return f" saw you're covering {city}, so thought i'd give you a go."
    if work:
        return f" saw you're doing {work}, so thought i'd give you a go."
    return ""



def niche_words(niche: str) -> str:
    """'Bathroom Renovation' -> 'bathroom fitters'. Falls back to 'trade firms'
    rather than dropping a raw label into the middle of a sentence."""
    n = (niche or "").strip().lower()
    if not n:
        return NICHE_DEFAULT
    for key, words in NICHE_WORDS:
        if key in n:
            return words
    return NICHE_DEFAULT


# Greetings, rotated so consecutive sends never open the same way. The "mate"
# set is used when we haven't got a name off the page, which is most of the
# time on a business Page.
GREETINGS_NAMED = [
    "alright {n}?",
    "hi {n},",
    "hey {n},",
    "alright {n} mate?",
    "hiya {n},",
    "hey up {n},",
]

GREETINGS_MATE = [
    "alright mate?",
    "hi mate,",
    "hey mate,",
    "alright mate,",
    "hiya mate,",
    "alright there mate?",
]


def funny_greeting(first_name: str, row: int) -> str:
    """Varied opener. Uses the first name ONLY if it's a known real first name.

    A blocklist of company words can't win this — "Top Loft Ltd" gives "Top",
    and there's an endless supply of Elite / Prime / Ace / Pro style business
    names. So this is an ALLOWLIST: the word has to appear in FIRST_NAMES (586
    common UK/US first names) or we say "mate". Getting "mate" wrong costs
    nothing; getting "alright Top?" wrong wastes the lead.
    """
    n = (first_name or "").strip().split(" ")[0].strip(".,!?-'’")
    if n.lower() not in FIRST_NAMES:
        n = ""
    if n:
        return GREETINGS_NAMED[row % len(GREETINGS_NAMED)].format(n=n.capitalize())
    return GREETINGS_MATE[row % len(GREETINGS_MATE)]


# The five. {greet} and {niche} are filled per lead. "image" is the photo that
# makes the joke land — a missing file means we fall back to the normal message
# rather than send "that's my foot" with no foot attached.
FUNNY_MESSAGES = [
    {
        "key": "foot",
        "image": ["foot.jpg", "foot 2.jpg"],
        "text": """{greet} that's my foot.

trying to get it in the door.{personal}

i build quoting assistants for {niche}. not to get you more work, you've got plenty. it sits on your site and estimates jobs, so you stop losing evenings to people who ask and then vanish.

made you one already. want it here or on WhatsApp?""",
    },
    {
        "key": "dark",
        "image": ["black.jpg"],
        "text": """{greet} this is a shot in the dark.

get it? shot and dark.{personal}

i make quoting assistants for {niche}. it's for lads already flat out, not for drumming up work. it estimates the job the second someone asks, so the right customers are the ones who actually get through.

yours is made. send it here or is WhatsApp easier?""",
    },
    {
        "key": "bush",
        "image": ["bush.jpg"],
        "text": """{greet} not going to beat around the bush.

so there's the bush. beaten.{personal}

i build quoting assistants for {niche}. if you're fully booked, good, that's exactly who it's for. it tells people what a job'll cost before they ring you, so it's the right customers getting through, not just more calls.

yours is done. want a look? here or WhatsApp.""",
    },
    {
        "key": "cards",
        "image": ["cards on the table.webp"],
        "text": """{greet} cards on the table.

quoting assistants for {niche}, that's what i do.{personal}

built you one without being asked. it does the bit you haven't got time for: estimates jobs on the spot, stops the time wasters reaching you at all. not more leads, just better ones.

want to see it? here or WhatsApp.""",
    },
    {
        "key": "cold",
        "image": ["ice.jpg"],
        "text": """{greet} this is a cold message. seems pretty cold to me.{personal}

i put quoting assistants on websites for {niche} and i've done you one already.

it estimates jobs on your site so you're not writing estimates out at 9pm for folk who go quiet. if you're booked solid, that's the whole point.

here or WhatsApp?""",
    },
]


def _funny_image(msg: dict, row: int):
    """(bytes, mime) for this message's photo, or None if the file is missing.

    Messages carrying two photos alternate, so the same picture doesn't go out
    every single time that message comes round.
    """
    names = msg["image"]
    name = names[(row // len(FUNNY_MESSAGES)) % len(names)]
    try:
        f = FUNNY_IMAGE_DIR / name
        if not f.is_file():
            print(f"  Funny image missing: {name}")
            return None
        mime = "image/webp" if f.suffix.lower() == ".webp" else "image/jpeg"
        data = f.read_bytes()
        print(f"  Funny image loaded: {name} ({len(data):,} bytes)")
        return (data, mime)
    except Exception as e:
        print(f"  Funny image failed to load ({name}): {e}")
        return None


def build_funny_message(row: int, first_name: str = "", niche: str = "",
                        city: str = ""):
    """(text, (image_bytes, mime), key) for this lead, or None if unusable.

    None means the caller keeps the normal personalised DM.
    """
    try:
        msg = FUNNY_MESSAGES[row % len(FUNNY_MESSAGES)]
        img = _funny_image(msg, row)
        if not img:
            return None
        text = msg["text"].format(
            greet=funny_greeting(first_name, row),
            niche=niche_words(niche),
            personal=funny_personal_line(niche, city),
        )
        return (text, img, msg["key"])
    except Exception as e:
        print(f"  Funny message build failed: {e}")
        return None


_auto_state = {
    "excel_path":    None,
    "current_row":   None,
    "sent_today":    0,
    "wb":            None,
    "ws":            None,
    "cols":          {},
    "last_screenshot": None,   # (bytes, mime) for the current lead's picture — set by /auto_next, consumed by /auto_copy_screenshot
}

# Separate state for the Owner Emails section so it never interferes with the
# Auto Sender (DM) flow — you can load a different list in each.
_email_state = {
    "excel_path":    None,
    "current_row":   None,
    "sent_today":    0,
    "wb":            None,
    "ws":            None,
    "cols":          {},
}

DAILY_CAP = 30

def _load_excel(path: str):
    wb = _xl.load_workbook(path)
    ws = wb.active
    headers = [str(c.value or "").strip().lower() for c in ws[1]]

    created = []
    def _find_or_create(matchers, label):
        """Return the index of the first header matched by the EARLIEST
        predicate in `matchers` that matches anything, creating a new column
        titled `label` at the end if none do.

        Ordered, not a single predicate, because a loose one-shot match binds
        to whatever happens to sit leftmost. Real damage this caused on
        HUNGARIAN_LEADS_MASTER_EVERYTHING.xlsx, whose headers are
        "... Status ... Followers ... DM Sent? ... Reply? ... Contacted? ...":
          - "Contacted?" was matched by `"contacted" in h or h == "status"`,
            and Status is column 6 while Contacted? is column 13 — so the app
            read and wrote Status, and the 387 real sends recorded in
            Contacted? were invisible to it.
          - "Follow-Ups" was matched by `"follow" in h`, which hits FOLLOWERS.
            The follower count was therefore read as the number of nudges
            already sent, so `fups >= FOLLOWUP_TOUCHES` was true for every page
            with 3+ followers and the entire sheet was silently excluded from
            the follow-up queue forever. Marking a nudge would also have
            overwritten the follower count with 1/2/3.
        Each list therefore runs exact match -> specific substring -> loose
        fallback, so an explicitly-named column always wins over a coincidence.
        """
        for match in matchers:
            idx = next((i for i, h in enumerate(headers) if match(h)), None)
            if idx is not None:
                return idx
        ws.cell(row=1, column=len(headers) + 1, value=label)
        headers.append(label.lower())
        created.append(label)
        return len(headers) - 1

    # Follow-up tracking columns are created on first load if absent.
    # Match the SENT-tracking column only — NOT the lead-contact column. On the
    # newer list "Contact (Email or Personal FB)" also contains "contact", so the
    # old loose "contact" match hijacked the email column and overwrote it with
    # "Skip"/"Yes". Match "contacted"/"outreach status" instead.
    contacted_idx = _find_or_create([
        lambda h: h.strip() in ("contacted?", "contacted"),
        lambda h: "contacted" in h,
        lambda h: "dm sent" in h,
        lambda h: "outreach" in h,
        lambda h: h.strip() == "status",
    ], "Contacted?")
    sent_date_idx = _find_or_create([
        lambda h: "sent date" in h,
        lambda h: "date" in h,
    ], "Sent Date")
    replied_idx   = _find_or_create([
        lambda h: h.strip() in ("replied?", "replied", "reply?", "reply"),
        lambda h: "repl" in h,
    ], "Replied?")
    # "follower" guard is the whole point — see the docstring above.
    followups_idx = _find_or_create([
        lambda h: "follow-up" in h or "followup" in h or "follow up" in h,
        lambda h: "follow" in h and "follower" not in h,
    ], "Follow-Ups")
    if created:
        wb.save(path)  # only write when we actually added a column

    cols = {
        "url":       next((i for i, h in enumerate(headers) if "url" in h or "facebook" in h), 1),
        "name":      next((i for i, h in enumerate(headers) if "company" in h or "name" in h), 0),
        "city":      next((i for i, h in enumerate(headers) if "city" in h or "location" in h), None),
        "niche":     next((i for i, h in enumerate(headers) if "niche" in h or "trade" in h), None),
        # The lead's primary contact (email OR personal FB) and its type. Used to
        # decide per lead whether we send a DM or write an owner email.
        "contact":      next((i for i, h in enumerate(headers) if "contact" in h and ("email" in h or "fb" in h or "personal" in h)), None),
        "contact_type": next((i for i, h in enumerate(headers) if "contact" in h and "type" in h), None),
        "company_fb":   next((i for i, h in enumerate(headers) if ("company" in h or "business" in h) and ("facebook" in h or "fb" in h or "page" in h)), None),
        "contacted": contacted_idx,
        "sent_date": sent_date_idx,
        "replied":   replied_idx,
        "followups": followups_idx,
    }
    return wb, ws, cols

def _save_wb():
    """Save the loaded workbook, retrying briefly if the file is locked
    (open in Excel, or mid OneDrive-sync). Raises a clear, actionable
    error if it stays locked so the UI can tell the user what to do."""
    import time as _t
    wb   = _auto_state.get("wb")
    path = _auto_state.get("excel_path")
    if wb is None or not path:
        return
    for attempt in range(6):
        try:
            wb.save(path)
            return
        except PermissionError:
            _t.sleep(0.4)
    raise PermissionError(
        "Can't write to the Excel file — it's open in Excel (or still syncing in "
        "OneDrive). Close the spreadsheet in Excel, wait for OneDrive to finish, "
        "then click again."
    )

def _today_str() -> str:
    from datetime import date
    return date.today().isoformat()

def _parse_date(v):
    """Coerce an Excel cell (datetime, date, or 'YYYY-MM-DD' string) to a date."""
    from datetime import datetime, date
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    if not s:
        return None
    try:
        return datetime.fromisoformat(s[:19]).date()
    except Exception:
        try:
            return datetime.strptime(s[:10], "%Y-%m-%d").date()
        except Exception:
            return None

def _to_int(v):
    try:
        return int(float(str(v).strip()))
    except Exception:
        return 0

def _next_lead(ws, cols, skip_row=None):
    """Next uncontacted DM lead (Auto Sender). DM-only, unchanged: walks rows
    whose URL column is a real link. Email leads are handled by the separate
    Owner Emails section via _next_email_lead()."""
    for row in ws.iter_rows(min_row=2):
        excel_row = row[0].row
        if skip_row is not None and excel_row == skip_row:
            continue
        url = str(row[cols["url"]].value or "").strip()
        contacted = str(row[cols["contacted"]].value or "").strip().lower() if cols["contacted"] is not None else ""
        if not url.startswith("http"):
            continue  # skip blank or invalid URL rows silently
        low = url.lower()
        if "facebook.com" not in low and "instagram.com" not in low:
            # Website-only lead (combined "Facebook / Website URL" column): there's
            # no page to DM, so don't open the company site. Skip it for the DM run.
            continue
        if contacted not in ("yes", "y", "sent", "1", "true", "skip"):
            return {
                "row":   excel_row,
                "url":   url,
                "name":  str(row[cols["name"]].value or "").strip(),
                "city":  str(row[cols["city"]].value or "").strip()  if cols["city"]  is not None else "",
                "niche": str(row[cols["niche"]].value or "").strip() if cols["niche"] is not None else "",
            }
    return None


def _next_email_lead(ws, cols, skip_row=None):
    """Next uncontacted EMAIL lead (Owner Emails section). Only rows whose
    contact is an email address (or Contact Type says Email). Returns the email
    address plus the Company Facebook page to scan for a personal line."""
    for row in ws.iter_rows(min_row=2):
        excel_row = row[0].row
        if skip_row is not None and excel_row == skip_row:
            continue

        def _v(key):
            i = cols.get(key)
            return str(row[i].value or "").strip() if i is not None else ""

        if _v("contacted").lower() in ("yes", "y", "sent", "1", "true", "skip"):
            continue

        contact = _v("contact") or _v("url")
        ctype   = _v("contact_type").lower()
        if not (("@" in contact) or ("email" in ctype)):
            continue  # not an email lead — leave it for the DM flow

        company_fb = _v("company_fb")
        scan_url   = company_fb if company_fb.startswith("http") else ""
        return {
            "row":   excel_row,
            "email": contact,
            "url":   scan_url,   # Company FB page to scan ("" if none)
            "name":  _v("name"),
            "city":  _v("city"),
            "niche": _v("niche"),
        }
    return None

def _find_message_box(ctx):
    """Find Messenger input across all open tabs."""
    import time as _time
    _time.sleep(1.2)
    selectors = [
        '[aria-label="Message"]', '[aria-label="Aa"]',
        'div[contenteditable="true"][role="textbox"]',
        'div[contenteditable="true"]',
    ]
    for pg in reversed(list(ctx.pages)):
        if "facebook.com" in pg.url or "messenger.com" in pg.url:
            for sel in selectors:
                try:
                    el = pg.locator(sel).first
                    if el.is_visible(timeout=1200):
                        return pg, el
                except:
                    continue
    return None, None

def _click_message_btn(page):
    selectors = [
        # English Facebook
        '[aria-label="Send message"]', '[aria-label="Message"]',
        # Hungarian Facebook
        '[aria-label="Üzenet küldése"]', '[aria-label="Üzenet"]',
        # Generic fallbacks
        'a[href*="messenger.com/t/"]',
        'div[role="button"]:has-text("Message")',
        'div[role="button"]:has-text("Üzenet")',
        'a:has-text("Send message")', 'a:has-text("Message")',
        'a:has-text("Üzenet küldése")',
    ]
    for sel in selectors:
        try:
            el = page.locator(sel).first
            if el.is_visible(timeout=800):
                el.click()
                page.wait_for_timeout(2000)
                return True
        except:
            continue
    return False


def _type_dm_into_box(page, text: str) -> bool:
    """Type the DM directly into the Messenger composer, instead of relying on
    clipboard + manual paste. This is what frees the clipboard up for the
    picture — if the text used the clipboard, copying the picture right after
    would overwrite it before it could be pasted. Best-effort, never raises;
    returns False if no box is found so the caller falls back to the existing
    copy-to-clipboard + manual-paste path.

    ONLY matches Facebook's two actual aria-labels for the message composer
    ("Message" / "Aa") — deliberately NOT a generic contenteditable selector.
    A bare 'div[contenteditable="true"]' matches ANY editable field on the
    page, including a comment box on the post itself, which is exactly how
    this typed a DM into a public comment instead of a private message once
    the specific selectors didn't match fast enough. Failing safe (returning
    False, falling back to clipboard) is much better than typing into the
    wrong box."""
    selectors = ['[aria-label="Message"]', '[aria-label="Aa"]']
    for sel in selectors:
        try:
            el = page.locator(sel).first
            if el.is_visible(timeout=1500):
                el.click()
                page.keyboard.type(text, delay=12)   # human-like pace
                return True
        except Exception:
            continue
    return False


def _find_composer(page, total_timeout_ms: int = 4000):
    """Search the main document AND every iframe on the page for Facebook's
    message composer. Diagnostics showed the page URL never changes after
    clicking Message — Facebook opens a docked chat popup in place rather
    than navigating, and that popup commonly renders inside an iframe. A
    plain page.locator() only ever searches the top-level document, so it
    silently never found the composer even though it existed. Polls for up
    to total_timeout_ms since the popup can take a moment to render. Returns
    the matching Locator, or None."""
    import time as _t
    # The real label, confirmed from a live diagnostic dump: Hungarian
    # Facebook labels the composer "Írás nekik: <Page Name>" (literally
    # "Write to them: <Page Name>") — NOT "Üzenet", which was a guess that
    # turned out wrong. The page name varies per lead, so this matches on the
    # "Írás nekik" prefix. "Write to" covers the English equivalent for the
    # same reason. Constrained to contenteditable+role=textbox so a
    # label-bearing BUTTON or a comment box ("Hozzászólás írása…") can never
    # match — confirmed via the same dump: two "Hozzászólás írása…" boxes
    # were present and correctly did NOT match this pattern.
    selectors = [
        'div[contenteditable="true"][role="textbox"][aria-label^="Írás nekik"]',
        'div[contenteditable="true"][role="textbox"][aria-label^="Write to"]',
        'div[contenteditable="true"][role="textbox"][aria-label="Aa"]',
        '[aria-label="Aa"]',
    ]
    deadline = _t.time() + (total_timeout_ms / 1000)
    while _t.time() < deadline:
        for fr in page.frames:
            for sel in selectors:
                try:
                    el = fr.locator(sel).first
                    if el.count() > 0 and el.is_visible(timeout=200):
                        print(f"  Composer found via {sel!r} in frame {fr.url[:70]!r}")
                        return el
                except Exception:
                    continue
        page.wait_for_timeout(300)
    print(f"  No composer found in any of {len(page.frames)} frame(s). Page URL: {page.url}")
    # Diagnostic: dump every editable textbox's aria-label so a future miss
    # tells us the actual label to add instead of leaving us guessing.
    try:
        labels = page.evaluate("""() =>
            Array.from(document.querySelectorAll('[contenteditable="true"], [role="textbox"]'))
                 .map(el => ({label: el.getAttribute('aria-label'), role: el.getAttribute('role'),
                              editable: el.getAttribute('contenteditable')}))
        """)
        print(f"  Editable elements on page: {labels}")
    except Exception:
        pass
    return None


def _paste_image_into_box(page, image_bytes: bytes) -> bool:
    """Attach the picture directly into the Messenger composer, so it's
    sitting there ready — no manual Ctrl+V needed for it at all. Only the DM
    TEXT is left for the user to handle (copy/paste or type themselves).

    Same strict, Messenger-only match as _type_dm_into_box() (via
    _find_composer) — this is what makes it safe: if the real composer isn't
    found, this does nothing rather than pasting an image into the wrong
    place (e.g. a comment box). Once focused, it sets the OS clipboard to the
    image and simulates Ctrl+V — Messenger's own paste handler turns that
    into an attached image preview, not a sent message. Nothing gets sent
    automatically; the user still adds their text and hits Enter themselves."""
    el = _find_composer(page)
    if el is None:
        return False
    try:
        el.click()
        if not copy_image_to_clipboard(image_bytes):
            print("  Clipboard image copy failed.")
            return False
        page.keyboard.press("Control+V")
        page.wait_for_timeout(600)   # let the attachment preview render
        print("  Ctrl+V sent to composer.")
        return True
    except Exception as _e:
        print(f"  Paste into composer failed: {_e}")
        return False


def _open_current_messenger(hint_url: str = "") -> bool:
    """Find the already-open Facebook profile tab (same matching as
    capture_profile_tab) and click its Message button, so the composer is
    open and focused and the user only needs to Ctrl+V. Best-effort, never
    raises — a failed auto-click just means the user opens the chat by hand
    like before."""
    import socket
    from playwright.sync_api import sync_playwright
    try:
        s = socket.create_connection(("127.0.0.1", CDP_PORT), timeout=3)
        s.close()
    except Exception:
        return False
    try:
        with sync_playwright() as p:
            browser = p.chromium.connect_over_cdp(f"http://localhost:{CDP_PORT}")
            context = browser.contexts[0]
            target_page = None
            if hint_url:
                hint_clean = hint_url.rstrip("/").split("?")[0].lower()
                for pg in context.pages:
                    if hint_clean in pg.url.lower():
                        target_page = pg
                        break
            if not target_page:
                for pg in context.pages:
                    u = pg.url.lower()
                    if "facebook.com/" in u and "messenger.com" not in u and "/messages" not in u:
                        target_page = pg
                        break
            if not target_page:
                return False
            return _click_message_btn(target_page)
    except Exception as e:
        print(f"  Could not auto-open Messenger: {e}")
        return False

@app.route("/auto_load", methods=["POST"])
def auto_load():
    """Load Excel and return first uncontacted lead info."""
    try:
        data = request.get_json()
        path = data.get("path", "").strip()

        # Auto-detect if no path given
        if not path:
            for f in _Path(__file__).parent.glob("*.xlsx"):
                nm = f.name.lower()
                if any(k in nm for k in ("lead", "dm", "contact", "prospect", "outreach")):
                    path = str(f); break
            if not path:
                files = list(_Path(__file__).parent.glob("*.xlsx"))
                path = str(files[0]) if files else ""

        if not path or not _Path(path).exists():
            return jsonify({"error": f"Excel file not found: {path}"}), 400

        wb, ws, cols = _load_excel(path)
        _auto_state.update({"excel_path": path, "wb": wb, "ws": ws, "cols": cols})

        lead = _next_lead(ws, cols)
        if not lead:
            return jsonify({"error": "No more leads — all done!"}), 200

        total = sum(1 for row in ws.iter_rows(min_row=2) if str(row[cols["url"]].value or "").strip())
        sent  = sum(1 for row in ws.iter_rows(min_row=2)
                    if cols["contacted"] is not None and
                    str(row[cols["contacted"]].value or "").strip().lower() in ("yes","y","sent","1","true"))

        return jsonify({"lead": lead, "total": total, "sent": sent})
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/auto_next", methods=["POST"])
def auto_next():
    """Open next lead in Chrome, generate DM, type it into Facebook."""
    try:

        ws   = _auto_state["ws"]
        cols = _auto_state["cols"]
        if ws is None:
            return jsonify({"error": "Load your Excel file first."}), 400

        # If we still have a current_row that was never marked sent/skip,
        # mark it as Skip now so it never resurfaces as a duplicate.
        prev_row = _auto_state["current_row"]
        if prev_row is not None and cols["contacted"] is not None:
            ws.cell(row=prev_row, column=cols["contacted"] + 1, value="Skip")
            _save_wb()

        lead = _next_lead(ws, cols, skip_row=prev_row)
        if not lead:
            return jsonify({"error": "No more leads — all done!"}), 200

        _auto_state["current_row"] = lead["row"]

        # Open profile in Chrome
        from playwright.sync_api import sync_playwright
        import socket
        connected = False
        for _ in range(5):
            try:
                s = socket.create_connection(("127.0.0.1", CDP_PORT), timeout=2)
                s.close()
                connected = True
                break
            except:
                import time as _t2; _t2.sleep(1)
        if not connected:
            return jsonify({"error": "Chrome not found. Close all Chrome windows, double-click START.bat, wait for Chrome to fully open, then try again."}), 500

        with sync_playwright() as p:
            browser = p.chromium.connect_over_cdp(f"http://localhost:{CDP_PORT}")
            ctx     = browser.contexts[0]

            # Close ALL old FB/Messenger tabs before opening the new lead. The old
            # filter excluded messenger.com/"/messages" tabs from cleanup, so any
            # lead where clicking Message navigated that tab into a Messenger
            # conversation left it open forever — after ~10 leads that's ~10 live
            # Messenger tabs (each a heavy React/WebSocket app) piling up, which is
            # what made things grind to a crawl a few leads in. By the time we're
            # asking for the next lead the previous conversation is already sent,
            # so it's always safe to close it.
            for _pg in list(ctx.pages):
                _u = _pg.url.lower()
                if "localhost" in _u:
                    continue
                if "facebook.com" in _u or "messenger.com" in _u or "m.me/" in _u:
                    try:
                        _pg.close()
                        print(f"  Closed old tab: {_u[:70]}")
                    except Exception:
                        pass

            page = ctx.new_page()
            page.bring_to_front()
            import time as _perf
            _t_load = _perf.time()
            try:
                page.goto(lead["url"], wait_until="domcontentloaded", timeout=45000)
            except Exception:
                try:
                    page.goto(lead["url"], wait_until="commit", timeout=30000)
                except Exception as e:
                    return jsonify({"error": f"Could not load page: {e}"}), 500
            print(f"  [timing] page load: {_perf.time() - _t_load:.1f}s")
            page.wait_for_timeout(500)   # settle after load (image-render wait below protects the screenshot)

            # Strip whatever overlay is dimming the page (a dialog or a full-screen
            # scrim). It dims the screenshot the vision model reads and can sit
            # invisibly on top of the Message button.
            try:
                page.keyboard.press("Escape")
            except Exception:
                pass
            _strip_overlays(page)

            # Extract text
            try:
                text = page.evaluate("""() => {
                    return document.body ? document.body.innerText : '';
                }""")
                import re as _re
                text = _re.sub(r"\n{3,}", "\n\n", text).strip()
                text = _re.sub(r"(?im)^\s*[\d,\.]+\s*(followers?|likes?|reviews?|ratings?|check.?ins?|people follow|people like).*$", "", text)
                text = _re.sub(r"\n{3,}", "\n\n", text).strip()[:12000]
            except:
                text = ""

            text_data = text
            if lead["niche"]: text_data = f"--- TRADE TYPE ---\n{lead['niche']}\n\n" + text_data
            if lead["name"]:  text_data = f"--- COMPANY NAME ---\n{lead['name']}\n\n" + text_data
            if lead["city"]:  text_data += f"\n\n--- LISTED CITY ---\n{lead['city']}"

            # Extract name — only from the company name or a bio self-intro,
            # validated against FIRST_NAMES. No name found → plain "Hi." (no guess).
            name_hint = extract_first_name(text or "", lead.get("name", ""))

            _auto_state["last_screenshot"] = None   # cleared per-lead; set below if we get one

            if MESSAGE_VARIANT == "funny":
                # No scan at all: the funny message is built from the SHEET
                # (trade type + city), so there's nothing on the page we need.
                # Skips the screenshot and the vision call — that's the ~8s AI
                # opener plus image capture gone, per lead. The real message is
                # swapped in just below; this is only a safe fallback in case
                # the funny build fails (e.g. a missing photo file).
                dm = _dm_case(tidy_message(pick_stripped_message(
                    text_data, name_hint,
                    lead.get("name", "") or extract_company_name(text_data))))
            elif MESSAGE_VARIANT == "stripped":
                # Direct one-line variant — no opener, so skip the screenshot AND
                # the vision call entirely (faster, and the photo isn't used).
                dm = _dm_case(tidy_message(pick_stripped_message(
                    text_data, name_hint,
                    lead.get("name", "") or extract_company_name(text_data))))
            else:
                # Niches with no finished job to photograph (dentists, movers,
                # storage) skip the whole visual pipeline: no post screenshot,
                # no vision call, no profile screenshot. That is the entire
                # per-lead wait — a minute or more once the vision retry kicks
                # in — spent producing an opener that the postprocess then
                # discarded, because nothing a model invents off a photo of a
                # van or a waiting room can be corroborated against the page.
                # The opener still gets personalised, just from the page TEXT:
                # CUSTOM_OPENER_POSTPROCESS normally runs inside call_gemini
                # (see its body), so on this path it has to be invoked directly.
                skip_vision = bool(SKIP_VISION_FOR and SKIP_VISION_FOR(text_data))

                if skip_vision:
                    print("  Vision SKIPPED for this niche — opener from page text (no screenshot, no model call).")
                    opener = (CUSTOM_OPENER_POSTPROCESS("", text_data) or "") if CUSTOM_OPENER_POSTPROCESS else ""
                    profile_shot = None
                else:
                    # Screenshot the first-post area so the vision model can see the actual job photos
                    images = []
                    shot = None
                    try:
                        import time as _tw
                        _scroll_to_freshest_post(page)
                        _tw.sleep(0.8)   # wait for images to render
                        shot = page.screenshot(type="jpeg", quality=72)
                        images = [(shot, "image/jpeg")]
                        print(f"  First-post screenshot: {len(shot)} bytes")
                    except Exception as _se:
                        print(f"  Screenshot skipped: {_se}")

                    # Generate DM — pass screenshot so vision model sees the actual job
                    import time as _perf2
                    _t_ai = _perf2.time()
                    opener = call_gemini(text_data=text_data, images=images if images else None)
                    print(f"  [timing] AI opener: {_perf2.time() - _t_ai:.1f}s")
                    opener = patch_fallback_opener(opener, text_data, lead.get("city", ""))

                # Picture to attach: always the top of their profile — logo,
                # cover photo, business name — not whatever post happens to be
                # freshest. A mid-feed post can be a text-only update or a
                # boring link-share card (as opposed to an actual job photo),
                # which reads as unclear/generic when sent as "the picture."
                # The header is consistently recognizable as "this is your
                # business page" regardless of what they last posted.
                # Skipped on the no-vision path: it costs a 900ms settle plus
                # the capture, and its only consumers are the "Copy picture"
                # button and the auto-paste, which HU disables anyway
                # (HU_ATTACH_SCREENSHOT=False in dm_app_hu.py). The button just
                # stays hidden for these leads, since has_screenshot is False.
                if not skip_vision:
                    try:
                        page.evaluate("window.scrollTo(0, 0)")
                        _strip_overlays(page)   # re-run fresh — the early cleanup's 6s watcher has long since given up by now
                        page.wait_for_timeout(900)   # was 400ms — too short, sometimes caught the page mid-render
                        # Screenshot just the main content column, not the full
                        # viewport. A full-page shot on desktop Facebook includes
                        # the right-hand contacts/chat sidebar — which shows every
                        # OTHER business page already messaged, visible to
                        # whoever this picture gets sent to. [role="main"] is
                        # Facebook's own landmark for the profile content only,
                        # so this crops the sidebar (and left nav) out entirely
                        # rather than guessing a pixel width to cut.
                        main = page.locator('[role="main"]').first
                        box = main.bounding_box() if main.count() > 0 and main.is_visible() else None
                        if box:
                            # element.screenshot() captures the FULL scrollable height
                            # of [role="main"], not just what's on screen — Facebook's
                            # main container extends way past the visible page, which
                            # produced a huge mostly-blank image. Clip to the element's
                            # x/width (still excludes the sidebar, which sits outside
                            # role="main" horizontally) but cap the height to the
                            # actual viewport instead of the element's full height.
                            viewport = page.viewport_size or {"width": 1280, "height": 900}
                            clip_h = max(200, min(box["height"], viewport["height"] - box["y"]))
                            profile_shot = page.screenshot(type="jpeg", quality=72, clip={
                                "x": box["x"], "y": box["y"], "width": box["width"], "height": clip_h,
                            })
                            print(f"  Profile screenshot (clipped to main, viewport height): {len(profile_shot)} bytes")
                        else:
                            profile_shot = page.screenshot(type="jpeg", quality=72)
                            print(f"  [role=main] not found — full-page screenshot (may include sidebar): {len(profile_shot)} bytes")
                        _auto_state["last_screenshot"] = (profile_shot, "image/jpeg")
                    except Exception as _se:
                        print(f"  Picture capture skipped: {_se}")
                        profile_shot = None
                # Reuse the same profile screenshot to read the actual logo
                # colour off it, if there is a clean one to read — no extra
                # screenshot, one bounded model call, "" (never a guess) if
                # the picture isn't a clear single-colour logo. Cropped to
                # the top of the page first: sending the WHOLE main-content
                # shot made the small circular avatar a tiny fraction of a
                # busy image, which is why obvious logos were still coming
                # back "none" — the cover-photo-plus-avatar area sits within
                # the top ~700px on every FB Page layout, so cropping there
                # makes the logo a much bigger, clearer part of what the
                # model actually sees.
                brand_color = ""
                if profile_shot:
                    try:
                        from PIL import Image
                        import io as _io
                        img = Image.open(_io.BytesIO(profile_shot))
                        top_crop = img.crop((0, 0, img.width, min(img.height, 700)))
                        buf = _io.BytesIO()
                        top_crop.convert("RGB").save(buf, format="JPEG", quality=90)
                        brand_color = detect_brand_color(buf.getvalue())
                        if brand_color:
                            print(f"  Brand colour detected: {brand_color}")
                    except Exception as _ce:
                        print(f"  Brand colour detection skipped: {_ce}")
                # Which market's wording to use — detected from the loaded
                # leads file, matching the US/UK buttons in the UI, so no
                # extra manual toggle is needed here.
                _excel_path = (_auto_state.get("excel_path") or "").lower()
                market = "UK" if "uk" in _excel_path else "US"
                opener = _apply_greeting(opener, name_hint, market)
                if MESSAGE_VARIANT == "british":
                    company_name = lead.get("name", "") or extract_company_name(text_data)
                    dm = _dm_case(tidy_message(pick_british_message(
                        opener, text_data, name_hint, company_name, brand_color, market=market)))
                else:
                    opener = opener.rstrip(".!,") + "."
                    trade  = detect_trade(text_data)
                    body   = pick_body(trade, text_data, lead.get("city", ""), lead.get("name", "") or extract_company_name(text_data))
                    dm     = _dm_case(tidy_message((opener + " " + body).strip()))

            # ── "Funny" style: swap in a joke message + its photo ─────────────
            # Done here, AFTER the normal DM is built, so a missing image file
            # just leaves the personalised message in place rather than sending
            # a punchline with nothing to look at.
            funny_key = ""
            if MESSAGE_VARIANT == "funny":
                # ONLY name_hint — that's already been validated against
                # FIRST_NAMES upstream. lead["name"] is the COMPANY name, so
                # falling back to it turns "Top Loft Ltd" into "alright Top?".
                _fname = locals().get("name_hint", "")
                _niche = lead.get("niche", "") or locals().get("trade", "")
                built  = build_funny_message(lead["row"], _fname, _niche,
                                             lead.get("city", ""))
                if built:
                    dm, _img, funny_key = built
                    # Replaces the profile screenshot — the joke photo IS the
                    # attachment for these, and only one image gets pasted.
                    _auto_state["last_screenshot"] = _img
                    print(f"  FUNNY message '{funny_key}' (row {lead['row']}, "
                          f"niche -> {niche_words(_niche)})")
                else:
                    print("  Funny message unavailable — kept personalised DM.")

            # Click Message button to open the chat
            msg_opened = _click_message_btn(page)
            picture_attached = False
            if msg_opened:
                shot_data = _auto_state.get("last_screenshot")
                if shot_data:
                    # Do this BEFORE copying the DM text — pasting the picture
                    # uses the clipboard internally (sets it to the image),
                    # which would silently overwrite the DM text if that had
                    # already been copied. Text goes to the clipboard last, so
                    # it's what's actually sitting there for the user to paste.
                    picture_attached = _paste_image_into_box(page, shot_data[0])
                    if picture_attached:
                        print(f"  Picture attached directly into message box "
                              f"({len(shot_data[0]):,} bytes, {shot_data[1]}"
                              f"{', FUNNY photo' if funny_key else ', profile screenshot'}).")
                    else:
                        print("  Could not attach picture automatically — use the Copy picture button.")
            else:
                print("  Could not click Message button — user can open chat manually.")

            # DM text is always just clipboard-copied — the user handles it
            # themselves (paste or type it in). Only the PICTURE gets attached
            # automatically, since that's a pure clipboard-set + paste with no
            # risk of landing in the wrong field the way typing text did.
            clipboard = copy_to_clipboard(dm)

        return jsonify({
            "dm":         dm,
            "lead":       lead,
            "clipboard":  clipboard,
            "msg_opened": msg_opened,
            "picture_attached": picture_attached,
            "funny_key":  funny_key,
            "sent_today": _auto_state["sent_today"],
            "has_screenshot": _auto_state["last_screenshot"] is not None,
        })

    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/auto_copy_screenshot", methods=["POST"])
def auto_copy_screenshot():
    """Copy the current lead's picture (set by /auto_next) to the clipboard,
    ready to Ctrl+V into the Messenger chat as a second message. Separate
    button/call from the DM generation on purpose — see copy_image_to_clipboard()
    for why text and image can't share one clipboard operation here."""
    shot = _auto_state.get("last_screenshot")
    if not shot:
        return jsonify({"error": "No picture for this lead — it may have fallen back to a generic profile screenshot that wasn't captured, or the page failed to load."}), 400
    image_bytes, _mime = shot
    ok = copy_image_to_clipboard(image_bytes)
    if not ok:
        return jsonify({"error": "Could not copy the picture to the clipboard."}), 500
    return jsonify({"ok": True})


@app.route("/auto_skip", methods=["POST"])
def auto_skip():
    """Skip current lead (mark as Skip) and return next lead info without opening browser."""
    try:
        wb   = _auto_state["wb"]
        ws   = _auto_state["ws"]
        cols = _auto_state["cols"]
        row  = _auto_state["current_row"]

        if ws is None:
            return jsonify({"error": "No Excel loaded."}), 400

        if row is not None and cols["contacted"] is not None:
            ws.cell(row=row, column=cols["contacted"] + 1, value="Skip")
            _save_wb()

        _auto_state["current_row"] = None

        lead = _next_lead(ws, cols)
        if not lead:
            return jsonify({"error": "No more leads — all done!"}), 200

        _auto_state["current_row"] = lead["row"]

        total = sum(1 for r in ws.iter_rows(min_row=2) if str(r[cols["url"]].value or "").strip())
        sent  = sum(1 for r in ws.iter_rows(min_row=2)
                    if cols["contacted"] is not None and
                    str(r[cols["contacted"]].value or "").strip().lower() in ("yes","y","sent","1","true"))

        return jsonify({"ok": True, "lead": lead, "sent": sent, "total": total})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/auto_mark_sent", methods=["POST"])
def auto_mark_sent():
    """Mark current lead as sent in Excel."""
    try:
        wb   = _auto_state["wb"]
        ws   = _auto_state["ws"]
        cols = _auto_state["cols"]
        row  = _auto_state["current_row"]

        if ws is None or row is None:
            return jsonify({"ok": False, "error": "Nothing to mark."}), 400

        if cols["contacted"] is not None:
            ws.cell(row=row, column=cols["contacted"] + 1, value="Yes")
        if cols.get("sent_date") is not None:
            ws.cell(row=row, column=cols["sent_date"] + 1, value=_today_str())
        _save_wb()

        _auto_state["sent_today"] += 1
        _auto_state["current_row"] = None

        sent  = sum(1 for r in ws.iter_rows(min_row=2)
                    if cols["contacted"] is not None and
                    str(r[cols["contacted"]].value or "").strip().lower() in ("yes","y","sent","1","true"))
        total = sum(1 for r in ws.iter_rows(min_row=2) if str(r[cols["url"]].value or "").strip())

        return jsonify({"ok": True, "sent_today": _auto_state["sent_today"], "sent": sent, "total": total})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


# ── FB Group Sender section ───────────────────────────────────────────────────
# A third, separate flow from Auto Sender. Auto Sender's whole job is to
# INVENT the personalised opener (screenshot the profile, ask the vision
# model). This flow is for leads collected by hand off Facebook groups
# (fb_group_scraper.py, or copied in manually) where the opener already
# exists as a column in the sheet — Milán read the post himself and typed the
# line. So this never screenshots, never calls Groq/Gemini: it opens the
# profile, drops the sheet's own line through the same greeting + pitch
# pipeline as every other flow (_apply_greeting/pick_body — HU vs EN resolved
# the same way, via dm_app_hu.py's module-level overrides), and copies the
# result to the clipboard. Own state dict so it can't collide with a
# different sheet loaded in Auto Sender or Owner Emails at the same time.
_fbgroup_state = {
    "excel_path":  None,
    "current_row": None,
    "sent_today":  0,
    "wb":          None,
    "ws":          None,
    "cols":        {},
}


def _load_fb_excel(path: str):
    """_load_excel plus a 'Personalized Line' column — the one thing this flow
    needs that no other sheet has.

    Delegates the tracking-column resolution to _load_excel rather than
    repeating its matcher list, which is how this function originally shipped:
    the duplicated predicates meant the Status/Followers mis-binding fixed in
    _load_excel would have quietly survived here."""
    wb, ws, cols = _load_excel(path)
    headers = [str(c.value or "").strip().lower() for c in ws[1]]
    # "Personalized Line" / "Personal line" / "Opener" — whatever he calls the
    # column he writes the opener into. Ordered like _load_excel's matchers so
    # an explicitly named column beats a loose "line" substring hit.
    cols["line"] = next(
        (i for i, h in enumerate(headers) if "personal" in h and "line" in h), None)
    if cols["line"] is None:
        cols["line"] = next(
            (i for i, h in enumerate(headers)
             if "personal" in h or "opener" in h or h.strip() == "line"), None)
    return wb, ws, cols


def _next_fb_lead(ws, cols, skip_row=None):
    """Next uncontacted FB-group lead that actually HAS a personalised line —
    a row with no line yet isn't ready to send (Milán hasn't written it in),
    so it's skipped rather than sent with nothing personal in it. Never
    dropped silently from the sheet, just not surfaced until the column is
    filled in."""
    for row in ws.iter_rows(min_row=2):
        excel_row = row[0].row
        if skip_row is not None and excel_row == skip_row:
            continue
        url = str(row[cols["url"]].value or "").strip()
        if not url.startswith("http"):
            continue
        if "facebook.com" not in url.lower() and "instagram.com" not in url.lower():
            continue
        line = str(row[cols["line"]].value or "").strip() if cols["line"] is not None else ""
        if not line:
            continue
        contacted = str(row[cols["contacted"]].value or "").strip().lower() if cols["contacted"] is not None else ""
        if contacted not in ("yes", "y", "sent", "1", "true", "skip"):
            return {
                "row":   excel_row,
                "url":   url,
                "name":  str(row[cols["name"]].value or "").strip(),
                "city":  str(row[cols["city"]].value or "").strip()  if cols["city"]  is not None else "",
                "niche": str(row[cols["niche"]].value or "").strip() if cols["niche"] is not None else "",
                "line":  line,
            }
    return None


def _save_fb_wb():
    import time as _t
    wb   = _fbgroup_state.get("wb")
    path = _fbgroup_state.get("excel_path")
    if wb is None or not path:
        return
    for attempt in range(6):
        try:
            wb.save(path)
            return
        except PermissionError:
            _t.sleep(0.4)
    raise PermissionError(
        "Can't write to the Excel file — it's open in Excel (or still syncing in "
        "OneDrive). Close the spreadsheet in Excel, wait for OneDrive to finish, "
        "then click again."
    )


@app.route("/fb_load", methods=["POST"])
def fb_load():
    try:
        data = request.get_json()
        path = data.get("path", "").strip()

        if not path:
            for f in _Path(__file__).parent.glob("*.xlsx"):
                nm = f.name.lower()
                if any(k in nm for k in ("fb_group", "group", "fb_lead")):
                    path = str(f); break

        if not path or not _Path(path).exists():
            return jsonify({"error": f"Excel file not found: {path}"}), 400

        wb, ws, cols = _load_fb_excel(path)
        if cols["line"] is None:
            return jsonify({"error": "No 'Personalized Line' column found in this sheet — "
                                      "add one and put your opener line for each lead in it."}), 400
        _fbgroup_state.update({"excel_path": path, "wb": wb, "ws": ws, "cols": cols})

        lead = _next_fb_lead(ws, cols)
        if not lead:
            return jsonify({"error": "No more leads with a personalized line ready — all done, "
                                      "or add lines for the rest in the sheet."}), 200

        total = sum(1 for row in ws.iter_rows(min_row=2) if str(row[cols["url"]].value or "").strip())
        sent  = sum(1 for row in ws.iter_rows(min_row=2)
                    if cols["contacted"] is not None and
                    str(row[cols["contacted"]].value or "").strip().lower() in ("yes","y","sent","1","true"))

        return jsonify({"lead": lead, "total": total, "sent": sent})
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/fb_next", methods=["POST"])
def fb_next():
    """Open next lead in Chrome and build the DM from the sheet's own
    personalised line — no screenshot, no AI call, just the greeting +
    pitch wrapped around what Milán already wrote."""
    try:
        ws   = _fbgroup_state["ws"]
        cols = _fbgroup_state["cols"]
        if ws is None:
            return jsonify({"error": "Load your Excel file first."}), 400

        prev_row = _fbgroup_state["current_row"]
        if prev_row is not None and cols["contacted"] is not None:
            ws.cell(row=prev_row, column=cols["contacted"] + 1, value="Skip")
            _save_fb_wb()

        lead = _next_fb_lead(ws, cols, skip_row=prev_row)
        if not lead:
            return jsonify({"error": "No more leads with a personalized line ready — all done, "
                                      "or add lines for the rest in the sheet."}), 200

        _fbgroup_state["current_row"] = lead["row"]

        from playwright.sync_api import sync_playwright
        import socket
        connected = False
        for _ in range(5):
            try:
                s = socket.create_connection(("127.0.0.1", CDP_PORT), timeout=2)
                s.close()
                connected = True
                break
            except:
                import time as _t2; _t2.sleep(1)
        if not connected:
            return jsonify({"error": "Chrome not found. Close all Chrome windows, double-click START.bat, wait for Chrome to fully open, then try again."}), 500

        with sync_playwright() as p:
            browser = p.chromium.connect_over_cdp(f"http://localhost:{CDP_PORT}")
            ctx     = browser.contexts[0]

            for _pg in list(ctx.pages):
                _u = _pg.url.lower()
                if "localhost" in _u:
                    continue
                if "facebook.com" in _u or "messenger.com" in _u or "m.me/" in _u:
                    try:
                        _pg.close()
                    except Exception:
                        pass

            page = ctx.new_page()
            page.bring_to_front()
            try:
                page.goto(lead["url"], wait_until="domcontentloaded", timeout=45000)
            except Exception:
                try:
                    page.goto(lead["url"], wait_until="commit", timeout=30000)
                except Exception as e:
                    return jsonify({"error": f"Could not load page: {e}"}), 500
            page.wait_for_timeout(500)

            try:
                page.keyboard.press("Escape")
            except Exception:
                pass
            _strip_overlays(page)

            # Text scrape is still worth doing — it's free (page is already
            # open to click Message anyway) and feeds name detection + niche
            # detection the same way Auto Sender's does. What it's NOT used
            # for is generating the opener — that's the sheet's "line".
            try:
                text = page.evaluate("""() => {
                    return document.body ? document.body.innerText : '';
                }""")
                text = re.sub(r"\n{3,}", "\n\n", text).strip()
                text = re.sub(r"(?im)^\s*[\d,\.]+\s*(followers?|likes?|reviews?|ratings?|check.?ins?|people follow|people like).*$", "", text)
                text = re.sub(r"\n{3,}", "\n\n", text).strip()[:12000]
            except:
                text = ""

            text_data = text
            if lead["niche"]: text_data = f"--- TRADE TYPE ---\n{lead['niche']}\n\n" + text_data
            if lead["name"]:  text_data = f"--- COMPANY NAME ---\n{lead['name']}\n\n" + text_data
            if lead["city"]:  text_data += f"\n\n--- LISTED CITY ---\n{lead['city']}"

            name_hint = extract_first_name(text or "", lead.get("name", ""))

            _excel_path = (_fbgroup_state.get("excel_path") or "").lower()
            market = "UK" if "uk" in _excel_path else "US"
            opener = lead["line"].strip()
            opener = _apply_greeting(opener, name_hint, market)
            opener = opener.rstrip(".!,") + "."
            trade  = detect_trade(text_data)
            body   = pick_body(trade, text_data, lead.get("city", ""), lead.get("name", "") or extract_company_name(text_data))
            dm     = _dm_case(tidy_message((opener + " " + body).strip()))

            msg_opened = _click_message_btn(page)
            if not msg_opened:
                print("  Could not click Message button — user can open chat manually.")

            clipboard = copy_to_clipboard(dm)

        return jsonify({
            "dm":         dm,
            "lead":       lead,
            "clipboard":  clipboard,
            "msg_opened": msg_opened,
            "sent_today": _fbgroup_state["sent_today"],
        })

    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/fb_skip", methods=["POST"])
def fb_skip():
    try:
        wb   = _fbgroup_state["wb"]
        ws   = _fbgroup_state["ws"]
        cols = _fbgroup_state["cols"]
        row  = _fbgroup_state["current_row"]

        if ws is None:
            return jsonify({"error": "No Excel loaded."}), 400

        if row is not None and cols["contacted"] is not None:
            ws.cell(row=row, column=cols["contacted"] + 1, value="Skip")
            _save_fb_wb()

        _fbgroup_state["current_row"] = None

        lead = _next_fb_lead(ws, cols)
        if not lead:
            return jsonify({"error": "No more leads with a personalized line ready — all done, "
                                      "or add lines for the rest in the sheet."}), 200

        _fbgroup_state["current_row"] = lead["row"]

        total = sum(1 for r in ws.iter_rows(min_row=2) if str(r[cols["url"]].value or "").strip())
        sent  = sum(1 for r in ws.iter_rows(min_row=2)
                    if cols["contacted"] is not None and
                    str(r[cols["contacted"]].value or "").strip().lower() in ("yes","y","sent","1","true"))

        return jsonify({"ok": True, "lead": lead, "sent": sent, "total": total})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/fb_mark_sent", methods=["POST"])
def fb_mark_sent():
    try:
        wb   = _fbgroup_state["wb"]
        ws   = _fbgroup_state["ws"]
        cols = _fbgroup_state["cols"]
        row  = _fbgroup_state["current_row"]

        if ws is None or row is None:
            return jsonify({"ok": False, "error": "Nothing to mark."}), 400

        if cols["contacted"] is not None:
            ws.cell(row=row, column=cols["contacted"] + 1, value="Yes")
        if cols.get("sent_date") is not None:
            ws.cell(row=row, column=cols["sent_date"] + 1, value=_today_str())
        _save_fb_wb()

        _fbgroup_state["sent_today"] += 1
        _fbgroup_state["current_row"] = None

        sent  = sum(1 for r in ws.iter_rows(min_row=2)
                    if cols["contacted"] is not None and
                    str(r[cols["contacted"]].value or "").strip().lower() in ("yes","y","sent","1","true"))
        total = sum(1 for r in ws.iter_rows(min_row=2) if str(r[cols["url"]].value or "").strip())

        return jsonify({"ok": True, "sent_today": _fbgroup_state["sent_today"], "sent": sent, "total": total})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


# ── Owner Emails section ──────────────────────────────────────────────────────
# A separate, self-contained flow for leads whose contact is an email address.
# It scans the lead's Company Facebook page for a personal line, writes one short
# owner email, and copies it to the clipboard. It uses its own _email_state so it
# never disturbs the Auto Sender (DM) flow.

def _save_email_wb():
    """Save the Owner-Emails workbook, retrying briefly if the file is locked."""
    import time as _t
    wb   = _email_state.get("wb")
    path = _email_state.get("excel_path")
    if wb is None or not path:
        return
    for _ in range(6):
        try:
            wb.save(path); return
        except PermissionError:
            _t.sleep(0.4)
    raise PermissionError(
        "Can't write to the Excel file — it's open in Excel (or still syncing in "
        "OneDrive). Close the spreadsheet, wait for OneDrive, then click again.")


def _email_counts(ws, cols):
    total = sum(
        1 for r in ws.iter_rows(min_row=2)
        if (cols.get("contact") is not None and "@" in str(r[cols["contact"]].value or ""))
        or (cols.get("contact_type") is not None and "email" in str(r[cols["contact_type"]].value or "").lower()))
    sent = sum(
        1 for r in ws.iter_rows(min_row=2)
        if cols["contacted"] is not None and
        str(r[cols["contacted"]].value or "").strip().lower() in ("yes", "y", "sent", "1", "true"))
    return total, sent


@app.route("/email_load", methods=["POST"])
def email_load():
    """Load Excel for the Owner Emails section; return the first email lead."""
    try:
        data = request.get_json() or {}
        path = (data.get("path") or "").strip()
        if not path:
            for f in _Path(__file__).parent.glob("*.xlsx"):
                if any(k in f.name.lower() for k in ("lead", "email", "contact", "prospect", "outreach")):
                    path = str(f); break
        if not path or not _Path(path).exists():
            return jsonify({"error": f"Excel file not found: {path}"}), 400
        wb, ws, cols = _load_excel(path)
        _email_state.update({"excel_path": path, "wb": wb, "ws": ws, "cols": cols, "current_row": None})
        lead = _next_email_lead(ws, cols)
        if not lead:
            return jsonify({"error": "No email leads found in this file."}), 200
        total, sent = _email_counts(ws, cols)
        return jsonify({"lead": lead, "total": total, "sent": sent})
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500


def _gmail_compose_url(to: str, subject: str, body: str) -> str:
    """The Gmail compose deep-link (full-screen compose, everything pre-filled)."""
    from urllib.parse import quote
    return ("https://mail.google.com/mail/?view=cm&fs=1"
            f"&to={quote(to or '')}"
            f"&su={quote(subject or '')}"
            f"&body={quote(body or '')}")


def _open_gmail_tab(to: str, subject: str, body: str):
    """Open the pre-filled Gmail compose in the user's DEFAULT browser (their
    normal Chrome) — NOT the isolated DMSender automation Chrome.

    Why default browser: Streak (and any email tracker) is a Gmail extension that
    only injects its tracking pixel in the Chrome where it's installed + logged in
    — i.e. the real, everyday Chrome. The DMSender profile is a clean, isolated
    instance with no extensions, so anything composed there sends UNTRACKED. By
    handing the compose URL to the OS default browser, Gmail (and Streak) open in
    the right place and the send is tracked. Best-effort: a failure here never
    breaks writing the email."""
    url = _gmail_compose_url(to, subject, body)
    try:
        import webbrowser
        webbrowser.open(url, new=2)   # new=2 → new tab in the running default browser
    except Exception as _ge:
        print(f"  Gmail auto-open skipped: {_ge}")


@app.route("/email_next", methods=["POST"])
def email_next():
    """Scan the next email lead's Company FB page and write the owner email."""
    try:
        ws   = _email_state["ws"]
        cols = _email_state["cols"]
        if ws is None:
            return jsonify({"error": "Load your Excel file first."}), 400

        prev_row = _email_state["current_row"]
        if prev_row is not None and cols["contacted"] is not None:
            ws.cell(row=prev_row, column=cols["contacted"] + 1, value="Skip")
            _save_email_wb()

        lead = _next_email_lead(ws, cols, skip_row=prev_row)
        if not lead:
            return jsonify({"error": "No more email leads — all done!"}), 200
        _email_state["current_row"] = lead["row"]

        scan_url = (lead.get("url") or "").strip()

        # No Company FB page → write from spreadsheet data only, no Chrome.
        if not scan_url.startswith("http"):
            subject, body = _build_email("", "", lead)
            clipboard = copy_to_clipboard(_email_clipboard(lead.get("email", ""), subject, body))
            _open_gmail_tab(lead.get("email", ""), subject, body)
            return jsonify({"subject": subject, "email": lead.get("email", ""),
                            "body": body, "lead": lead, "clipboard": clipboard,
                            "sent_today": _email_state["sent_today"]})

        from playwright.sync_api import sync_playwright
        import socket
        connected = False
        for _ in range(5):
            try:
                s = socket.create_connection(("127.0.0.1", CDP_PORT), timeout=2); s.close()
                connected = True; break
            except:
                import time as _t2; _t2.sleep(1)
        if not connected:
            return jsonify({"error": "Chrome not found. Double-click START UK.bat, wait for Chrome to open, then try again."}), 500

        with sync_playwright() as p:
            browser = p.chromium.connect_over_cdp(f"http://localhost:{CDP_PORT}")
            ctx     = browser.contexts[0]
            for _pg in list(ctx.pages):
                _u = _pg.url.lower()
                if "localhost" in _u:
                    continue
                if "facebook.com" in _u or "messenger.com" in _u or "m.me/" in _u:
                    try: _pg.close()
                    except Exception: pass
            page = ctx.new_page()
            page.bring_to_front()
            try:
                page.goto(scan_url, wait_until="domcontentloaded", timeout=45000)
            except Exception:
                try:
                    page.goto(scan_url, wait_until="commit", timeout=30000)
                except Exception as e:
                    return jsonify({"error": f"Could not load page: {e}"}), 500
            page.wait_for_timeout(1000)
            try: page.keyboard.press("Escape")
            except Exception: pass
            try:
                page.evaluate(r"""() => {
                    if (!document.getElementById('dm-overlay-killer')) {
                        const st = document.createElement('style'); st.id = 'dm-overlay-killer';
                        st.textContent = '[role="dialog"]{display:none!important;}html,body{overflow:auto!important;}';
                        (document.head || document.documentElement).appendChild(st);
                    }
                    const nuke = () => document.querySelectorAll('[role="dialog"]').forEach(el => el.remove());
                    nuke(); const obs = new MutationObserver(nuke);
                    obs.observe(document.documentElement, { childList: true, subtree: true });
                    setTimeout(() => obs.disconnect(), 6000);
                }""")
            except Exception: pass
            try:
                text = page.evaluate("""() => {
                    return document.body ? document.body.innerText : '';
                }""")
                import re as _re
                text = _re.sub(r"\n{3,}", "\n\n", text).strip()[:12000]
            except:
                text = ""

            text_data = text
            if lead["niche"]: text_data = f"--- TRADE TYPE ---\n{lead['niche']}\n\n" + text_data
            if lead["city"]:  text_data += f"\n\n--- LISTED CITY ---\n{lead['city']}"

            opener = ""
            try:
                import time as _tw
                _scroll_to_freshest_post(page)
                _tw.sleep(1.4)
                shot = page.screenshot(type="jpeg", quality=72)
                page.evaluate("window.scrollTo(0, 0)")
                raw = call_gemini(text_data=text_data, images=[(shot, "image/jpeg")])
                opener = patch_fallback_opener(raw.rstrip(".!,") + ".", text_data, lead.get("city", ""))
            except Exception as _se:
                print(f"  Email opener skipped: {_se}")

            try:
                page_title = page.evaluate(
                    "() => { const h = document.querySelector('h1');"
                    " const o = document.querySelector('meta[property=\"og:title\"]');"
                    " return (h && h.innerText) || (o && o.content) || document.title || ''; }")
            except Exception:
                page_title = ""

            subject, body = _build_email(opener, text, lead, page_title)
            clipboard = copy_to_clipboard(_email_clipboard(lead.get("email", ""), subject, body))
            # Email is ready → open Gmail compose automatically in the user's
            # DEFAULT Chrome (where Streak lives), so the send is tracked.
            _open_gmail_tab(lead.get("email", ""), subject, body)

        return jsonify({"subject": subject, "email": lead.get("email", ""),
                        "body": body, "lead": lead, "clipboard": clipboard,
                        "sent_today": _email_state["sent_today"]})
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/email_skip", methods=["POST"])
def email_skip():
    """Skip the current email lead and return the next one (no Chrome)."""
    try:
        ws = _email_state["ws"]; cols = _email_state["cols"]; row = _email_state["current_row"]
        if ws is None:
            return jsonify({"error": "No Excel loaded."}), 400
        if row is not None and cols["contacted"] is not None:
            ws.cell(row=row, column=cols["contacted"] + 1, value="Skip"); _save_email_wb()
        _email_state["current_row"] = None
        lead = _next_email_lead(ws, cols)
        if not lead:
            return jsonify({"error": "No more email leads — all done!"}), 200
        _email_state["current_row"] = lead["row"]
        total, sent = _email_counts(ws, cols)
        return jsonify({"ok": True, "lead": lead, "sent": sent, "total": total})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/email_mark_sent", methods=["POST"])
def email_mark_sent():
    """Mark the current email lead as sent."""
    try:
        ws = _email_state["ws"]; cols = _email_state["cols"]; row = _email_state["current_row"]
        if ws is None or row is None:
            return jsonify({"ok": False, "error": "Nothing to mark."}), 400
        if cols["contacted"] is not None:
            ws.cell(row=row, column=cols["contacted"] + 1, value="Yes")
        if cols.get("sent_date") is not None:
            ws.cell(row=row, column=cols["sent_date"] + 1, value=_today_str())
        _save_email_wb()
        _email_state["sent_today"] += 1; _email_state["current_row"] = None
        total, sent = _email_counts(ws, cols)
        return jsonify({"ok": True, "sent_today": _email_state["sent_today"], "sent": sent, "total": total})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


# ── Follow-up tracking ───────────────────────────────────────────────────────
# A lead is "due" for a nudge when it's been contacted, hasn't replied, the
# sequence isn't finished (max 3 nudges), and enough days have passed since the
# LAST touch. Gaps widen: nudge #1 at 3 days, #2 at +7, #3 at +11.
_FOLLOWUP_GAP = {0: 3, 1: 7, 2: 11}

@app.route("/followups_due", methods=["POST"])
def followups_due():
    """Return every contacted-but-no-reply lead that's due for a follow-up."""
    try:
        ws   = _auto_state["ws"]
        cols = _auto_state["cols"]
        if ws is None:
            return jsonify({"error": "Load your leads on the Auto Sender tab first."}), 400

        from datetime import date
        today = date.today()
        due = []
        for row in ws.iter_rows(min_row=2):
            r   = row[0].row
            url = str(row[cols["url"]].value or "").strip()
            if not url.startswith("http"):
                continue
            contacted = (str(row[cols["contacted"]].value or "").strip().lower()
                         if cols["contacted"] is not None else "")
            if contacted not in ("yes", "y", "sent", "1", "true"):
                continue  # only leads we've actually messaged
            replied = (str(row[cols["replied"]].value or "").strip().lower()
                       if cols.get("replied") is not None else "")
            if replied in ("yes", "y", "1", "true"):
                continue  # they replied — never nudge
            fups = _to_int(row[cols["followups"]].value) if cols.get("followups") is not None else 0
            if fups >= FOLLOWUP_TOUCHES:
                continue  # sequence finished
            last = _parse_date(row[cols["sent_date"]].value) if cols.get("sent_date") is not None else None
            # No date logged (e.g. sent before date-tracking existed) — we can't
            # tell HOW overdue it is, but it's a contacted-no-reply lead that's
            # been silently dropped. Surface it as due (no_date) rather than
            # hiding it forever. days=None so the UI shows it without a count.
            if last is None:
                due.append({
                    "row":   r,
                    "name":  str(row[cols["name"]].value or "").strip(),
                    "url":   url,
                    "city":  str(row[cols["city"]].value or "").strip() if cols["city"] is not None else "",
                    "niche": str(row[cols["niche"]].value or "").strip() if cols["niche"] is not None else "",
                    "days":  None,
                    "no_date": True,
                    "touch": fups + 1,
                })
                continue
            days = (today - last).days
            if days < _FOLLOWUP_GAP.get(fups, 99):
                continue  # not due yet
            due.append({
                "row":   r,
                "name":  str(row[cols["name"]].value or "").strip(),
                "url":   url,
                "city":  str(row[cols["city"]].value or "").strip() if cols["city"] is not None else "",
                "niche": str(row[cols["niche"]].value or "").strip() if cols["niche"] is not None else "",
                "days":  days,
                "touch": fups + 1,
            })
        # No-date leads first (oldest, silently-dropped), then most overdue.
        due.sort(key=lambda d: (0 if d.get("no_date") else 1, -(d["days"] or 0)))
        return jsonify({"due": due, "count": len(due)})
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500

@app.route("/followup_send", methods=["POST"])
def followup_send():
    """Prepare the next nudge for one lead: copy it to the clipboard, log the
    touch (advance the cadence), and return the text + profile URL to open."""
    try:
        ws   = _auto_state["ws"]
        cols = _auto_state["cols"]
        if ws is None:
            return jsonify({"error": "Load your leads first."}), 400
        r = int(request.get_json(force=True).get("row", 0))
        if r < 2:
            return jsonify({"error": "Bad row."}), 400

        name = str(ws.cell(row=r, column=cols["name"] + 1).value or "").strip()
        url  = str(ws.cell(row=r, column=cols["url"] + 1).value or "").strip()
        niche = (str(ws.cell(row=r, column=cols["niche"] + 1).value or "").strip()
                 if cols.get("niche") is not None else "")
        fups = _to_int(ws.cell(row=r, column=cols["followups"] + 1).value) if cols.get("followups") is not None else 0
        text = _followup_text(fups + 1, name, niche)

        if cols.get("followups") is not None:
            ws.cell(row=r, column=cols["followups"] + 1, value=fups + 1)
        if cols.get("sent_date") is not None:
            ws.cell(row=r, column=cols["sent_date"] + 1, value=_today_str())
        _save_wb()

        clipboard = copy_to_clipboard(text)
        return jsonify({"ok": True, "text": text, "url": url, "clipboard": clipboard, "touch": fups + 1})
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500

@app.route("/replies_pending", methods=["POST"])
def replies_pending():
    """Every contacted lead with no reply logged yet, newest send first, plus
    the reply-rate stats.

    Exists because /followups_due — the only screen that ever offered a "Got a
    reply" button — deliberately hides a lead until its nudge is DUE (3+ days,
    see _FOLLOWUP_GAP). A reply that lands on day 0-2 therefore had nowhere to
    be recorded, and by day 3 the conversation has already happened, so
    clicking the button pays the user nothing in the moment. Measured result:
    387 sends and 0 replies ever logged across the master sheet, which makes
    every copy decision unfalsifiable.

    This route has no due-date filter at all: if it was sent and no answer is
    recorded, it shows up. The stats come back with it so the screen is worth
    opening for its own sake rather than being pure data entry.
    """
    try:
        ws   = _auto_state["ws"]
        cols = _auto_state["cols"]
        if ws is None:
            return jsonify({"error": "Load your leads on the Auto Sender tab first."}), 400
        if cols.get("replied") is None:
            return jsonify({"error": "This sheet has no Replied? column."}), 400

        from datetime import date
        today = date.today()
        pending, sent_total, replied_total = [], 0, 0

        for row in ws.iter_rows(min_row=2):
            r   = row[0].row
            url = str(row[cols["url"]].value or "").strip()
            if not url.startswith("http"):
                continue
            contacted = (str(row[cols["contacted"]].value or "").strip().lower()
                         if cols["contacted"] is not None else "")
            if contacted not in ("yes", "y", "sent", "1", "true"):
                continue
            sent_total += 1

            replied = str(row[cols["replied"]].value or "").strip().lower()
            if replied:
                # Any non-empty value counts as "answered" for the stats, but
                # only an affirmative counts as an actual reply — so a future
                # "No" can be logged explicitly without inflating the rate.
                if replied in ("yes", "y", "1", "true", "igen"):
                    replied_total += 1
                continue

            last = _parse_date(row[cols["sent_date"]].value) if cols.get("sent_date") is not None else None
            pending.append({
                "row":   r,
                "name":  str(row[cols["name"]].value or "").strip(),
                "url":   url,
                "niche": str(row[cols["niche"]].value or "").strip() if cols.get("niche") is not None else "",
                "days":  (today - last).days if last else None,
                "sent":  last.isoformat() if last else "",
            })

        # Newest first: the freshest sends are the ones the user can actually
        # remember, so they're the cheapest to answer. Undated rows last.
        pending.sort(key=lambda p: (p["days"] is None, p["days"] if p["days"] is not None else 0))

        answered = sent_total - len(pending)
        return jsonify({
            "ok": True,
            "pending": pending,
            "count":   len(pending),
            "sent":    sent_total,
            "replied": replied_total,
            "logged":  answered,
            "rate":    round(replied_total / answered * 100, 1) if answered else None,
        })
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/mark_reply", methods=["POST"])
def mark_reply():
    """Log a yes/no answer for one lead from the Replies sweep.

    Separate from /mark_replied (which only ever writes "Yes", for the
    follow-up list's "Got a reply" button). Writing an explicit "No" matters:
    it's what distinguishes "this lead never answered" from "I haven't got
    round to checking this one", and without that distinction the pending list
    never empties and the reply rate has no denominator.
    """
    try:
        ws   = _auto_state["ws"]
        cols = _auto_state["cols"]
        if ws is None:
            return jsonify({"error": "Load your leads first."}), 400
        data = request.get_json(force=True)
        r = int(data.get("row", 0))
        if r < 2 or cols.get("replied") is None:
            return jsonify({"error": "Can't mark this row."}), 400
        ws.cell(row=r, column=cols["replied"] + 1,
                value="Yes" if data.get("replied") else "No")
        _save_wb()
        return jsonify({"ok": True})
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/mark_replied", methods=["POST"])
def mark_replied():
    """Mark a lead as replied so it drops off the follow-up list for good."""
    try:
        ws   = _auto_state["ws"]
        cols = _auto_state["cols"]
        if ws is None:
            return jsonify({"error": "Load your leads first."}), 400
        r = int(request.get_json(force=True).get("row", 0))
        if r < 2 or cols.get("replied") is None:
            return jsonify({"error": "Can't mark this row."}), 400
        ws.cell(row=r, column=cols["replied"] + 1, value="Yes")
        _save_wb()
        return jsonify({"ok": True})
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(debug=False, port=5001)

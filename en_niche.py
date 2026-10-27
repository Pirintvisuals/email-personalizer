"""
English niche detector — name-first classifier + weighted text scoring + opener validator.
English mirror of hu_niche.py (same machinery, UK trades, English copy).

Public API:
    detect_niche(text_data, company_name="") -> str
    detect_niche_from_name(name) -> str
    opener_matches_niche(opener, niche) -> bool
    NICHE_OPENER_INSTRUCTIONS: dict[str, str]
"""
import re
import unicodedata

# "" = truly unknown (no trade signal anywhere). Kept DISTINCT from the literal
# "general building" returned by the general-signal / multi-trade checks, so the
# caller can label a confirmed renovator as "renovation" while leaving a genuine
# unknown neutral ("businesses").
_DEFAULT_NICHE = ""


def _norm(text: str) -> str:
    """Lowercase + strip accent marks (NFKD decompose, drop combining chars)."""
    nfkd = unicodedata.normalize("NFKD", text.lower())
    return "".join(c for c in nfkd if not unicodedata.combining(c))


# ── Name-first rules ─────────────────────────────────────────────────────────────
# Each tuple: (niche_label, [regex patterns on normalised name])
# Checked top-to-bottom; first match wins.
# NOTE: garage doors / paving are placed before "windows and doors" and
# "landscaping" so "garage door" and "driveway" don't get mis-grabbed.

_NAME_RULES: list[tuple[str, list[str]]] = [
    # ── 1. Garage doors and shading ───────────────────────────────────────────────
    ("garage doors and shading", [
        r"\bgarage[\s\-]?door",
        r"\broller[\s\-]?(door|shutter)",
        r"\bsectional[\s\-]?door",
        r"\bawning", r"\bcanop", r"\bpergola",
        r"\bshutter[\s\-]?(co|ltd|services|specialist)",
        r"\bsomfy\b", r"\bhormann\b",
    ]),
    # ── 2. Windows and doors ──────────────────────────────────────────────────────
    ("windows and doors", [
        r"\bwindow",                          # Window World, Bolton Windows…
        r"\bglazing\b", r"\bglaz",            # Glazing Co, double glazing
        r"\bupvc\b",
        r"\bconservator",
        r"\bbi[\s\-]?fold", r"\bbifold",
        r"\bfront[\s\-]?door", r"\bcomposite[\s\-]?door",
        r"\brehau\b", r"\bveka\b", r"\bliniar\b", r"\bschuco\b", r"\bschueco\b",
        r"\borigin[\s\-]?(door|window|frames)", r"\baluk\b",
    ]),
    # ── 3. Air conditioning ───────────────────────────────────────────────────────
    ("air conditioning", [
        r"\bair[\s\-]?con",                   # air con, aircon, air conditioning
        r"\bac[\s\-]?(cool|climate|install|services|solutions|systems|ltd|co)\b",
        r"\bcool[\s\-]?(air|ing|tech|breeze|comfort)",
        r"\b(comfort|fresh)[\s\-]?air\b",
        r"\bhvac\b",
        r"\bheat[\s\-]?pump",
        r"\bclimate[\s\-]?(control|solutions)",
        r"\brefrigerat",
        r"\bdaikin\b", r"\bmitsubishi\b", r"\bfujitsu\b",
    ]),
    # ── 4. Heating (boiler / gas) ─────────────────────────────────────────────────
    ("heating", [
        r"\bheating\b",
        r"\bboiler",
        r"\bgas[\s\-]?(safe|engineer|fitting|works|services)",
        r"\bheating[\s\-]?(engineer|services|solutions|specialist)",
        r"\bcentral[\s\-]?heating",
        r"\bvaillant\b", r"\bworcester\b", r"\bbaxi\b", r"\bideal[\s\-]?(boiler|heating)",
        r"\bglow[\s\-]?worm\b", r"\bviessmann\b",
    ]),
    # ── 5. Solar ──────────────────────────────────────────────────────────────────
    ("solar", [
        r"\bsolar",                           # solar, solartech, solarpro…
        r"\bphotovolt",
        r"\bpv[\s\-]?(system|install|solutions|ltd|co)",
        r"\bpv\b",
        r"\brenewable",
        r"\bsolaredge\b", r"\bgivenergy\b",
    ]),
    # ── 6. Roofing ────────────────────────────────────────────────────────────────
    ("roofing", [
        r"\broofing\b",
        r"\broofer",
        r"\broofs?\b",
        r"\bflat[\s\-]?roof",
        r"\bpitched[\s\-]?roof",
        r"\bleadwork\b", r"\blead[\s\-]?work",
        r"\bguttering\b", r"\bfascia",
        r"\bre[\s\-]?roof",
        r"\bmarley\b", r"\bredland\b", r"\bsandtoft\b",
    ]),
    # ── 7. Electrical ─────────────────────────────────────────────────────────────
    ("electrical", [
        r"\belectric",
        r"\belectrician",
        r"\bsparky\b", r"\bsparks\b",
        r"\bniceic\b", r"\bnapit\b",
        r"\bpat[\s\-]?test",
        r"\belectro[\s\-]?(tech|services|ltd)",
    ]),
    # ── Bathroom renovation — placed BEFORE tiling & plumbing so names like
    #    "X Plumbing & Bathrooms" or "Tile & Bathroom" resolve to bathroom,
    #    which is the core target niche. ─────────────────────────────────────────
    ("bathroom renovation", [
        r"\bbathroom",
        r"\bwet[\s\-]?room",
        r"\bensuite\b", r"\ben[\s\-]?suite\b",
        r"\bshower[\s\-]?room",
    ]),
    # ── 8. Tiling ─────────────────────────────────────────────────────────────────
    ("tiling", [
        r"\btiling\b",
        r"\btiler",
        r"\bwall[\s\-]?and[\s\-]?floor[\s\-]?til",
        r"\bceramic[\s\-]?til",
    ]),
    # ── 9. Painting and decorating ────────────────────────────────────────────────
    ("painting and decorating", [
        r"\bpaint",
        r"\bdecorat",
        r"\bpainter",
    ]),
    # ── 10. Bricklaying ───────────────────────────────────────────────────────────
    ("bricklaying", [
        r"\bbricklay",
        r"\bbrickwork\b",
        r"\bbrick[\s\-]?and[\s\-]?block",
    ]),
    # ── 11. Plastering ────────────────────────────────────────────────────────────
    ("plastering", [
        r"\bplaster",
        r"\bdrylin",
        r"\bskim",
        r"\bvenetian[\s\-]?plaster",
    ]),
    # ── 12. Joinery ───────────────────────────────────────────────────────────────
    ("joinery", [
        r"\bjoiner",
        r"\bjoinery\b",
        r"\bcarpent",
        r"\bbespoke[\s\-]?(furniture|kitchen|joinery)",
        r"\bkitchen[\s\-]?fitt",
        r"\bfitted[\s\-]?(furniture|wardrobe)",
    ]),
    # ── 13. Plumbing ──────────────────────────────────────────────────────────────
    ("plumbing", [
        r"\bplumb",
        r"\bplumber",
    ]),
    # ── 14. Paving ────────────────────────────────────────────────────────────────
    ("paving", [
        r"\bpaving\b",
        r"\bblock[\s\-]?pav",
        r"\bdriveway",
        r"\bresin[\s\-]?(driveway|bound|drive)",
        r"\btarmac",
        r"\bmarshalls\b", r"\bbradstone\b", r"\bbrett[\s\-]?(landscap|paving)",
    ]),
    # ── 15. External wall insulation ──────────────────────────────────────────────
    ("external wall insulation", [
        r"\bexternal[\s\-]?wall[\s\-]?insulation",
        r"\bewi\b",
        r"\brendering\b", r"\brenderer",
        r"\bk[\s\-]?rend\b",
        r"\bsilicone[\s\-]?render",
    ]),
    # ── 16. Landscaping ───────────────────────────────────────────────────────────
    ("landscaping", [
        r"\blandscap",
        r"\bgarden",
        r"\bground[\s\-]?work",
        r"\bturf",
        r"\bfencing\b", r"\bfence",
        r"\bartificial[\s\-]?grass", r"\bartificial[\s\-]?turf",
        r"\bhedg", r"\btree[\s\-]?(surgeon|work|services)",
    ]),
    # ── 17. Decking ───────────────────────────────────────────────────────────────
    ("decking", [
        r"\bdecking\b",
        r"\bcomposite[\s\-]?deck",
        r"\btrex\b", r"\bmillboard\b", r"\bcladco\b",
    ]),
    # ── 18. Pool building ─────────────────────────────────────────────────────────
    ("pool building", [
        r"\bswimming[\s\-]?pool",
        r"\bpool[\s\-]?(build|install|construct|company|specialist)",
        r"\bhot[\s\-]?tub",
    ]),
    # ── 19. Flooring ──────────────────────────────────────────────────────────────
    ("flooring", [
        r"\bflooring\b",
        r"\bfloor[\s\-]?layer",
        r"\blaminate[\s\-]?floor",
        r"\bparquet\b",
        r"\bkarndean\b", r"\bamtico\b", r"\bquick[\s\-]?step\b",
    ]),
]


def detect_niche_from_name(name: str) -> str:
    """
    Fast check: normalise company name, test each niche's patterns.
    Returns niche label or "" if nothing matches.
    """
    if not name or not name.strip():
        return ""
    n = _norm(name)
    for niche, patterns in _NAME_RULES:
        for pat in patterns:
            if re.search(pat, n):
                return niche
    return ""


# ── Weighted text scoring ────────────────────────────────────────────────────────
# weight: 4 = uniquely niche-specific, 3 = very specific, 2 = specific, 1 = weak

_TEXT_KEYWORDS: dict[str, list[tuple[str, int]]] = {
    "windows and doors": [
        ("windows", 3), ("window", 3), ("double glazing", 4), ("glazing", 3),
        ("upvc", 4), ("conservatory", 4), ("bi-fold", 4), ("bifold", 4),
        ("composite door", 4), ("front door", 3),
        ("rehau", 4), ("veka", 4), ("liniar", 4), ("schuco", 4),
        ("origin doors", 4), ("residence 9", 4),
    ],
    "air conditioning": [
        ("air conditioning", 4), ("air con", 4), ("aircon", 4),
        ("heat pump", 4), ("hvac", 4), ("air source", 4),
        ("split system", 3), ("daikin", 4), ("mitsubishi electric", 4),
        ("fujitsu", 4), ("panasonic", 3), ("samsung", 2),
        ("refrigeration", 3), ("cooling", 2),
    ],
    "heating": [
        ("heating engineer", 4), ("central heating", 4), ("boiler", 4),
        ("gas safe", 4), ("combi boiler", 4), ("boiler install", 4),
        ("underfloor heating", 4), ("radiator", 3), ("gas engineer", 4),
        ("worcester bosch", 4), ("vaillant", 3), ("baxi", 3), ("ideal boiler", 4),
        ("glow worm", 3), ("viessmann", 3),
        ("gas boiler", 4), ("boiler service", 4), ("boiler servicing", 4),
        ("boiler repair", 4), ("boiler replacement", 4), ("boiler installation", 4),
        ("gas safe registered", 4), ("gas fitter", 4), ("power flush", 4),
        ("powerflush", 4), ("landlord gas", 4), ("gas certificate", 4),
        ("heating system", 3), ("heating and plumbing", 3), ("unvented", 3),
    ],
    "solar": [
        ("solar panel", 4), ("solar pv", 4), ("photovoltaic", 4),
        ("solar energy", 4), ("renewable energy", 4),
        ("longi", 4), ("jinko", 4), ("canadian solar", 4),
        ("solaredge", 4), ("givenergy", 4), ("tesla powerwall", 4),
        ("battery storage", 3),
    ],
    "roofing": [
        ("roofing", 4), ("roofer", 4), ("flat roof", 4),
        ("pitched roof", 4), ("re-roof", 4), ("leadwork", 4), ("lead flashing", 4),
        ("guttering", 3), ("fascia", 3), ("soffit", 3), ("epdm", 4),
        ("marley", 4), ("redland", 4), ("velux", 3), ("fibreglass roof", 4),
    ],
    "electrical": [
        ("electrician", 4), ("electrical", 4), ("rewire", 4),
        ("niceic", 4), ("napit", 4), ("consumer unit", 4),
        ("fuse board", 3), ("pat testing", 4), ("ev charger", 3),
    ],
    "tiling": [
        ("tiling", 4), ("tiler", 4),
        ("wall and floor tiling", 4), ("ceramic tiling", 4),
        ("porcelain tiles", 4), ("mosaic", 3),
        ("wall tiling", 4), ("floor tiling", 4), ("tile fitter", 4),
        ("tiling services", 4), ("tiling specialist", 4),
    ],
    "painting and decorating": [
        ("painting and decorating", 4), ("painter and decorator", 4),
        ("decorator", 4), ("interior painting", 4), ("exterior painting", 4),
        ("spray painting", 3),
    ],
    "bricklaying": [
        ("bricklaying", 4), ("bricklayer", 4), ("brickwork", 4),
        ("brick and block", 4), ("blockwork", 3),
    ],
    "plastering": [
        ("plastering", 4), ("plasterer", 4), ("skimming", 4),
        ("dry lining", 4), ("drylining", 4), ("venetian plaster", 4),
    ],
    "joinery": [
        ("joinery", 4), ("joiner", 4), ("carpenter", 4), ("carpentry", 4),
        ("bespoke furniture", 4), ("fitted wardrobe", 4), ("kitchen fitting", 3),
        # Kitchen / bespoke specialists describe themselves in many ways in the bio:
        ("bespoke kitchen", 4), ("fitted kitchen", 4), ("handmade kitchen", 4),
        ("shaker kitchen", 4), ("kitchen design", 4), ("kitchen installation", 4),
        ("kitchen fitter", 4), ("bespoke joinery", 4), ("bespoke cabinetry", 4),
        ("fitted bedroom", 4), ("bespoke bedroom", 4), ("made to measure", 3),
        ("kitchens", 2), ("kitchen", 2), ("cabinet maker", 4), ("wardrobes", 2),
    ],
    "plumbing": [
        ("plumbing", 4), ("plumber", 4), ("plumbing and heating", 3),
        ("leak repair", 3), ("bathroom plumbing", 3),
        ("plumbing services", 4), ("emergency plumber", 4), ("blocked drain", 3),
        ("drainage", 3), ("burst pipe", 3), ("tap repair", 2), ("toilet repair", 2),
    ],
    "paving": [
        ("block paving", 4), ("paving", 4), ("driveway", 4),
        ("resin driveway", 4), ("resin bound", 4), ("tarmac", 3),
        ("marshalls", 4), ("bradstone", 4), ("indian sandstone driveway", 4),
    ],
    "external wall insulation": [
        ("external wall insulation", 4), ("ewi", 4),
        ("rendering", 4), ("k rend", 4), ("silicone render", 4),
        ("through colour render", 4),
    ],
    "landscaping": [
        ("landscaping", 4), ("landscaper", 4), ("garden design", 4),
        ("groundworks", 3), ("turfing", 4), ("artificial grass", 4),
        ("fencing", 3), ("patio", 3), ("sleepers", 3), ("tree surgery", 3),
    ],
    "decking": [
        ("composite decking", 4), ("decking", 4), ("trex", 4),
        ("millboard", 4), ("cladco", 4), ("hardwood decking", 4),
    ],
    "pool building": [
        ("swimming pool", 4), ("pool construction", 4), ("pool install", 4),
        ("hot tub", 3), ("pool refurbishment", 4),
    ],
    "flooring": [
        ("flooring", 4), ("floor layer", 4), ("laminate flooring", 4),
        ("luxury vinyl", 4), ("karndean", 4), ("amtico", 4), ("parquet", 3),
    ],
    "bathroom renovation": [
        ("bathroom renovation", 4), ("bathroom fitting", 4),
        ("bathroom installation", 4), ("wet room", 4), ("ensuite", 3),
        ("bathroom fitter", 4), ("bathroom refurbishment", 4), ("bathroom design", 4),
        ("bathroom specialist", 4), ("shower room", 4), ("mobility bathroom", 4),
        ("disabled bathroom", 4), ("walk in shower", 3), ("bathroom suite", 3),
        ("bathrooms", 3), ("bathroom", 2),
    ],
}

_GENERAL_SIGNALS = {
    "general builder", "general building", "building contractor",
    "general contractor", "full house renovation", "complete renovation",
    "complete refurbishment", "full refurbishment", "property renovation",
    "all trades", "all aspects of building", "construction company",
    "design and build",
    # ── Generic renovators / handymen ────────────────────────────────────────
    # These do a bit of everything, so a single trade keyword in a post must
    # NOT lock them to that trade. Catch them here → "general building" →
    # defaults to the safe "renovation" label in dm_app.pick_body().
    "house fixer", "handyman", "handyperson", "home improvement",
    "home improvements", "home renovation", "house renovation",
    "house renovations", "home refurbishment", "house refurbishment",
    "renovation company", "renovation specialist", "renovations and",
    "property maintenance", "property refurbishment", "home repairs",
    # ── Hungarian renovation terms (accent-stripped by _norm) ────────────────
    "lakasfelujitas", "felujitas", "ingatlanfelujitas", "teljes felujitas",
    # ── Extensions / conversions / fit-outs = general build work ─────────────
    "house extension", "home extension", "extensions and renovations",
    "loft conversion", "loft conversions", "garage conversion",
    "barn conversion", "new build", "extension specialist",
}

_MIN_SCORE  = 3   # must score at least this to claim a niche
_MIN_MARGIN = 2   # must beat second place by this much


def _score_text(text_data: str) -> str:
    t_orig = text_data.lower()
    t_norm = _norm(text_data)

    # Immediate general contractor detection
    for sig in _GENERAL_SIGNALS:
        if sig in t_orig or sig in t_norm:
            return "general building"

    scores: dict[str, int] = {}
    for niche, kw_list in _TEXT_KEYWORDS.items():
        s = sum(w for kw, w in kw_list if _norm(kw) in t_norm)
        if s > 0:
            scores[niche] = s

    if not scores:
        return _DEFAULT_NICHE

    # 4+ different trades scored → general contractor
    if len(scores) >= 4:
        return "general building"

    top = sorted(scores.items(), key=lambda x: -x[1])
    best_niche, best_score = top[0]

    # Minimum threshold
    if best_score < _MIN_SCORE:
        return _DEFAULT_NICHE

    # Margin guard: if two niches are too close, fall back to default
    if len(top) >= 2 and (best_score - top[1][1]) < _MIN_MARGIN:
        return _DEFAULT_NICHE

    return best_niche


def detect_niche(text_data: str, company_name: str = "") -> str:
    """
    Primary entry point. Name-first → text scoring → default.
    """
    if company_name:
        name_niche = detect_niche_from_name(company_name)
        if name_niche:
            return name_niche
    if text_data and text_data.strip():
        return _score_text(text_data)
    return _DEFAULT_NICHE


# ── Per-niche opener instructions ────────────────────────────────────────────────
# Injected into the prompt before the vision call so the model knows what to look for.

NICHE_OPENER_INSTRUCTIONS: dict[str, str] = {
    "windows and doors": (
        "IMPORTANT CONTEXT: This is a window / door fitting business. "
        "If the photo shows decking, paving, a garden, an AC outdoor unit, solar or a roof, "
        "IGNORE IT COMPLETELY — those are other businesses' photos showing up at random. "
        "On a window/door photo look for: "
        "frame colour (anthracite, white, golden oak, dark brown) and how it suits the facade, "
        "opening type (tilt and turn, sliding / lift-and-slide, bi-fold), "
        "reveal depth (deeply set frame vs flush), "
        "cill slope (fall for drainage), "
        "frame-to-wall seal (taped vs just rendered), "
        "on a front door the threshold meeting the floor, "
        "on a roof window the flashing return, "
        "integrated roller shutter box built into the wall. "
        "If no window/door work is visible → FALLBACK."
    ),
    "air conditioning": (
        "IMPORTANT CONTEXT: This is an air conditioning business. "
        "If the photo shows decking, paving, a garden, a window, a roof or a boiler, IGNORE IT. "
        "On an AC photo look for: Daikin / Mitsubishi Electric / Fujitsu / Panasonic / LG / Samsung brand, "
        "outdoor unit placement, neatness of the pipework run, "
        "indoor unit position and type. "
        "If there's no AC / split / heat pump work in the photo → FALLBACK."
    ),
    "heating": (
        "IMPORTANT CONTEXT: This is a heating and gas business. "
        "If the photo shows decking, paving, a garden, a window, a roof or an AC outdoor unit, IGNORE IT. "
        "On a heating photo look for: "
        "Worcester Bosch / Vaillant / Ideal / Baxi / Viessmann boiler, "
        "pipework run on the wall, underfloor heating manifold, radiator connections, "
        "gas meter box, plant room. "
        "If there's no boiler / heating system work in the photo → FALLBACK."
    ),
    "solar": (
        "IMPORTANT CONTEXT: This is a solar business. "
        "If the photo shows a garden, decking, a window, AC or a boiler, IGNORE IT. "
        "On a solar photo look for: LONGi / JA Solar / Jinko / Canadian Solar / Trina panels, "
        "inverter (SolarEdge / SMA / Fronius / GivEnergy), "
        "panel symmetry and even pitch, cable management, mounting rails. "
        "If there's no solar system in the photo → FALLBACK."
    ),
    "roofing": (
        "IMPORTANT CONTEXT: This is a roofing business. "
        "If the photo shows AC, decking, a garden or a window, IGNORE IT. "
        "On a roofing photo look for: Marley / Redland / Sandtoft tiles, "
        "lead flashing detail, EPDM/GRP flat roof, Velux / Fakro roof window, "
        "ridge line, verge detail, gutter fall, chimney flashing. "
        "If there's no roof in the photo → FALLBACK."
    ),
    "electrical": (
        "IMPORTANT CONTEXT: This is an electrical business. "
        "On an electrical photo look for: consumer unit / fuse board layout, "
        "breaker arrangement, cable and conduit runs, EV charger install, "
        "downlight layout. If there's no electrical work in the photo → FALLBACK."
    ),
    "tiling": (
        "IMPORTANT CONTEXT: This is a tiling business. "
        "On a tiling photo look for: porcelain / ceramic / natural stone, "
        "grout line consistency, floor-to-wall transition, "
        "layout pattern (brick bond, herringbone, diagonal), mosaic, trim edge. "
        "If there's no tiling in the photo → FALLBACK."
    ),
    "painting and decorating": (
        "IMPORTANT CONTEXT: This is a painting and decorating business. "
        "On a decorating photo look for: even coverage on a facade, colour, "
        "cutting-in accuracy around door frames, wall-to-ceiling line, "
        "finish quality. If there's no painting work in the photo → FALLBACK."
    ),
    "bricklaying": (
        "IMPORTANT CONTEXT: This is a bricklaying business. "
        "On a bricklaying photo look for: brick bond, pointing, "
        "straightness of the corners (quoins), bed joints, course consistency. "
        "If there's no brickwork in the photo → FALLBACK."
    ),
    "plastering": (
        "IMPORTANT CONTEXT: This is a plastering / dry lining business. "
        "On the photo look for: flatness of the skim, internal/external angles, "
        "ceiling finish, render coverage. If there's no plastering work → FALLBACK."
    ),
    "joinery": (
        "IMPORTANT CONTEXT: This is a joinery / carpentry business. "
        "On the photo look for: type of bespoke piece, timber finish, "
        "hinges and handles, accuracy of the joints. "
        "If there's no joinery work → FALLBACK."
    ),
    "plumbing": (
        "IMPORTANT CONTEXT: This is a plumbing business. "
        "On the photo look for: pipework layout, pump, taps and fittings, "
        "sanitary install. If there's no plumbing work → FALLBACK."
    ),
    "garage doors and shading": (
        "IMPORTANT CONTEXT: This is a garage door / shading business. "
        "If the photo shows a window, roof, garden or AC, IGNORE IT. "
        "On the photo look for: sectional / roller door type, Somfy / Hörmann motor, "
        "awning, canopy, pergola. If there's no garage door / shading → FALLBACK."
    ),
    "paving": (
        "IMPORTANT CONTEXT: This is a paving business. "
        "On a paving photo look for: Marshalls / Bradstone / Brett block type, "
        "joint consistency, edging course, laying pattern, resin finish. "
        "If there's no paving / driveway in the photo → FALLBACK."
    ),
    "external wall insulation": (
        "IMPORTANT CONTEXT: This is an external wall insulation / rendering business. "
        "On the photo look for: K Rend / silicone render finish, corner beads, "
        "smoothing around window reveals, mesh/fixing system. "
        "If there's no render / EWI work → FALLBACK."
    ),
    "landscaping": (
        "IMPORTANT CONTEXT: This is a landscaping business. "
        "On the photo look for: paving, sleepers, retaining levels, turf / artificial grass evenness, "
        "planting layout, fencing line, irrigation. "
        "If there's no landscaping / garden work in the photo → FALLBACK."
    ),
    "decking": (
        "IMPORTANT CONTEXT: This is a decking business. "
        "On the photo look for: composite / hardwood board type and colour, "
        "even board spacing, fascia/edge trim, substructure. "
        "If there's no decking in the photo → FALLBACK."
    ),
    "pool building": (
        "IMPORTANT CONTEXT: This is a swimming pool business. "
        "On the photo look for: pool finish type and colour, filtration, "
        "step detail, coping edge. If there's no pool in the photo → FALLBACK."
    ),
    "flooring": (
        "IMPORTANT CONTEXT: This is a flooring business. "
        "On the photo look for: laminate / LVT / parquet type, layout pattern, "
        "skirting/beading, floor-to-wall meeting. "
        "If there's no flooring in the photo → FALLBACK."
    ),
    "bathroom renovation": (
        "IMPORTANT CONTEXT: This is a bathroom renovation business. "
        "On the photo look for: tile / panel type, sanitaryware layout, "
        "taps and shower install, grout line. If there's no bathroom in the photo → FALLBACK."
    ),
    "general building": (
        "IMPORTANT CONTEXT: This is a general builder / contractor. "
        "They do a range of trades. React to whatever finished trade work you can see in the photo — "
        "be it a roof, paving, tiling, brickwork or anything else."
    ),
}


# ── Opener post-validation ───────────────────────────────────────────────────────
# If opener contains these niche-contradicting terms → reject and use fallback.

_NICHE_VOCAB: dict[str, dict[str, list[str]]] = {
    "bathroom renovation": {"contradict": [
        "decking", "paving", "driveway", "block pav", "patio", "tarmac",
        "fence", "fencing", "turf", "artificial grass", "garden",
        "media wall", "soundbar", "tv unit",
        "roof tile", "flat roof", "solar", "air con", "aircon",
        "window frame", "window", "glazing", "conservatory", "bi-fold", "upvc",
        # heating/boiler — a boiler opener on a bathroom lead is a clash
        # (NB: not "radiator" — bathroom fitters do towel rails)
        "boiler", "flue", "combi", "central heating",
        "worcester bosch", "vaillant", "baxi", "viessmann",
    ]},
    "windows and doors": {"contradict": [
        "decking", "paving", "driveway", "block pav",
        "air con", "aircon", "solar", "boiler",
        "roof tile", "garden", "turf",
    ]},
    "air conditioning": {"contradict": [
        "decking", "paving", "driveway", "window", "front door", "back door", "composite door", "door frame",
        "boiler", "radiator", "solar", "roof tile",
        "garden", "turf",
    ]},
    "heating": {"contradict": [
        "decking", "paving", "driveway", "window", "front door", "back door", "composite door", "door frame",
        "air con", "aircon", "solar", "roof tile",
        "garden", "turf",
    ]},
    "solar": {"contradict": [
        "decking", "paving", "driveway", "window", "front door", "back door", "composite door", "door frame",
        "air con", "aircon", "boiler", "roof tile of",
        "garden", "turf",
    ]},
    "roofing": {"contradict": [
        "decking", "paving", "driveway", "window", "front door", "back door", "composite door", "door frame",
        "air con", "aircon", "boiler", "solar panel",
        "garden", "turf",
    ]},
    "garage doors and shading": {"contradict": [
        "decking", "paving", "air con", "aircon",
        "solar", "window frame", "roof tile",
    ]},
    "paving": {"contradict": [
        "air con", "aircon", "boiler",
        "solar", "window", "roof tile",
        "decking board",
    ]},
    "landscaping": {"contradict": [
        "air con", "aircon", "boiler",
        "solar panel", "roof tile", "window frame",
    ]},
}


def opener_matches_niche(opener: str, niche: str) -> bool:
    """Return False if opener contradicts the niche (wrong-niche hallucination)."""
    if not opener or not niche:
        return True
    vocab = _NICHE_VOCAB.get(niche)
    if vocab is None:
        return True
    o = _norm(opener)
    for term in vocab.get("contradict", []):
        if _norm(term) in o:
            print(f"  Opener rejected: '{term}' in '{opener}' (niche={niche})")
            return False
    return True

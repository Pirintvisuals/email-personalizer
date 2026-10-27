"""
German niche detector — name-first classifier + weighted text scoring + opener validator.
German/Austrian mirror of en_niche.py (same machinery, DACH trades, German copy).

The niche LABELS (dict keys) stay in English on purpose — they are internal
identifiers shared with dm_app_de.py (_canon_niche, _pitch_for,
NICHE_OPENER_INSTRUCTIONS, _NICHE_VOCAB). Only the keywords and the prose are
German, so the detector works on German-language Facebook pages and Excel columns.

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
# unknown neutral ("Betriebe").
_DEFAULT_NICHE = ""


def _norm(text: str) -> str:
    """Lowercase + strip accent marks (NFKD decompose, drop combining chars).
    Strips German umlauts too (ä->a, ö->o, ü->u) so 'Türen' matches 'turen'."""
    nfkd = unicodedata.normalize("NFKD", text.lower())
    return "".join(c for c in nfkd if not unicodedata.combining(c))


# ── Name-first rules ─────────────────────────────────────────────────────────────
# Each tuple: (niche_label, [regex patterns on normalised name])
# Checked top-to-bottom; first match wins.
# NOTE: garage doors / paving are placed before "windows and doors" and
# "landscaping" so "Garagentor" and "Einfahrt" don't get mis-grabbed.
# Patterns run on _norm() output, so umlauts are already stripped (türen->turen).

_NAME_RULES: list[tuple[str, list[str]]] = [
    # ── 1. Garage doors and shading ───────────────────────────────────────────────
    ("garage doors and shading", [
        r"\bgaragentor", r"\bgaragen[\s\-]?tor",
        r"\brolltor", r"\bsektionaltor",
        r"\bmarkise", r"\bpergola", r"\bbeschattung", r"\bsonnenschutz",
        r"\brolllad", r"\brolladen",
        r"\bsomfy\b", r"\bhormann\b",
    ]),
    # ── 2. Windows and doors ──────────────────────────────────────────────────────
    ("windows and doors", [
        r"\bfenster",                          # Fensterbau, Fenster Müller…
        r"\bverglas",                          # Verglasung
        r"\bkunststofffenster", r"\bkunststoff[\s\-]?fenster",
        r"\bwintergart",                       # Wintergarten
        r"\bhaustur", r"\bhaus[\s\-]?tur",
        r"\bschiebetur",
        r"\brehau\b", r"\bveka\b", r"\binternorm\b", r"\bschuco\b", r"\baluk\b",
        r"\btur(en)?\b",
    ]),
    # ── 3. Air conditioning ───────────────────────────────────────────────────────
    ("air conditioning", [
        r"\bklimaanlage", r"\bklima[\s\-]?(technik|service|technologie|systeme|anlagen)",
        r"\bklima\b",
        r"\bwarmepump",                        # Wärmepumpe
        r"\bkalte[\s\-]?(technik|anlagen|service)",
        r"\bkaltetechnik",
        r"\bsplit[\s\-]?(klima|gerat|anlage)",
        r"\bdaikin\b", r"\bmitsubishi\b", r"\bfujitsu\b",
    ]),
    # ── 4. Heating (boiler / gas) ─────────────────────────────────────────────────
    ("heating", [
        r"\bheizung",
        r"\btherme", r"\bgastherme",            # AT/DE for boiler
        r"\bgas[\s\-]?(technik|installateur|service)",
        r"\bheizungs?[\s\-]?(bau|technik|service|installateur)",
        r"\bzentralheiz",
        r"\bfussbodenheiz", r"\bfussboden[\s\-]?heiz",
        r"\bvaillant\b", r"\bviessmann\b", r"\bbuderus\b", r"\bjunkers\b", r"\bwolf[\s\-]?heiz",
    ]),
    # ── 5. Solar ──────────────────────────────────────────────────────────────────
    ("solar", [
        r"\bsolar",                            # Solar, Solartechnik…
        r"\bphotovolt",
        r"\bpv[\s\-]?(anlage|technik|systeme)",
        r"\bpv\b",
        r"\berneuerbare",
        r"\bsolaredge\b", r"\bfronius\b",
    ]),
    # ── 6. Roofing ────────────────────────────────────────────────────────────────
    ("roofing", [
        r"\bdachdeck",
        r"\bdach\b", r"\bdach[\s\-]?(bau|technik|service|deckerei)",
        r"\bflachdach",
        r"\bsteildach",
        r"\bspengler",                          # AT sheet-metal roofer
        r"\bspenglerei",
        r"\bdachrinne",
        r"\bzimmerei",                          # often does roofs (timber)
        r"\bmarley\b", r"\bbramac\b", r"\btondach\b",
    ]),
    # ── 7. Electrical ─────────────────────────────────────────────────────────────
    ("electrical", [
        r"\belektro",
        r"\belektriker",
        r"\belektrotechnik",
        r"\belektroinstall",
        r"\belektro[\s\-]?(technik|service|installation)",
    ]),
    # ── Bathroom renovation — placed BEFORE tiling & plumbing so names like
    #    "X Installateur & Bäder" or "Fliesen & Bad" resolve to bathroom,
    #    which is the core target niche. ─────────────────────────────────────────
    ("bathroom renovation", [
        r"\bbadezimmer",
        r"\bbadsanier",
        r"\bbadgestalt",
        r"\bbad[\s\-]?(bau|design|studio|profi|sanierung)",
        r"\bnasszelle", r"\bfeuchtraum",
        r"\bbader\b",                           # Bäder
    ]),
    # ── 8. Tiling ─────────────────────────────────────────────────────────────────
    ("tiling", [
        r"\bfliesenleg",
        r"\bfliesen",
        r"\bfliesenfach",
        r"\bnaturstein[\s\-]?(verleg|fliesen)",
    ]),
    # ── 9. Painting and decorating ────────────────────────────────────────────────
    ("painting and decorating", [
        r"\bmaler",
        r"\bmalerei",
        r"\banstrich",
        r"\blackier",
    ]),
    # ── 10. Bricklaying ───────────────────────────────────────────────────────────
    ("bricklaying", [
        r"\bmaurer",
        r"\bmaurermeist",
        r"\bmauerwerk",
    ]),
    # ── 11. Plastering ────────────────────────────────────────────────────────────
    ("plastering", [
        r"\bverputz",
        r"\bputz\b",
        r"\bstuckateur",
        r"\btrockenbau",
        r"\bspachtel",
    ]),
    # ── 12. Joinery ───────────────────────────────────────────────────────────────
    ("joinery", [
        r"\btischler",                          # AT joiner
        r"\bschreiner",                         # DE joiner
        r"\bzimmermann",
        r"\bmobel",                             # Möbel
        r"\bkuchenbau", r"\bkuchen[\s\-]?(studio|profi|werkstatt)",
        r"\beinbaukuch",
    ]),
    # ── 13. Plumbing ──────────────────────────────────────────────────────────────
    ("plumbing", [
        r"\binstallateur",                      # AT plumber
        r"\bsanitar",
        r"\bklempner",                          # DE plumber
        r"\bhaustechnik",
    ]),
    # ── 14. Paving ────────────────────────────────────────────────────────────────
    ("paving", [
        r"\bpflaster",
        r"\bpflasterung",
        r"\beinfahrt",
        r"\bhof[\s\-]?(pflaster|gestaltung)",
        r"\basphalt",
        r"\bgalabau",                           # Garten- u. Landschaftsbau (also paving)
    ]),
    # ── 15. External wall insulation ──────────────────────────────────────────────
    ("external wall insulation", [
        r"\bvollwarmeschutz",                   # AT term for EWI
        r"\bwarmedammung", r"\bwarme[\s\-]?dammung",
        r"\bwdvs\b",
        r"\bfassaden[\s\-]?(bau|sanierung|technik)",
        r"\bvollwarme",
    ]),
    # ── 16. Landscaping ───────────────────────────────────────────────────────────
    ("landscaping", [
        r"\blandschaftsbau",
        r"\bgartengestalt",
        r"\bgartenbau",
        r"\bgarten[\s\-]?(service|profi|design)",
        r"\bgarten\b",
        r"\bzaun", r"\bzaunbau",
        r"\brasen", r"\bkunstrasen",
        r"\bhecke", r"\bbaumpfleg",
    ]),
    # ── 17. Decking ───────────────────────────────────────────────────────────────
    ("decking", [
        r"\bterrassendiel",
        r"\bterrassen[\s\-]?(bau|profi)",
        r"\bwpc\b",
        r"\bholzterrass",
    ]),
    # ── 18. Pool building ─────────────────────────────────────────────────────────
    ("pool building", [
        r"\bschwimmbad",
        r"\bschwimmbecken",
        r"\bpool[\s\-]?(bau|technik|profi|service)",
        r"\bpoolbau",
        r"\bwhirlpool",
    ]),
    # ── 19. Flooring ──────────────────────────────────────────────────────────────
    ("flooring", [
        r"\bbodenleg",
        r"\bparkett",
        r"\blaminat",
        r"\bvinylboden",
        r"\bboden[\s\-]?(verleg|studio|profi)",
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
# All keywords are matched accent-stripped via _norm(), so umlauts are fine.

_TEXT_KEYWORDS: dict[str, list[tuple[str, int]]] = {
    "windows and doors": [
        ("fenster", 3), ("fenstertausch", 4), ("fenstermontage", 4),
        ("verglasung", 3), ("isolierglas", 4), ("kunststofffenster", 4),
        ("wintergarten", 4), ("haustür", 3), ("schiebetür", 4),
        ("rehau", 4), ("veka", 4), ("internorm", 4), ("schüco", 4),
        ("aluminiumfenster", 4), ("holzfenster", 3),
    ],
    "air conditioning": [
        ("klimaanlage", 4), ("klimatechnik", 4), ("klima", 3),
        ("wärmepumpe", 4), ("kältetechnik", 4), ("luftwärmepumpe", 4),
        ("splitgerät", 4), ("daikin", 4), ("mitsubishi electric", 4),
        ("fujitsu", 4), ("panasonic", 3), ("samsung", 2),
        ("klimagerät", 4), ("kühlung", 2),
    ],
    "heating": [
        ("heizungsbau", 4), ("zentralheizung", 4), ("therme", 4),
        ("gastherme", 4), ("heizungsanlage", 4), ("heizungstausch", 4),
        ("fußbodenheizung", 4), ("heizkörper", 3), ("gasinstallateur", 4),
        ("vaillant", 3), ("viessmann", 3), ("buderus", 3), ("junkers", 4),
        ("wolf heizung", 3), ("pelletsheizung", 4),
        ("gasheizung", 4), ("thermenwartung", 4), ("thermenservice", 4),
        ("heizungsreparatur", 4), ("kesseltausch", 4), ("heizungsinstallation", 4),
        ("gas wasser heizung", 4), ("ölheizung", 4), ("heizungssystem", 3),
        ("heizung und sanitär", 3), ("nahwärme", 3),
    ],
    "solar": [
        ("solaranlage", 4), ("solar pv", 4), ("photovoltaik", 4),
        ("solarenergie", 4), ("erneuerbare energie", 4),
        ("pv anlage", 4), ("stromspeicher", 3), ("fronius", 4),
        ("solaredge", 4), ("batteriespeicher", 4),
        ("solarstrom", 4),
    ],
    "roofing": [
        ("dachdecker", 4), ("dachdeckerei", 4), ("flachdach", 4),
        ("steildach", 4), ("dachsanierung", 4), ("spengler", 4), ("spenglerei", 4),
        ("dachrinne", 3), ("dachstuhl", 3), ("zimmerei", 3), ("dachfenster", 4),
        ("bramac", 4), ("tondach", 4), ("dacheindeckung", 4),
    ],
    "electrical": [
        ("elektriker", 4), ("elektrotechnik", 4), ("neuverkabelung", 4),
        ("elektroinstallation", 4), ("verteilerkasten", 4), ("sicherungskasten", 3),
        ("e-check", 3), ("ladestation", 3), ("wallbox", 3),
    ],
    "tiling": [
        ("fliesenleger", 4), ("fliesen", 4),
        ("wand und bodenfliesen", 4), ("natursteinfliesen", 4),
        ("feinsteinzeug", 4), ("mosaik", 3),
        ("wandfliesen", 4), ("bodenfliesen", 4), ("fliesenverlegung", 4),
        ("fliesenarbeiten", 4), ("fliesenfachbetrieb", 4),
    ],
    "painting and decorating": [
        ("maler und anstreicher", 4), ("malermeister", 4),
        ("malerei", 4), ("innenanstrich", 4), ("fassadenanstrich", 4),
        ("lackierung", 3),
    ],
    "bricklaying": [
        ("maurer", 4), ("maurermeister", 4), ("mauerwerk", 4),
        ("rohbau", 4), ("ziegelbau", 3),
    ],
    "plastering": [
        ("verputz", 4), ("verputzer", 4), ("stuckateur", 4),
        ("trockenbau", 4), ("trockenausbau", 4), ("spachtelarbeiten", 4),
    ],
    "joinery": [
        ("tischlerei", 4), ("tischler", 4), ("schreiner", 4), ("schreinerei", 4),
        ("möbeltischler", 4), ("einbauschrank", 4), ("küchenmontage", 3),
        # Kitchen / bespoke specialists describe themselves in many ways in the bio:
        ("einbauküche", 4), ("massivholzküche", 4), ("küche nach maß", 4),
        ("küchenstudio", 4), ("küchenplanung", 4), ("tischlerarbeiten", 4),
        ("möbel nach maß", 4), ("möbelbau", 4), ("innenausbau", 3),
        ("schlafzimmer nach maß", 4), ("nach maß", 3),
        ("küchen", 2), ("küche", 2), ("kleiderschrank", 2),
    ],
    "plumbing": [
        ("installateur", 4), ("sanitär", 4), ("gas wasser installateur", 3),
        ("rohrbruch", 3), ("badinstallation", 3),
        ("sanitärinstallation", 4), ("notdienst installateur", 4), ("rohrreinigung", 3),
        ("entwässerung", 3), ("wasserrohrbruch", 3), ("klempner", 4), ("haustechnik", 3),
    ],
    "paving": [
        ("pflasterung", 4), ("pflasterarbeiten", 4), ("einfahrt", 4),
        ("hofpflaster", 4), ("gehweg", 3), ("asphaltierung", 3),
        ("granitpflaster", 4), ("betonpflaster", 4), ("natursteinpflaster", 4),
    ],
    "external wall insulation": [
        ("vollwärmeschutz", 4), ("wärmedämmung", 4),
        ("wdvs", 4), ("fassadensanierung", 4), ("fassadendämmung", 4),
        ("thermische sanierung", 4),
    ],
    "landscaping": [
        ("landschaftsbau", 4), ("gartengestaltung", 4), ("gartenbau", 4),
        ("erdbau", 3), ("rasenverlegung", 4), ("kunstrasen", 4),
        ("zaunbau", 3), ("naturpool", 3), ("hochbeet", 3), ("baumpflege", 3),
    ],
    "decking": [
        ("terrassendielen", 4), ("holzterrasse", 4), ("wpc terrasse", 4),
        ("terrassenbau", 4), ("bambusterrasse", 4),
    ],
    "pool building": [
        ("schwimmbad", 4), ("poolbau", 4), ("schwimmbecken", 4),
        ("whirlpool", 3), ("poolsanierung", 4),
    ],
    "flooring": [
        ("bodenleger", 4), ("parkett", 4), ("laminat", 4),
        ("vinylboden", 4), ("designboden", 4), ("parkettboden", 3),
    ],
    "bathroom renovation": [
        ("badsanierung", 4), ("badezimmer", 4),
        ("badumbau", 4), ("nasszelle", 4), ("feuchtraum", 3),
        ("badgestaltung", 4), ("badsanierung aus einer hand", 4), ("baddesign", 4),
        ("badprofi", 4), ("duschbad", 4), ("seniorengerechtes bad", 4),
        ("barrierefreies bad", 4), ("begehbare dusche", 3), ("komplettbad", 3),
        ("bäder", 3), ("bad", 2),
    ],
}

_GENERAL_SIGNALS = {
    "baumeister", "bauunternehmen", "baufirma",
    "generalunternehmer", "komplettsanierung", "komplettumbau",
    "hausbau", "hausumbau", "schlüsselfertig", "schlusselfertig",
    "alle gewerke", "rund ums bau", "bau gmbh",
    "planung und ausführung", "planung und ausfuhrung",
    # ── Generic renovators / handymen ────────────────────────────────────────
    # These do a bit of everything, so a single trade keyword in a post must
    # NOT lock them to that trade. Catch them here → "general building" →
    # defaults to the safe "renovation" label in dm_app_de.pick_body().
    "hausmeister", "haus reparatur", "renovierung",
    "renovierungen", "wohnungssanierung", "haussanierung",
    "sanierungen", "sanierung und", "altbausanierung",
    "hausverwaltung", "objektsanierung", "instandsetzung",
    # ── Hungarian renovation terms (accent-stripped by _norm) ────────────────
    "lakasfelujitas", "felujitas", "ingatlanfelujitas", "teljes felujitas",
    # ── Extensions / conversions / fit-outs = general build work ─────────────
    "zubau", "anbau", "aufstockung", "umbau und sanierung",
    "dachausbau", "dachbodenausbau", "kellerausbau",
    "neubau", "zubau spezialist",
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
        "WICHTIGER KONTEXT: Das ist ein Fenster-/Türenbetrieb. "
        "Wenn das Foto eine Terrasse, Pflasterung, einen Garten, ein Klima-Außengerät, Solar oder ein Dach zeigt, "
        "IGNORIER ES KOMPLETT — das sind Fotos von anderen Betrieben, die zufällig auftauchen. "
        "Auf einem Fenster-/Türen-Foto achte auf: "
        "Rahmenfarbe (anthrazit, weiß, golden oak, dunkelbraun) und wie sie zur Fassade passt, "
        "Öffnungsart (Dreh-Kipp, Schiebe-/Hebe-Schiebe, Faltschiebe), "
        "Einbautiefe (tief eingesetzter Rahmen vs. bündig), "
        "Gefälle der Fensterbank (Wasserablauf), "
        "Anschluss Rahmen zur Wand (verklebt vs. nur verputzt), "
        "bei einer Haustür die Schwelle zum Bodenbelag, "
        "bei einem Dachfenster die Eindeckung rundherum, "
        "integrierter Rollladenkasten in der Wand. "
        "Wenn keine Fenster-/Türenarbeit zu sehen ist → FALLBACK."
    ),
    "air conditioning": (
        "WICHTIGER KONTEXT: Das ist ein Klimatechnik-Betrieb. "
        "Wenn das Foto eine Terrasse, Pflasterung, einen Garten, ein Fenster, ein Dach oder eine Therme zeigt, IGNORIER ES. "
        "Auf einem Klima-Foto achte auf: Daikin / Mitsubishi Electric / Fujitsu / Panasonic / LG / Samsung, "
        "Platzierung des Außengeräts, saubere Leitungsführung, "
        "Position und Typ des Innengeräts. "
        "Wenn keine Klima-/Split-/Wärmepumpen-Arbeit auf dem Foto ist → FALLBACK."
    ),
    "heating": (
        "WICHTIGER KONTEXT: Das ist ein Heizungs- und Gasbetrieb. "
        "Wenn das Foto eine Terrasse, Pflasterung, einen Garten, ein Fenster, ein Dach oder ein Klima-Außengerät zeigt, IGNORIER ES. "
        "Auf einem Heizungs-Foto achte auf: "
        "Vaillant / Viessmann / Buderus / Junkers Therme oder Kessel, "
        "Leitungsführung an der Wand, Fußbodenheizungs-Verteiler, Heizkörperanschlüsse, "
        "Gaszähler, Heizraum. "
        "Wenn keine Therme-/Heizungsarbeit auf dem Foto ist → FALLBACK."
    ),
    "solar": (
        "WICHTIGER KONTEXT: Das ist ein Solar-/Photovoltaik-Betrieb. "
        "Wenn das Foto einen Garten, eine Terrasse, ein Fenster, Klima oder eine Therme zeigt, IGNORIER ES. "
        "Auf einem Solar-Foto achte auf: Module (LONGi / JA Solar / Jinko / Canadian Solar / Trina), "
        "Wechselrichter (Fronius / SolarEdge / SMA), "
        "Symmetrie und gleichmäßige Neigung der Module, Kabelführung, Montageschienen. "
        "Wenn keine Solaranlage auf dem Foto ist → FALLBACK."
    ),
    "roofing": (
        "WICHTIGER KONTEXT: Das ist ein Dachdecker-/Spengler-Betrieb. "
        "Wenn das Foto Klima, eine Terrasse, einen Garten oder ein Fenster zeigt, IGNORIER ES. "
        "Auf einem Dach-Foto achte auf: Ziegel (Bramac / Tondach), "
        "Spenglerdetail / Verblechung, Flachdach (EPDM/Folie), Dachfenster, "
        "Firstlinie, Ortgangdetail, Dachrinnengefälle, Kaminanschluss. "
        "Wenn kein Dach auf dem Foto ist → FALLBACK."
    ),
    "electrical": (
        "WICHTIGER KONTEXT: Das ist ein Elektrobetrieb. "
        "Auf einem Elektro-Foto achte auf: Verteiler-/Sicherungskasten-Aufbau, "
        "Anordnung der Sicherungen, Kabel- und Leitungsführung, Wallbox-Montage, "
        "Beleuchtungs-/Spotanordnung. Wenn keine Elektroarbeit auf dem Foto ist → FALLBACK."
    ),
    "tiling": (
        "WICHTIGER KONTEXT: Das ist ein Fliesenleger-Betrieb. "
        "Auf einem Fliesen-Foto achte auf: Feinsteinzeug / Keramik / Naturstein, "
        "gleichmäßige Fugen, Übergang Boden zu Wand, "
        "Verlegemuster (Verband, Fischgrät, diagonal), Mosaik, Abschlussprofil. "
        "Wenn keine Fliesenarbeit auf dem Foto ist → FALLBACK."
    ),
    "painting and decorating": (
        "WICHTIGER KONTEXT: Das ist ein Maler-Betrieb. "
        "Auf einem Maler-Foto achte auf: gleichmäßige Deckung an einer Fassade, Farbe, "
        "saubere Kanten rund um Türstöcke, Linie Wand zu Decke, "
        "Oberflächenqualität. Wenn keine Malerarbeit auf dem Foto ist → FALLBACK."
    ),
    "bricklaying": (
        "WICHTIGER KONTEXT: Das ist ein Maurer-Betrieb. "
        "Auf einem Maurer-Foto achte auf: Mauerverband, Verfugung, "
        "Geradheit der Ecken, Lagerfugen, gleichmäßige Schichten. "
        "Wenn kein Mauerwerk auf dem Foto ist → FALLBACK."
    ),
    "plastering": (
        "WICHTIGER KONTEXT: Das ist ein Verputzer-/Trockenbau-Betrieb. "
        "Auf dem Foto achte auf: Ebenheit der Spachtelung, Innen-/Außenkanten, "
        "Deckenoberfläche, Putzdeckung. Wenn keine Verputzarbeit zu sehen ist → FALLBACK."
    ),
    "joinery": (
        "WICHTIGER KONTEXT: Das ist eine Tischlerei / Schreinerei. "
        "Auf dem Foto achte auf: Art des Möbelstücks, Holzoberfläche, "
        "Beschläge und Griffe, Genauigkeit der Verbindungen. "
        "Wenn keine Tischlerarbeit zu sehen ist → FALLBACK."
    ),
    "plumbing": (
        "WICHTIGER KONTEXT: Das ist ein Installateur-/Sanitär-Betrieb. "
        "Auf dem Foto achte auf: Leitungsführung, Pumpe, Armaturen und Anschlüsse, "
        "Sanitärinstallation. Wenn keine Installateurarbeit zu sehen ist → FALLBACK."
    ),
    "garage doors and shading": (
        "WICHTIGER KONTEXT: Das ist ein Garagentor-/Beschattungs-Betrieb. "
        "Wenn das Foto ein Fenster, Dach, einen Garten oder Klima zeigt, IGNORIER ES. "
        "Auf dem Foto achte auf: Sektional-/Rolltor-Typ, Somfy / Hörmann Antrieb, "
        "Markise, Vordach, Pergola. Wenn kein Garagentor / keine Beschattung → FALLBACK."
    ),
    "paving": (
        "WICHTIGER KONTEXT: Das ist ein Pflasterer-Betrieb. "
        "Auf einem Pflaster-Foto achte auf: Steintyp (Beton / Granit / Naturstein), "
        "gleichmäßige Fugen, Randstein, Verlegemuster, Asphaltoberfläche. "
        "Wenn keine Pflasterung / Einfahrt auf dem Foto ist → FALLBACK."
    ),
    "external wall insulation": (
        "WICHTIGER KONTEXT: Das ist ein Fassaden-/Vollwärmeschutz-Betrieb. "
        "Auf dem Foto achte auf: Putzoberfläche der Fassade, Kantenprofile, "
        "saubere Ausführung an Fensterlaibungen, Dämm-/Befestigungssystem. "
        "Wenn keine Fassaden-/Dämmarbeit zu sehen ist → FALLBACK."
    ),
    "landscaping": (
        "WICHTIGER KONTEXT: Das ist ein Garten-/Landschaftsbau-Betrieb. "
        "Auf dem Foto achte auf: Pflasterung, Stützmauern, Höhenführung, gleichmäßiger Rasen / Kunstrasen, "
        "Bepflanzung, Zaunlinie, Bewässerung. "
        "Wenn keine Garten-/Landschaftsarbeit auf dem Foto ist → FALLBACK."
    ),
    "decking": (
        "WICHTIGER KONTEXT: Das ist ein Terrassenbau-Betrieb. "
        "Auf dem Foto achte auf: Dielentyp und -farbe (WPC / Hartholz), "
        "gleichmäßiger Dielenabstand, Abschlussprofil, Unterkonstruktion. "
        "Wenn keine Terrasse auf dem Foto ist → FALLBACK."
    ),
    "pool building": (
        "WICHTIGER KONTEXT: Das ist ein Schwimmbad-/Poolbau-Betrieb. "
        "Auf dem Foto achte auf: Beckenoberfläche und -farbe, Filtertechnik, "
        "Stufendetail, Randstein. Wenn kein Pool auf dem Foto ist → FALLBACK."
    ),
    "flooring": (
        "WICHTIGER KONTEXT: Das ist ein Bodenleger-Betrieb. "
        "Auf dem Foto achte auf: Boden-Typ (Laminat / Vinyl / Parkett), Verlegemuster, "
        "Sockelleiste, Übergang Boden zu Wand. "
        "Wenn kein Boden auf dem Foto ist → FALLBACK."
    ),
    "bathroom renovation": (
        "WICHTIGER KONTEXT: Das ist ein Badsanierungs-Betrieb. "
        "Auf dem Foto achte auf: Fliesen-/Plattentyp, Anordnung der Sanitärobjekte, "
        "Armaturen und Dusche, Fugenbild. Wenn kein Bad auf dem Foto ist → FALLBACK."
    ),
    "general building": (
        "WICHTIGER KONTEXT: Das ist ein Baumeister / Generalunternehmer. "
        "Die machen verschiedene Gewerke. Reagier auf jede fertige Handwerksarbeit, die du auf dem Foto siehst — "
        "egal ob Dach, Pflasterung, Fliesen, Mauerwerk oder etwas anderes."
    ),
}


# ── Opener post-validation ───────────────────────────────────────────────────────
# If opener contains these niche-contradicting terms → reject and use fallback.

_NICHE_VOCAB: dict[str, dict[str, list[str]]] = {
    "bathroom renovation": {"contradict": [
        "terrasse", "pflaster", "einfahrt", "asphalt",
        "zaun", "kunstrasen", "garten",
        "media wall", "soundbar", "tv-wand",
        "dachziegel", "flachdach", "solar", "klimaanlage", "klima",
        "fensterrahmen", "fenster", "verglasung", "wintergarten",
        # heating/boiler — a boiler opener on a bathroom lead is a clash
        # (NB: not "Heizkörper" — bathroom fitters do towel rails)
        "therme", "gastherme", "kessel", "zentralheizung",
        "vaillant", "viessmann", "buderus",
    ]},
    "windows and doors": {"contradict": [
        "terrasse", "pflaster", "einfahrt",
        "klimaanlage", "klima", "solar", "therme",
        "dachziegel", "garten", "kunstrasen",
    ]},
    "air conditioning": {"contradict": [
        "terrasse", "pflaster", "einfahrt", "fenster", "tür",
        "therme", "heizkörper", "solar", "dachziegel",
        "garten", "kunstrasen",
    ]},
    "heating": {"contradict": [
        "terrasse", "pflaster", "einfahrt", "fenster", "tür",
        "klimaanlage", "klima", "solar", "dachziegel",
        "garten", "kunstrasen",
    ]},
    "solar": {"contradict": [
        "terrasse", "pflaster", "einfahrt", "fenster", "tür",
        "klimaanlage", "klima", "therme",
        "garten", "kunstrasen",
    ]},
    "roofing": {"contradict": [
        "terrasse", "pflaster", "einfahrt", "fenster", "tür",
        "klimaanlage", "klima", "therme", "solarmodul",
        "garten", "kunstrasen",
    ]},
    "garage doors and shading": {"contradict": [
        "terrasse", "pflaster", "klimaanlage", "klima",
        "solar", "fensterrahmen", "dachziegel",
    ]},
    "paving": {"contradict": [
        "klimaanlage", "klima", "therme",
        "solar", "fenster", "dachziegel",
        "terrassendiele",
    ]},
    "landscaping": {"contradict": [
        "klimaanlage", "klima", "therme",
        "solarmodul", "dachziegel", "fensterrahmen",
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

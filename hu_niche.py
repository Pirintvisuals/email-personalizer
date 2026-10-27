"""
Hungarian niche detector — name-first classifier + weighted text scoring + opener validator.

Public API:
    _hu_detect_niche(text_data, company_name="") -> str
    _hu_detect_niche_from_name(name) -> str
    _opener_matches_niche(opener, niche) -> bool
    NICHE_OPENER_INSTRUCTIONS: dict[str, str]
"""
import re
import unicodedata

_DEFAULT_NICHE = "felújító kivitelező"


def _norm(text: str) -> str:
    """Lowercase + strip accent marks (NFKD decompose, drop combining chars)."""
    nfkd = unicodedata.normalize("NFKD", text.lower())
    return "".join(c for c in nfkd if not unicodedata.combining(c))


# ── Name-first rules ─────────────────────────────────────────────────────────────
# Each tuple: (niche_label, [regex patterns on normalised name])
# Checked top-to-bottom; first match wins.

_NAME_RULES: list[tuple[str, list[str]]] = [
    # ── 1. Nyílászáró (ablak / window) ───────────────────────────────────────────
    ("nyílászáró", [
        r"\bablak",                           # Ablak Arzenál, Ablak-Bau, Benyó Ablak…
        r"\bnyilaszar",                       # Nyílászáró Centrum, Nyilaszaro Bt
        r"ablak[\s\-]?(szere|epites|bepites|csere|mester|gyart|kft|bt)",
        r"nyilaszaro[\s\-]?(szere|epites|bepites|csere|keszites|kft|bt)",
        r"\bschuco\b", r"\brehau\b", r"\bveka\b", r"\binternorm\b", r"\bgealan\b",
        r"\bpvc[\s\-]?ablak",
        r"\bbejarat[oi][\s\-]?ajto",
        r"\bfa[\s\-]?ablak\b",
    ]),
    # ── 2. Klíma ─────────────────────────────────────────────────────────────────
    ("klíma", [
        r"\bklima\b",                         # Klíma Kft, Klíma-Tech
        r"\blegkondi\b", r"\blegkondicionalo",
        r"comfort[\s\-]?air", r"cool[\s\-]?air", r"fresh[\s\-]?air",
        r"klima[\s\-]?(szere|bepites|telepites|rendszer|technik|center|kft|bt|hu)",
        r"\bdaikin\b",
        r"\bkaisai\b",
        r"heat[\s\-]?pump[\s\-]?(kft|bt|szere|szol)",
        r"\bhoszivatty",                      # hőszivattyú → hoszivattu after norm
        r"\bac[\s\-]?serv\b", r"\bac[\s\-]?tech\b", r"\bac[\s\-]?kft\b",
        r"\bhvac\b",
        r"\bsplit[\s\-]?(szere|kft|bt|air)",
    ]),
    # ── 2b. Tüzeléstechnikai / gázégő szerviz ──────────────────────────────────────
    # Narrower than fűtésszerelő — a real rejection confirmed the gap: "Gázégő
    # Tüzeléstechnikai Bolt és Szerviz" was pitched as "fűtésszerelő cégeknek"
    # and replied "nem fűtésszereléssel foglalkozunk hanem gázkazán szereléssel"
    # (we don't do heating installation, we do gas boiler/burner work) — this
    # company self-identifies as a narrow burner/combustion specialist, not a
    # general heating installer, and neither _NAME_RULES nor _TEXT_KEYWORDS had
    # any entry for "gázégő"/"tüzeléstechnika" at all, so it fell through to
    # the generic "kazán" (weight 3) text-scoring match for fűtésszerelő —
    # exactly clearing _MIN_SCORE with nothing to distinguish it. Must be
    # checked BEFORE fűtésszerelő below (first match wins).
    ("tüzeléstechnikai", [
        r"\bgazego\b",                         # Gázégő ...
        r"\btuzeleztechnik",                   # Tüzeléstechnika, Tüzeléstechnikai
        r"\begoszerviz\b",                     # Égőszerviz
        r"\bkazanszerviz\b",                   # Kazánszerviz (as a standalone shop/service name)
    ]),
    # ── 3. Fűtésszerelő ──────────────────────────────────────────────────────────
    ("fűtésszerelő", [
        r"\bfutesszere",                      # fűtésszerelő, fűtésszerelés
        r"\bgazszerelo\b",                    # gázszerelő
        r"\bgaz[\s\-]?mester\b",
        r"\bviz[\s\-]?gaz\b", r"\bvizgaz\b",  # Víz-Gáz, Vízgáz
        r"\bepuletgepesz",                    # épületgépész (Zengl-type companies)
        r"\bfutomester\b", r"\bfutoszerelo\b",
        r"\bkazanszere", r"\bkazan[\s\-]?(szere|technik|kft)",
        r"futes[\s\-]?(rendszer|szere|szol|mester|kft|bt|technik)",
        r"gaz[\s\-]?(szere|vezetek|bekotes|mester|kft|bt)",
        r"\bvaillant\b", r"\bviessmann\b", r"\bbuderus\b",
        r"viz[\s\-]?futes",                   # Víz-Fűtés Kecskemét ← ground truth case
        r"\bhoközpont\b", r"\bhoenergetik",
    ]),
    # ── 4. Napelemes ─────────────────────────────────────────────────────────────
    ("napelemes", [
        r"\bnapelem",
        r"\bnapkollekt",
        r"\bphotovolt\b", r"\bpv[\s\-]?(rendszer|szere|kft|bt)",
        r"solar[\s\-]?(rendszer|power|kft|bt|technik|energy|pro|hu|epito)",
        r"\bsolar\b",                         # SolarPro Kft, Solar Kft
        r"nap[\s\-]?(energia|panel|rendszer)",
        r"\blongi\b", r"\bjinko\b", r"\bfronius\b",
        r"\bsma[\s\-]?(inverter|solar)",
    ]),
    # ── 5. Tetőfedő ──────────────────────────────────────────────────────────────
    ("tetőfedő", [
        r"\btetofed",
        r"\btetoszere",
        r"\bteto[\s\-]?(buri|fed|szere|epitesz|munk)",
        r"\bacs[\s\-]?(mester|kft|bt)\b",
        r"\bacs\b(?![\w\-])",                 # standalone ács (not part of "asztalos")
        r"\bbramac\b", r"\btondach\b", r"\beternit\b", r"\bcreaton\b", r"\bicopal\b",
        r"\bkorcolt[\s\-]?lemez",
    ]),
    # ── 6. Villanyszerelő ────────────────────────────────────────────────────────
    ("villanyszerelő", [
        r"\bvillanyszere",
        r"\bvillszerelo\b",
        r"\belektromos[\s\-]?(szere|halozat|kivitele|mester|kft|bt)",
        r"\belectro\b", r"\belektro[\s\-]?(kft|bt|mester|szere|technik)",
        r"\bnapelemszere",
        r"\bvillamossagi\b",
    ]),
    # ── 7. Burkoló ───────────────────────────────────────────────────────────────
    ("burkoló", [
        r"\bburkolo\b",
        r"burkol[aá](s|si|t)[\s\-]?(kft|bt|mester|munk|szol)",
        r"\bcsempe[\s\-]?(rako|lerak|kft|bt|mester)",
        r"\bgranit[\s\-]?bur",
        r"\bpadloburkol",
    ]),
    # ── 8. Festő-mázoló ──────────────────────────────────────────────────────────
    ("festő-mázoló", [
        r"\bfesto[\s\-]?(mester|kft|bt|munk|szol|ceg|mes)\b",
        r"\bmazolo\b",
        r"\bfestes[\s\-]?(kft|mester|bt|szol)",
        r"\bhomlokzatfestes\b",
        r"\bszobafesto\b",
        r"festomester|mazolomester",
    ]),
    # ── 9. Kőműves ───────────────────────────────────────────────────────────────
    ("kőműves", [
        r"\bkomuves[\s\-]?(mester|kft|bt|munk)?\b",
        r"\bfalaz[oa]\b",
        r"\bvakolos[\s\-]?(kft|bt|mester|munk)",
        r"\btegla[\s\-]?falaz",
    ]),
    # ── 10. Gipszkartonos ────────────────────────────────────────────────────────
    ("gipszkartonos", [
        r"\bgipszkarton",
        r"\bszarazepites",
        r"\balmenyezet[\s\-]?(epites|kft|bt)",
        r"\bknauf\b", r"\brigips\b",
    ]),
    # ── 11. Asztalos ─────────────────────────────────────────────────────────────
    ("asztalos", [
        r"\basztalos\b",                      # Asztalos Bútor Bt, standalone
        r"\basztalos[\s\-]?(mester|kft|bt|munk|szol|ipari)\b",
        r"\bbutoraszt",
        r"\begyedi[\s\-]?butor",
        r"\bbeepitett[\s\-]?butor",
        r"\bkonyhabutor[\s\-]?(keszites|gyart|kft|bt)",
    ]),
    # ── 12. Vízvezetékszerelő ────────────────────────────────────────────────────
    ("vízvezetékszerelő", [
        r"\bvizvezetekszere",
        r"\bvizszere[\s\-]?(mester|kft|bt)\b",
        r"\bcsoszere[\s\-]?(kft|bt|mester)\b",
        r"\bszaniter[\s\-]?(kft|bt|szere|mester)",
    ]),
    # ── 13. Garázskapu és árnyékoló ──────────────────────────────────────────────
    ("garázskapu és árnyékoló", [
        r"\bgarazskapu\b",
        r"\bgarazs[\s\-]?(kapu|ajto)",
        r"\bsomfy\b",
        r"\bhormann\b",
        r"\bkapuautomat",
        r"\btologapu\b",
        r"\bszekcionalt[\s\-]?kapu",
        r"\bzip[\s\-]?arnyek",
        r"\bpergola\b",
        r"\brolltech\b",
    ]),
    # ── 14. Térkövező ────────────────────────────────────────────────────────────
    ("térkövező", [
        r"\bterko\b",                         # térkő → terko after norm
        r"\bterkovez",
        r"\bfelhajto\b",                      # felhajtó → felhajto
        r"\bsemmelrock\b", r"\bbradstone\b", r"\bmarshalls\b",
    ]),
    # ── 15. Homlokzatszigetelő ───────────────────────────────────────────────────
    ("homlokzatszigetelő", [
        r"\bhomlokzat[\s\-]?(szigetel|felujit|burk|vakolos)",
        r"\bhoszigetel\b",
        r"\bkulso[\s\-]?szigetel",
        r"\bdryvit\b",
        r"\bnemesvakolat\b",
        r"\bstyropor\b",
    ]),
    # ── 16. Kertépítő ────────────────────────────────────────────────────────────
    ("kertépítő", [
        r"\bkertepites\b", r"\bkertepito\b",
        r"\bkertrendez",
        r"\bkerteszet[\s\-]?(kft|bt|szol)",
        r"\bontozorendszer\b",
        r"\bfunyiro[\s\-]?(kft|bt|szol|munk)",
    ]),
    # ── 17. Terasz és deszka ─────────────────────────────────────────────────────
    ("terasz és deszka", [
        r"\bwpc[\s\-]?(terasz|deszka|burk|kft|bt)",
        r"\bterasz[\s\-]?(epites|burk|deszka|kft|bt)",
        r"\bfaterasz\b",
        r"\bbangkirai\b",
        r"kompozit[\s\-]?deszka",
    ]),
    # ── 18. Medenceépítő ─────────────────────────────────────────────────────────
    ("medenceépítő", [
        r"\bmedence[\s\-]?(epites|felujit|kft|bt|szere)",
        r"\buszoda[\s\-]?(epites|felujit)",
        r"\bjacuzzi[\s\-]?(epites|kft|bt)",
        r"\bpool[\s\-]?(build|kft|bt|construct|technik)",
    ]),
    # ── 19. Padlóburkoló ─────────────────────────────────────────────────────────
    ("padlóburkoló", [
        r"\bparketta[\s\-]?(lera|fek|csisz|kft|bt|mester)\b",
        r"\blaminalt[\s\-]?(padlo|fektet)",
        r"\bvinyl[\s\-]?(padlo|fektet)",
    ]),
    # ── 20. Fürdőszoba felújító ──────────────────────────────────────────────────
    ("fürdőszoba felújító", [
        r"\bfurdo[\s\-]?(szoba[\s\-]?)?(felujit|atalak|korszerus)",
        r"\bfurdoszoba[\s\-]?(kft|bt|szere|tervez|centr)",
    ]),
    # ── 21-23. NON-TRADE niches ──────────────────────────────────────────────────
    # Appended at the END on purpose: first match wins, so every existing trade
    # rule gets first refusal. These three are not construction trades — they
    # produce no finished physical artifact to photograph, which is why they
    # also need the text-sourced opener path in dm_app_hu.py rather than the
    # vision "compliment a craft detail" one. See _HU_TEXT_OPENER_HOOKS there.
    ("költöztető", [
        r"\bkoltozte",                        # költöztetés, költöztető
        r"\bkoltozes\b",
        r"\bbutorszallit",                    # bútorszállítás
        r"\bfuvaroz",                         # fuvarozás (name-only, safe)
        r"\blomtalanit",
        r"\bteherfuvar",
        r"\bzongoraszallit",
        r"\bmoving[\s\-]?(kft|bt|hu|team|service)",
        r"\btrans[\s\-]?(koltoz|butor)",
    ]),
    ("fogászat", [
        r"\bfogasz",                          # fogászat, fogászati
        r"\bfogorvos",
        r"\bdental\b", r"\bdent\b",           # X-Dent, Dental Kft
        r"\bimplantolog",
        r"\bszajsebesz",                      # szájsebészet
        r"\bfogszabalyoz",
        r"\bfogtechnik",
        r"\bdentiszt",
    ]),
    ("raktárbérlő", [
        r"\braktarber",                       # raktárbérlés, raktárbérlő
        r"\braktar[\s\-]?(kiado|berbead|park|centrum|hu)",
        r"\bonkiszolgalo[\s\-]?raktar",
        r"\bself[\s\-]?storage\b",
        r"\bstorage\b",
        r"\btarolo[\s\-]?(ber|kiado|box|centrum)",
        r"\bkontener[\s\-]?(ber|tarol|kiado)",
    ]),
    # ── 24-29. Lookup-table pricing niches ───────────────────────────────────────
    # All six price off a small set of known variables rather than a site visit,
    # which is the cleanest possible fit for a quoting bot. All are text-sourced
    # (see _HU_NON_PHOTO_NICHES in dm_app_hu.py) — none has a "finished craft
    # job" photo, and two of them have photos that must never be commented on.
    # Still appended AFTER every trade rule so existing niches keep first refusal.
    ("autóüveg", [
        r"\bauto[\s\-]?uveg",                 # autóüveg, "Auto Uveg Centrum"
        r"\bszelvedo",                        # szélvédő, szélvédőjavítás
        r"\bcarglass\b", r"\bglass[\s\-]?(car|auto|service)",
        r"\buvegkar\b",
        r"\bkofelverod",                      # kőfelverődés
    ]),
    ("háztartásigép-szerviz", [
        r"\bhaztartasi[\s\-]?gep",
        r"\bgepjavit", r"\bgepszerviz\b",
        r"\bmosogep[\s\-]?(szerviz|javit)",
        r"\bhuto[\s\-]?(szerviz|javit|gepesz)",
        r"\bmosogatogep[\s\-]?(szerviz|javit)",
        r"\bbeepitheto[\s\-]?gep",
        r"\bwhite[\s\-]?goods\b",
    ]),
    ("telefon- és laptopszerviz", [
        r"\btelefonszerviz\b", r"\bmobilszerviz\b",
        r"\blaptopszerviz\b", r"\bszamitogep[\s\-]?szerviz",
        r"\bgsm[\s\-]?szerviz", r"\bgsm[\s\-]?javit",
        r"\biphone[\s\-]?(szerviz|javit)", r"\bmacbook[\s\-]?(szerviz|javit)",
        r"\bkijelzocsere\b",
        r"\bmobiltelefon[\s\-]?javit",
    ]),
    ("fordítóiroda", [
        r"\bfordito[\s\-]?(iroda|kft|bt|szolg)",
        r"\bforditas[\s\-]?(kft|bt|iroda|szolg)",
        r"\bszakfordit",
        r"\btolmacs",
        r"\btranslat",                        # Translation, Translations Kft
        r"\bnyelvi[\s\-]?(szolg|kozpont|iroda)",
    ]),
    ("autókozmetika", [
        r"\bautokozmetik",
        r"\bcar[\s\-]?detail", r"\bdetailing\b",
        r"\bautomoso\b", r"\bkezi[\s\-]?autómoso", r"\bkezi[\s\-]?automoso",
        r"\bpolirozas\b",
        r"\bkeramia[\s\-]?bevonat",
        r"\bautofoliaz",
    ]),
    ("hajbeültetés", [
        r"\bhajbeultet",
        r"\bhajatultet",                      # hajátültetés
        r"\bhair[\s\-]?transplant",
        r"\bhajklinik",
        r"\bhajgyogyasz",
    ]),
]


def _hu_detect_niche_from_name(name: str) -> str:
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
    "nyílászáró": [
        ("nyílászáró", 4), ("nyilaszaro", 4), ("ablakcsere", 4), ("ablakbeépítés", 4),
        ("ablak", 3), ("pvc ablak", 4), ("fa ablak", 3),
        ("schüco", 4), ("schuco", 4), ("rehau", 4), ("veka", 4),
        ("internorm", 4), ("gealan", 4), ("kbe ablak", 4),
        ("bejárati ajtó", 3), ("ablakszerviz", 4), ("ablakfelújítás", 4),
    ],
    "klíma": [
        ("klíma", 4), ("klima", 4), ("klimaszerelő", 4), ("légkondicionáló", 4),
        ("hőszivattyú", 4), ("hoszivattu", 4),
        ("split", 3), ("daikin", 4), ("mitsubishi electric", 4),
        ("kaisai", 4), ("midea", 3), ("gree", 3),
        ("panasonic klíma", 4), ("lg klíma", 4), ("samsung klíma", 4),
        ("klíma telepítés", 4), ("klíma beépítés", 4),
    ],
    "tüzeléstechnikai": [
        ("gázégő", 4), ("gazego", 4), ("tüzeléstechnika", 4), ("tuzeleztechnika", 4),
        ("tüzeléstechnikai", 4), ("égőszerviz", 4), ("egoszerviz", 4),
        ("kazánszerviz", 4), ("kazanszerviz", 4), ("kazáncsere", 3), ("kazancsere", 3),
    ],
    "fűtésszerelő": [
        ("fűtésszerelő", 4), ("futesszerelo", 4), ("gázszerelő", 4), ("gazszerelo", 4),
        ("víz-gáz", 4), ("viz-gaz", 4), ("vízgáz", 4),
        ("épületgépész", 4), ("epuletgepesz", 4),
        ("kazán", 3), ("kazan", 3), ("padlófűtés", 4), ("padlofutes", 4),
        ("gázvezeték", 3), ("gazvezetek", 3),
        ("vaillant", 3), ("viessmann", 3), ("buderus", 3),
        ("víz-fűtés", 4), ("viz-futes", 4),
        ("hőközpont", 4), ("hokozpont", 4), ("hőenergetika", 4),
    ],
    "napelemes": [
        ("napelem", 4), ("napelemrendszer", 4), ("napkollekt", 4),
        ("pv rendszer", 4), ("photovoltaic", 4),
        ("longi", 4), ("jinko", 4), ("fronius", 4), ("sma inverter", 4),
        ("canadian solar", 4), ("risen", 3), ("trina", 3),
    ],
    "tetőfedő": [
        ("tetőfedő", 4), ("tetofedo", 4), ("tetőfedés", 4),
        ("korcolt", 4), ("oromszegély", 4), ("ereszcsatorna", 3),
        ("bramac", 4), ("tondach", 4), ("eternit", 4), ("creaton", 4), ("icopal", 4),
        ("velux", 3), ("fakro", 3), ("lapostető", 3), ("héjazat", 4),
    ],
    "villanyszerelő": [
        ("villanyszerelő", 4), ("villanyszere", 4), ("elektromos szerelés", 4),
        ("kismegszakító", 4), ("elosztó szekrény", 3), ("áramkör", 3),
    ],
    "burkoló": [
        ("burkoló", 4), ("burkolás", 4),
        ("csempe", 3), ("járólap", 3), ("gránit burkolat", 4),
        ("mozaik burkolat", 4), ("fürdőszoba burk", 4),
    ],
    "festő-mázoló": [
        ("festő-mázoló", 4), ("szobafestő", 4),
        ("homlokzatfestés", 4), ("beltéri festés", 4),
        ("falfesték", 3), ("diszperzit", 3),
    ],
    "kőműves": [
        ("kőműves", 4), ("komuves", 4), ("falazó mester", 4),
        ("kőművesmunka", 4), ("téglafalazás", 4),
    ],
    "gipszkartonos": [
        ("gipszkarton", 4), ("szárazépítés", 4), ("szarazepites", 4),
        ("álmennyezet", 4), ("knauf", 4), ("rigips", 4),
    ],
    "asztalos": [
        ("asztalos", 4), ("bútorasztalos", 4), ("egyedi bútor", 4),
        ("beépített bútor", 4), ("konyhabútor", 3),
    ],
    "vízvezetékszerelő": [
        ("vízvezetékszerelő", 4), ("vízszerelő", 4), ("vizszerelo", 4),
        ("csőszerelő", 4), ("szaniter", 3),
    ],
    "garázskapu és árnyékoló": [
        ("garázskapu", 4), ("garazskapu", 4), ("kapuautomatika", 4),
        ("tolókapu", 4), ("szekcionált kapu", 4), ("somfy", 4),
        ("raffstore", 4), ("zip árnyékoló", 4), ("pergola", 3),
        ("redőny szerelő", 3),
    ],
    "térkövező": [
        ("térkő", 4), ("terko", 4), ("térkövezés", 4),
        ("semmelrock", 4), ("bradstone", 4), ("marshalls", 4),
        ("kocsibeálló burkolat", 4),
    ],
    "homlokzatszigetelő": [
        ("homlokzatszigetelés", 4), ("külső szigetelés", 4),
        ("hőszigetelés", 3), ("nemesvakolat", 4), ("dryvit", 4), ("styropor", 3),
    ],
    "kertépítő": [
        ("kertépítés", 4), ("kertepites", 4), ("kertrendezés", 4),
        ("öntözőrendszer", 4), ("gyepszőnyeg", 3),
    ],
    "terasz és deszka": [
        ("wpc terasz", 4), ("wpc deszka", 4), ("bangkirai", 4),
        ("kompozit deszka", 4), ("teraszburkolat", 4), ("faterasz", 4),
    ],
    "medenceépítő": [
        ("medenceépítés", 4), ("medenceepites", 4), ("úszómedence", 4),
        ("jacuzzi", 3), ("wellness medence", 4),
    ],
    "padlóburkoló": [
        ("parketta", 3), ("laminált padló", 4), ("laminalt padlo", 4),
        ("padlóburkolás", 4), ("vinyl padló", 4),
    ],
    "fürdőszoba felújító": [
        ("fürdőszoba felújítás", 4), ("furdoszoba felujitas", 4),
        ("komplett fürdő", 4), ("fürdőszoba átalakítás", 4),
    ],
    # ── Non-trade niches ─────────────────────────────────────────────────────
    "költöztető": [
        ("költöztetés", 4), ("koltoztetes", 4), ("költöztető", 4),
        ("bútorszállítás", 4), ("butorszallitas", 4),
        ("lomtalanítás", 4), ("lomtalanitas", 4),
        ("zongoraszállítás", 4), ("zongoraszallitas", 4),
        ("irodaköltöztetés", 4), ("irodakoltoztetes", 4),
        ("teherfuvarozás", 3), ("fuvarozás", 3),
        ("csomagolóanyag", 3), ("bútorszerelés", 3),
    ],
    "fogászat": [
        ("fogászat", 4), ("fogaszat", 4), ("fogorvos", 4),
        ("implantátum", 4), ("implantatum", 4), ("implantológia", 4),
        ("all-on-4", 4), ("all on 4", 4),
        ("porcelán héj", 4), ("porcelan hej", 4), ("héjak", 3),
        ("fogszabályozás", 4), ("invisalign", 4),
        ("szájsebészet", 4), ("szajsebeszet", 4),
        ("fogtechnikai labor", 4), ("koronák", 3), ("fogpótlás", 4),
        ("cerec", 4), ("digitális lenyomat", 4),
    ],
    "raktárbérlő": [
        ("raktárbérlés", 4), ("raktarberles", 4), ("raktárbérlet", 4),
        ("önkiszolgáló raktár", 4), ("onkiszolgalo raktar", 4),
        ("self storage", 4), ("tárolóbox", 4), ("tarolobox", 4),
        ("raktárkapacitás", 4), ("konténerbérlés", 4),
        ("fűtött raktár", 4), ("futott raktar", 4),
        ("logisztikai csarnok", 3), ("raktározás", 3),
    ],
    # ── Lookup-table pricing niches ──────────────────────────────────────────
    # Weights kept honest: 4/3 only for terms that belong to exactly one of
    # these niches. Generic service words ("garancia", "alkatrész", "sürgős")
    # are 1-2 deliberately — dm_app_hu.py's _SPECIFIC_TRADE_TERMS harvests
    # every weight>=3 keyword from EVERY niche into the corroboration guard,
    # so promoting a common word here would start demanding corroboration for
    # it in the openers of unrelated trades.
    "autóüveg": [
        ("szélvédő", 4), ("szelvedo", 4), ("autóüveg", 4), ("autouveg", 4),
        ("kőfelverődés", 4), ("kofelverodes", 4),
        ("szélvédőjavítás", 4), ("szélvédőcsere", 4),
        ("oldalüveg", 4), ("hátsó szélvédő", 4),
        ("kamerakalibrálás", 4), ("adas kalibrál", 4),
        ("kaszkó", 2), ("casco", 2),
    ],
    "háztartásigép-szerviz": [
        ("háztartási gép", 4), ("haztartasi gep", 4),
        ("mosógépjavítás", 4), ("mosogepjavitas", 4),
        ("hűtőgépjavítás", 4), ("hutogepjavitas", 4),
        ("mosogatógép javítás", 4), ("szárítógép", 3),
        ("beépíthető gép", 4), ("gépjavítás", 3),
        ("mosógép", 2), ("hűtőszekrény", 2), ("alkatrész", 1),
    ],
    "telefon- és laptopszerviz": [
        ("kijelzőcsere", 4), ("kijelzocsere", 4),
        ("akkumulátorcsere", 4), ("alaplapjavítás", 4),
        ("mobilszerviz", 4), ("telefonszerviz", 4), ("laptopszerviz", 4),
        ("gsm szerviz", 4), ("adatmentés", 3),
        ("iphone javítás", 4), ("macbook javítás", 4),
        ("ázott telefon", 4),
    ],
    "fordítóiroda": [
        ("szakfordítás", 4), ("szakforditas", 4),
        ("hiteles fordítás", 4), ("hiteles forditas", 4),
        ("tolmácsolás", 4), ("tolmacsolas", 4),
        ("lektorálás", 4), ("lektoralas", 4),
        ("fordítóiroda", 4), ("forditoiroda", 4),
        ("nyelvpár", 4), ("honlapfordítás", 4),
        ("fordítás", 2), ("leütés", 2),
    ],
    "autókozmetika": [
        ("autókozmetika", 4), ("autokozmetika", 4),
        ("kerámiabevonat", 4), ("keramiabevonat", 4),
        ("polírozás", 4), ("polirozas", 4),
        ("kárpittisztítás", 4), ("karpittisztitas", 4),
        ("ózongenerátoros", 4), ("motormosás", 3),
        ("autófóliázás", 4), ("belső tisztítás", 3),
        ("viaszolás", 3),
    ],
    "hajbeültetés": [
        ("hajbeültetés", 4), ("hajbeultetes", 4),
        ("hajátültetés", 4), ("hajatultetes", 4),
        ("fue technika", 4), ("dhi technika", 4),
        ("graft", 4), ("szakállbeültetés", 4),
        ("szemöldökbeültetés", 4), ("hajklinika", 4),
        ("prp kezelés", 3),
    ],
}

_GENERAL_SIGNALS = {
    "generálkivitelező", "generalkivitelez",
    "általános kivitelező", "altalanos kivitelez",
    "komplex kivitelezés", "komplex kivitelezes",
    "komplex felújítás", "komplex felujitas",
    "teljes körű felújítás", "teljes koru felujitas",
    "építési vállalkozás", "epitesi vallalkozas",
    "kivitelezési vállalkozás", "kivitelezesi vallalkozas",
    "általános építési", "altalanos epitesi",
    "komplex építési", "komplex epitesi",
}

_MIN_SCORE  = 3   # must score at least this to claim a niche
_MIN_MARGIN = 2   # must beat second place by this much


def _score_text(text_data: str) -> str:
    t_orig = text_data.lower()
    t_norm = _norm(text_data)

    # Immediate general contractor detection
    for sig in _GENERAL_SIGNALS:
        if sig in t_orig or sig in t_norm:
            return "felújító kivitelező"

    scores: dict[str, int] = {}
    for niche, kw_list in _TEXT_KEYWORDS.items():
        s = sum(w for kw, w in kw_list if _norm(kw) in t_norm)
        if s > 0:
            scores[niche] = s

    if not scores:
        return _DEFAULT_NICHE

    # 4+ different trades scored → general contractor
    if len(scores) >= 4:
        return "felújító kivitelező"

    top = sorted(scores.items(), key=lambda x: -x[1])
    best_niche, best_score = top[0]

    # Minimum threshold
    if best_score < _MIN_SCORE:
        return _DEFAULT_NICHE

    # Margin guard: if two niches are too close, fall back to default
    if len(top) >= 2 and (best_score - top[1][1]) < _MIN_MARGIN:
        return _DEFAULT_NICHE

    return best_niche


def _hu_detect_niche(text_data: str, company_name: str = "") -> str:
    """
    Primary entry point. Name-first → text scoring → default.
    """
    if company_name:
        name_niche = _hu_detect_niche_from_name(company_name)
        if name_niche:
            return name_niche
    if text_data and text_data.strip():
        return _score_text(text_data)
    return _DEFAULT_NICHE


# ── Per-niche opener instructions ────────────────────────────────────────────────
# Injected into the prompt before the vision call so the model knows what to look for.

NICHE_OPENER_INSTRUCTIONS: dict[str, str] = {
    "nyílászáró": (
        "FONTOS KONTEXTUS: Ez egy ablak / nyílászáró cég. "
        "Ha a képen terasz, WPC deszka, kert, klíma kültéri egység, napelem vagy kő látható, "
        "TELJESEN FIGYELMEN KÍVÜL HAGYOD — ezek más cégek fotói véletlenszerűen megjelenve. "
        "Az ablak/nyílászáró fotón a következőket keress: "
        "profilszín (antracit, fehér, aranytölgy, sötétbarna), "
        "nyitástípus (bukó-nyíló, toló/emelt-toló, harmonika szárny), "
        "tokreveal mélysége (mélyen beépített vs. lapos), "
        "külső párkány dőlésszöge (lejtés a vízvezetéshez), "
        "fal-tok csatlakozás (szalagos-e), "
        "bejárati ajtónál küszöb és padlóburkolat illesztése, "
        "tetőtéri ablaknál villogópajzs visszavezetése, "
        "integrált redőnytok: beépített-e a falba. "
        "Ha nem látható ablak/ajtó munka → FALLBACK."
    ),
    "klíma": (
        "FONTOS KONTEXTUS: Ez egy klímaszerelő cég. "
        "Ha a képen terasz, burkolat, kert, ablak, tető, kő vagy kazán látható, "
        "FIGYELMEN KÍVÜL HAGYOD. "
        "Klíma fotón keress: Daikin / Mitsubishi / KAISAI / Midea / Gree / Panasonic / LG / Samsung brand, "
        "kültéri egység elhelyezése, csőelvezetés minősége, "
        "beltéri egység pozíciója és típusa. "
        "Ha nem klíma/split/hőszivattyú munka van a képen → FALLBACK."
    ),
    "fűtésszerelő": (
        "FONTOS KONTEXTUS: Ez egy fűtés- és gázszerelő cég. "
        "Ha a képen terasz, burkolat, kert, ablak, tető, klíma kültéri egység látható, "
        "FIGYELMEN KÍVÜL HAGYOD. "
        "Fűtésszerelő fotón keress: "
        "Vaillant / Viessmann / Bosch / Buderus / Junkers / Baxi / Ariston kazán, "
        "csőelvezetés az aljzaton, padlófűtés csőrendszer, radiátor bekötés, "
        "gázmérő szekrény, hőközpont. "
        "Ha nem kazán/fűtésrendszer munka van a képen → FALLBACK."
    ),
    "tüzeléstechnikai": (
        "FONTOS KONTEXTUS: Ez egy gázégő / tüzeléstechnikai szakcég, NEM általános "
        "fűtésszerelő — szűkebb szakterület: égők, kazánok szervize és cseréje. "
        "Ha a képen terasz, burkolat, kert, ablak, tető, klíma kültéri egység látható, "
        "FIGYELMEN KÍVÜL HAGYOD. "
        "Fotón keress: gázégő beépítés, kazán csere vagy javítás, csőelvezetés, "
        "égéstermék-elvezető, szerviz munka közben. "
        "Ha nem égő/kazán szerviz munka van a képen → FALLBACK."
    ),
    "napelemes": (
        "FONTOS KONTEXTUS: Ez egy napelemes cég. "
        "Ha a képen kert, terasz, ablak, klíma vagy kazán látható, FIGYELMEN KÍVÜL HAGYOD. "
        "Napelem fotón keress: LONGi / Jinko / Canadian Solar / Risen / Trina panel, "
        "inverter (Huawei / SMA / Fronius), "
        "panel szimmetria és dőlésszög egyenletessége, "
        "kábelvezetés, tartószerkezet. "
        "Ha nem napelemes rendszer van a képen → FALLBACK."
    ),
    "tetőfedő": (
        "FONTOS KONTEXTUS: Ez egy tetőfedő / ács cég. "
        "Ha a képen klíma, terasz, kert vagy ablak látható, FIGYELMEN KÍVÜL HAGYOD. "
        "Tetőfedő fotón keress: Bramac / Tondach / Eternit / Creaton cserép, "
        "korcolt fémtető, EPDM lapostető, Velux / Fakro tetőtéri ablak, "
        "oromszegély vonala, gerinccsempe illesztés, ereszcsatorna lejtése, "
        "kéménykörüli villásszeg. Ha nem tető van a képen → FALLBACK."
    ),
    "villanyszerelő": (
        "FONTOS KONTEXTUS: Ez egy villanyszerelő cég. "
        "Villanyszerelő fotón keress: elosztószekrény, kismegszakítók rendje, "
        "kábel- és csőelrendezés, napelemes inverter bekötés, épületautomatika. "
        "Ha nem elektromos szerelési munka van a képen → FALLBACK."
    ),
    "burkoló": (
        "FONTOS KONTEXTUS: Ez egy burkoló cég. "
        "Burkoló fotón keress: porcelán / gránit / mészkő lap, "
        "fugaköz egyenletessége, padló-fal átmenet, "
        "mintaillesztés (szálkás, kötött, diagonális), mozaikburkolat, szegélycsempe. "
        "Ha nem burkolat van a képen → FALLBACK."
    ),
    "festő-mázoló": (
        "FONTOS KONTEXTUS: Ez egy festő-mázoló cég. "
        "Festő fotón keress: homlokzatfestés egyenletes lefedése, szín, "
        "ajtókeretek körüli elvágás pontossága, beltéri fal-mennyezet találkozás, "
        "színátmenet minősége. Ha nem festési munka van a képen → FALLBACK."
    ),
    "kőműves": (
        "FONTOS KONTEXTUS: Ez egy kőműves cég. "
        "Kőműves fotón keress: téglafal minősége, vakolatfelület, "
        "sarkélek egyenessége, fugaköz, átmenetek. "
        "Ha nem kőműves munka van a képen → FALLBACK."
    ),
    "gipszkartonos": (
        "FONTOS KONTEXTUS: Ez egy gipszkartonos / szárazépítő cég. "
        "Fotón keress: álmennyezet kivágások, CD/UD profil rendszer, "
        "hézagolás, sarok illesztések. Ha nem szárazépítési munka van → FALLBACK."
    ),
    "asztalos": (
        "FONTOS KONTEXTUS: Ez egy asztalos / bútoripari cég. "
        "Fotón keress: egyedi bútor típusa, fa felület, pántok, fogantyúk, "
        "illesztések pontossága. Ha nem bútorasztalos munka van → FALLBACK."
    ),
    "vízvezetékszerelő": (
        "FONTOS KONTEXTUS: Ez egy vízvezetékszerelő / szaniter cég. "
        "Fotón keress: csőrendszer elrendezése, szivattyú, csaptelep, "
        "szaniter felszerelés. Ha nem vízvezetékes munka van → FALLBACK."
    ),
    "garázskapu és árnyékoló": (
        "FONTOS KONTEXTUS: Ez egy garázskapu / árnyékoló cég. "
        "Ha a képen ablak, tető, kert vagy klíma látható, FIGYELMEN KÍVÜL HAGYOD. "
        "Fotón keress: szekcionált / harmonika kapu típusa, Somfy / Hörmann motor, "
        "zip árnyékoló, raffstore, pergola. Ha nem garázskapu/árnyékoló van → FALLBACK."
    ),
    "térkövező": (
        "FONTOS KONTEXTUS: Ez egy térkövező cég. "
        "Térkő fotón keress: Semmelrock / Bradstone / Marshalls kő típusa, "
        "fugaköz egyenletessége, szegélykő illesztése, mintaelrendezés. "
        "Ha nem térkő/felhajtó van a képen → FALLBACK."
    ),
    "homlokzatszigetelő": (
        "FONTOS KONTEXTUS: Ez egy homlokzatszigetelő cég. "
        "Fotón keress: dryvit / nemesvakolat felület, sarkélek, "
        "ablakkeretek körüli simítás, rögzítő hálórendszer. Ha nem homlokzat van → FALLBACK."
    ),
    "kertépítő": (
        "FONTOS KONTEXTUS: Ez egy kertépítő cég. "
        "Fotón keress: öntözőrendszer, tereplépcsők vonala, "
        "gyepszőnyeg egységessége, növényzetrendezés, kerítés. "
        "Ha nem kertészeti munka van a képen → FALLBACK."
    ),
    "terasz és deszka": (
        "FONTOS KONTEXTUS: Ez egy terasz / WPC deszka cég. "
        "Fotón keress: WPC / bangkirai / kompozit deszka típusa és színe, "
        "deszkatávolság egyenletessége, szegélyléc illesztése, tartószerkezet. "
        "Ha nem terasz/deszka van a képen → FALLBACK."
    ),
    "medenceépítő": (
        "FONTOS KONTEXTUS: Ez egy medenceépítő cég. "
        "Fotón keress: medence burkolat típusa és színe, szűrőrendszer, "
        "lépcső kialakítása, peremelemek. Ha nem medence van a képen → FALLBACK."
    ),
    "padlóburkoló": (
        "FONTOS KONTEXTUS: Ez egy padlóburkoló cég. "
        "Fotón keress: parketta/laminált típusa, illesztési minta, "
        "szegélyléc, padló-fal találkozás. Ha nem padlóburkolat van → FALLBACK."
    ),
    "fürdőszoba felújító": (
        "FONTOS KONTEXTUS: Ez egy fürdőszoba felújító cég. "
        "Fotón keress: csempe/lap típusa, szaniter elrendezés, "
        "csap és zuhanyzó felszerelés, fugaköz. Ha nem fürdőszoba van → FALLBACK."
    ),
    "felújító kivitelező": (
        "FONTOS KONTEXTUS: Ez egy általános felújító / kivitelező cég. "
        "Különböző szakmákat végeznek. Reagálj arra a szakipari munkára amit látni a képen — "
        "legyen az tető, burkolat, festés, villanyszerelés vagy bármi más."
    ),
    # ── Non-trade niches ─────────────────────────────────────────────────────
    # These three never produce a "finished craft job" photo to compliment, so
    # the whole vision premise fails. dm_app_hu.py discards the model's opener
    # for these niches and builds a text-sourced one instead — these entries
    # are a backstop so that if that bypass is ever removed, the model falls
    # back to the safe city line rather than inventing a compliment.
    "költöztető": (
        "FONTOS KONTEXTUS: Ez egy költöztető cég, NEM szakipari kivitelező. "
        "Nincs 'elkészült munka' amit meg lehetne dicsérni — a fotón furgon, "
        "dobozok, rakodó csapat vagy egy bútor lesz. Ezekre NE reagálj "
        "szakmai dicsérettel, mert nevetségesen hangzik. "
        "HASZNÁLD A FALLBACKET."
    ),
    "fogászat": (
        "FONTOS KONTEXTUS: Ez egy fogászat, NEM szakipari cég. "
        "A fotón nagy eséllyel egy VALÓDI PÁCIENS arca vagy fogai láthatók. "
        "SOHA ne reagálj a páciens fogaira, mosolyára vagy az arcára — "
        "ez egy hideg üzenetben visszataszító és etikátlan. "
        "HASZNÁLD A FALLBACKET, kivétel nélkül."
    ),
    "raktárbérlő": (
        "FONTOS KONTEXTUS: Ez egy raktár- vagy tárolóbérléssel foglalkozó cég, "
        "NEM szakipari kivitelező. A fotón üres csarnok, tárolóbox, kapu vagy "
        "kerítés lesz — nincs benne szakmai részlet amit dicsérni lehetne. "
        "HASZNÁLD A FALLBACKET."
    ),
    # ── Lookup-table pricing niches ──────────────────────────────────────────
    # These never reach the vision model at all (dm_app_hu.py skips the whole
    # visual pipeline for them via SKIP_VISION_FOR), so these entries are a
    # backstop: if that skip is ever removed, the model must fall back rather
    # than invent a compliment.
    "autóüveg": (
        "FONTOS KONTEXTUS: Ez egy autóüveg / szélvédő cég, NEM szakipari "
        "kivitelező. A fotón autó, repedt szélvédő vagy szerelés közbeni kép "
        "lesz. Rendszámra, autó típusára vagy a tulajdonosra SOHA ne reagálj. "
        "HASZNÁLD A FALLBACKET."
    ),
    "háztartásigép-szerviz": (
        "FONTOS KONTEXTUS: Ez egy háztartási gép szerviz, NEM szakipari "
        "kivitelező. A fotón mosógép, hűtő vagy szerelő lesz — nincs elkészült "
        "munka amit dicsérni lehetne. HASZNÁLD A FALLBACKET."
    ),
    "telefon- és laptopszerviz": (
        "FONTOS KONTEXTUS: Ez egy telefon- vagy laptopszerviz, NEM szakipari "
        "kivitelező. A fotón készülék, alkatrész vagy szétszedett telefon lesz. "
        "HASZNÁLD A FALLBACKET."
    ),
    "fordítóiroda": (
        "FONTOS KONTEXTUS: Ez egy fordítóiroda, NEM szakipari kivitelező. "
        "Nincs fizikai munka amit meg lehetne nézni; a fotó szöveg, logó vagy "
        "irodai kép lesz. HASZNÁLD A FALLBACKET."
    ),
    "autókozmetika": (
        "FONTOS KONTEXTUS: Ez egy autókozmetika / detailing cég, NEM szakipari "
        "kivitelező. A fotón letisztított autó lesz. Rendszámra és a "
        "tulajdonosra SOHA ne reagálj. HASZNÁLD A FALLBACKET."
    ),
    "hajbeültetés": (
        "FONTOS KONTEXTUS: Ez egy hajbeültetéssel foglalkozó klinika, NEM "
        "szakipari cég. A fotón nagy eséllyel egy VALÓDI PÁCIENS fejbőre, "
        "hajvonala vagy előtte-utána képe látható. SOHA ne reagálj a páciens "
        "hajára, fejbőrére vagy külsejére — ez egy hideg üzenetben "
        "visszataszító és etikátlan. HASZNÁLD A FALLBACKET, kivétel nélkül."
    ),
}


# ── Opener post-validation ───────────────────────────────────────────────────────
# If opener contains these niche-contradicting terms → reject and use fallback.

# Shared by every non-trade niche (moving, dental, storage, and the six
# lookup-table niches). None of them has any legitimate reason to mention a
# construction material or detail, so one broad list serves them all — keeping
# six near-identical copies in sync by hand is how one of them ends up missing
# a term. Per-niche extras (patient references, number plates) are appended at
# the use site.
_NON_TRADE_CONTRADICT = [
    "terasz", "deszka", "wpc", "cserép", "cserep",
    "klíma", "klima", "kazán", "kazan", "napelem",
    "ablak", "nyilaszaro", "térkő", "terko",
    "burkolat", "fugakoz", "vakolat", "homlokzat", "tető", "teto",
]

_NICHE_VOCAB: dict[str, dict[str, list[str]]] = {
    "nyílászáró": {"contradict": [
        "terasz", "deszka", "wpc", "bangkirai",
        "klíma", "klima", "napelem",
        "kazán", "kazan", "vaillant", "viessmann",
        "cserép", "cserepe", "tető teteje",
        "térkő", "terko", "kertepites",
    ]},
    "klíma": {"contradict": [
        "terasz", "deszka", "wpc", "bangkirai",
        "ablak", "nyilaszaro",
        "kazán", "kazan", "vaillant", "viessmann", "buderus",
        "napelem", "cserép",
        "térkő", "terko", "kertepites",
    ]},
    "fűtésszerelő": {"contradict": [
        "terasz", "deszka", "wpc",
        "ablak", "nyilaszaro",
        "daikin", "mitsubishi", "kaisai", "midea", "gree",
        "napelem", "cserép",
        "térkő", "terko", "kertepites",
    ]},
    "napelemes": {"contradict": [
        "terasz", "deszka", "wpc",
        "ablak", "nyilaszaro",
        "klíma", "klima", "kazán", "kazan",
        "térkő", "terko", "kertepites",
    ]},
    "tetőfedő": {"contradict": [
        "terasz", "deszka", "wpc",
        "ablak", "nyilaszaro",
        "klíma", "klima", "kazán", "kazan",
        "térkő", "terko", "kertepites",
    ]},
    "garázskapu és árnyékoló": {"contradict": [
        "terasz", "wpc", "klíma", "klima",
        "napelem", "ablak", "nyilaszaro",
    ]},
    "térkövező": {"contradict": [
        "klíma", "klima", "kazán", "kazan",
        "napelem", "ablak", "nyilaszaro",
        "terasz deszka", "wpc",
    ]},
    "burkoló": {"contradict": [
        "terasz deszka", "wpc", "bangkirai",
        "klíma", "klima", "kazán", "kazan",
        "napelem", "cserép tető",
        "ablak nyilaszaro",
    ]},
    # Non-trade niches: ANY construction vocabulary in the opener is a
    # hallucination here, so the shared list is deliberately broad — there is
    # no legitimate reason a moving/dental/storage/repair opener mentions a
    # trade detail. Defined once as _NON_TRADE_CONTRADICT above.
    "költöztető": {"contradict": _NON_TRADE_CONTRADICT},
    "fogászat": {"contradict": _NON_TRADE_CONTRADICT + [
        # Never comment on the patient themselves.
        "mosoly", "fogsor", "foga ", "fogai",
    ]},
    "raktárbérlő": {"contradict": _NON_TRADE_CONTRADICT},
    # Same rule for every lookup-table niche: any construction vocabulary in
    # the opener is a hallucination, so they all share one list rather than
    # six near-identical copies. The two with patient/customer photos also ban
    # commenting on the person or their property.
    "autóüveg":                  {"contradict": _NON_TRADE_CONTRADICT + ["rendszám", "rendszam"]},
    "háztartásigép-szerviz":     {"contradict": _NON_TRADE_CONTRADICT},
    "telefon- és laptopszerviz": {"contradict": _NON_TRADE_CONTRADICT},
    "fordítóiroda":              {"contradict": _NON_TRADE_CONTRADICT},
    "autókozmetika":             {"contradict": _NON_TRADE_CONTRADICT + ["rendszám", "rendszam"]},
    "hajbeültetés": {"contradict": _NON_TRADE_CONTRADICT + [
        # Never comment on the patient themselves — same rule as fogászat.
        "hajad", "haja ", "hajvonal", "fejbőr", "fejbor", "kopasz",
    ]},
}


def _opener_matches_niche(opener: str, niche: str) -> bool:
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

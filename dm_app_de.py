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
from de_niche import (
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
CDP_PORT         = int(os.environ.get("CDP_PORT", "9224"))
DEFAULT_EXCEL    = os.environ.get("DEFAULT_EXCEL", r"C:\Users\pirin\OneDrive\Asztali gép\Austrian_Leads_100.xlsx")
GRAMMAR_CHECK    = False  # set to True in dm_app_hu.py
GREETING_PREFIX  = "Servus"
STRICT_NAME_DETECTION = True    # mirror HU: only greet with a confidently detected name
ALWAYS_GREET     = True          # mirror HU: always greet ("Servus. " when no name found)
GREETING_NAME_FMT = "Servus, {name}. "   # (legacy, unused) kept for reference
GREETING_SOLO_FMT = "Servus. "            # (legacy, unused) kept for reference


def _apply_greeting(opener: str, name_hint: str) -> str:
    """Prepend a rotating casual Austrian greeting. 'Servus' / 'Grüß dich' / 'Hey'
    read like a real person rather than a flat 'Hallo.', and it's the first thing
    shown in the Facebook message preview, so it earns the open. A greeting ending
    in . ! ? keeps the opener capitalised; a comma greeting flows into a lowercase
    opener ('Servus Wade, die Therme ist sauber geworden.')."""
    import random as _r
    if name_hint:
        g = _r.choice([
            f"Servus {name_hint}, ",
            f"Grüß dich {name_hint}, ",
            f"Hey {name_hint}, ",
            f"Hallo {name_hint}, ",
        ])
    else:
        g = _r.choice(["Servus, ", "Grüß dich, ", "Hey, ", "Hallo, "])
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
SOCIAL_PROOF = "hab schon ein paar Betriebe, die so eins nutzen"
CUSTOM_NAME_EXTRACTOR = None               # if set, overrides extract_first_name (callable: text -> str)
CUSTOM_PROMPT_BUILDER = None               # if set, called as (full_prompt, text_data) -> str before vision call
CUSTOM_OPENER_POSTPROCESS = None           # if set, called as (opener, text_data) -> str after generation
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
        return v if v in ("full", "stripped", "british") else "full"
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
Du schreibst die erste Zeile einer kalten Facebook-DM. Gib NUR die erste Zeile zurück — nichts anderes. Keine Erklärung, kein Emoji, kein Zusatztext. Schreib auf Deutsch.

Du siehst einen Screenshot der Facebook-Seite eines Handwerksbetriebs (AT/DE). Schau dir das Foto genau an. Reagier auf eine fertige Arbeit — so wie wenn du einem Kollegen schreibst, der das Handwerk wirklich kennt.

PRIORITÄT — FACHLICHES DETAIL:
Wenn du eines davon siehst, reagier darauf — das überzeugt weit mehr als ein allgemeines "schaut gut aus":
- Markenname (siehe Liste unten)
- Materialart (Feinsteinzeug, Naturstein, EPDM, WPC, Schiefer, Blech, Betonpflaster usw.)
- Ausführung / Qualität (Fugenbreite, Gefälle, Leitungsführung, Randstein, sauberer Mustermatch)

PAARUNG — ZWEI DETAILS, DIE ZUSAMMENPASSEN:
Wenn zwei Elemente im Foto sich offensichtlich ergänzen, nenn beide. Das ist das Detail, das nur jemand bemerkt, der wirklich hingeschaut hat. Beispiele für starke Paarungen:
- Fliesenfarbe/-finish + Armaturenfarbe ("die kupferfarbenen Fliesen mit den mattschwarzen Armaturen")
- Rahmenfarbe + Öffnungsart ("anthrazit Dreh-Kipp, tief in die Laibung gesetzt")
- Dielenfarbe + Abschlussprofil ("WPC-Dielen mit dem Alu-Abschluss")
- Thermenmarke + ein konkretes Anschlussdetail ("Vaillant Therme, Magnetfilter sauber am Rücklauf")
Wenn eine Paarung offensichtlich ist, bau die Zeile so: [Hauptelement], [zweites passendes Element] [Verb]. Insgesamt unter 16 Wörter.
- Größe / Umfang des Projekts (wie viele Module, wie große Fläche, große Anlage)
- Ein Detail, das nur ein Profi bemerkt (Lüftung, Entwässerungsgefälle, Unterkonstruktion, Verblechungsdetail, Firstlinie, Dachrinnengefälle, Kaminanschluss, Dielenabstand, sichtbare Dämmlage)

FENSTER- / TÜREN-FOTOS — schau gezielt auf:
- Rahmenfarbe: anthrazit, dunkelbraun, golden oak, weiß — wie sie zur Fassade passt
- Öffnungsart: Dreh-Kipp, Schiebe-/Hebe-Schiebe, Faltschiebe — sofort fachlich
- Gefälle der Fensterbank: gibt es ein sichtbares Gefälle zum Ablauf (eine flache Bank ist ein Fehler)
- Einbautiefe: tief gesetzter Rahmen vs. bündig — bemerkt nur ein Profi
- Anschluss Rahmen zur Wand: verklebt oder nur verputzt — verklebt ist das Premium-Detail
- Bei einer Haustür: Schwelle trifft auf den Bodenbelag, Dichtung drückt sauber an
- Bei einem Dachfenster (Velux/Fakro): Verblechung eingelegt, Ziegel sauber rundherum
- Bei integriertem Rollladen: ist der Kasten in der Wand verbaut oder steht er vor

MARKENNAMEN NACH NISCHE — wenn du sie siehst, verwende sie:
- Klima / Split / Wärmepumpe: Daikin, Mitsubishi, LG, Samsung, Panasonic, Fujitsu, Midea, Gree, Toshiba
- Therme / Heizung: Vaillant, Viessmann, Buderus, Junkers, Wolf, Bosch
- Solar: LONGi, JA Solar, Jinko, Canadian Solar, Trina, SolarEdge (Wechselrichter), SMA, Fronius
- Dach: Bramac, Tondach, Velux, Fakro, EPDM, Blech
- Fenster / Türen: Internorm, Schüco, Rehau, Veka, Aluk
- Pflaster / Einfahrt: Semmelrock, Friedl, Granit, Naturstein
- Terrasse: WPC, Hartholz, Lärche, Bambus

FACHLICHE BEISPIELE — so klingt es, wenn jemand das Handwerk kennt:
Klima / Split (variiere den Fokus — das Außengerät, das Innengerät oder die Kassette, die gedämmte Leitung, die Wandhalterung, der Kabelkanal, der Kondensatablauf, das Gitter, der Standort — NICHT immer "Leitungsführung". Nur was wirklich sichtbar ist):
"Die Daikin-Außeneinheit in Graz ist sauber geworden, Leitung ordentlich gedämmt."
"Das Mitsubishi-Split in Linz sitzt top, Innengerät schön gerade an der Wand."
"Die Deckenkassette in Wien passt perfekt, Gitter bündig in der Decke."
"Das Außengerät in Salzburg ist sauber, Kabelkanal eng an der Wand geklemmt."
"Das Multisplit in Innsbruck schaut ordentlich aus, zwei Innengeräte sauber aufeinander."
"Der Kondensatablauf in Klagenfurt passt, sauber vom Gerät weggeführt."
Therme / Heizung (variiere den Fokus — Abgasführung, Kondensatleitung, Gasleitung, Absperrhähne, Magnetfilter, Thermenmodell & Position, Wandmontage, Speicher, Verteiler — NICHT immer "Leitung an der Wand". Nur was wirklich sichtbar ist):
"Die Vaillant in Wien schaut ordentlich aus, Abgasführung sauber durch die Wand."
"Der Thermentausch in Graz ist sauber geworden, Absperrhähne schön ausgerichtet."
"Der Fußbodenheizungs-Verteiler in Linz passt, saubere Schleifen an den Anschlüssen."
"Die Viessmann in Salzburg sitzt top, Magnetfilter sauber am Rücklauf."
"Die Therme in St. Pölten ist sauber, Kondensatleitung eng an der Wand."
"Der Speicher-Einbau in Villach ist top geworden, Pumpengruppe sauber gesetzt."
"Die Gasleitung in Wels ist ordentlich, schnurgerade an der Wand geklemmt."
Solar:
"Die LONGi-Anlage in Wien schaut top aus, gleichmäßige Neigung über alle Module."
"Das Solardach in Graz wirkt ordentlich, saubere Symmetrie im Layout."
Dach:
"Das Bramac-Ziegeldach in Linz schaut top aus, gleichmäßige Schichtung über alles."
"Das Blechdach in Salzburg ist sauber geworden, klare Ortganglinie."
"Das EPDM-Flachdach in Wien schaut ordentlich aus, sauberer Randabschluss."
"Die Verblechung in Innsbruck passt, sauberes Detail rund um den Kamin."
Pflaster / Einfahrt:
"Das Pflaster in Graz schaut top aus, gleichmäßige Fugen quer durch."
"Die Granit-Einfahrt in Klagenfurt ist sauber geworden, echt sauberer Abschluss."
Fliesen / Bad (variiere den Fokus — Waschbecken, Sanitärobjekte, Duschwand, Nische, Wannenverkleidung, Fliesenmuster, Armaturen — statt immer Fugen. Diese zeigen NUR TON und AUFBAU: nenn nur ein Detail, das du WIRKLICH SIEHST, übernimm diese Beispiele nie, wenn sie nicht im Foto sind):
"Die begehbare Dusche in Wien ist sauber geworden, Fliesennische schön gerade."
"Das Bad in Graz passt, sauberer Übergang vom Boden zur Wand."
"Die freistehende Wanne in Linz sitzt top, saubere Verkleidung rundherum."
"Die Metro-Fliesen in Salzburg sind ordentlich, klare Verbandlinien."
"Die Nasszelle in Villach ist top geworden, Gefälle zum Ablauf passt genau."
"Der Waschtisch in Wels schaut leiwand aus, Becken und Armatur sauber ausgerichtet."
"Der Fischgrätboden in Wien passt, enge Fugen quer durch."
Fenster / Türen:
"Das Schüco-Fenster in anthrazit in Wien schaut top aus, tief gesetzter Rahmen."
"Das Internorm Dreh-Kipp in Graz schaut ordentlich aus, sauberes Gefälle an der Bank."
"Die Faltschiebetür in Linz ist sauber geworden, schmales Rahmenprofil durchgehend."
"Die Haustür in Salzburg passt, sauberer Übergang zur Schwelle."
"Das Velux in Innsbruck ist sauber, Verblechung ordentlich eingelegt."
Garten / Terrasse:
"Die WPC-Terrasse in Wien schaut top aus, gleichmäßiger Dielenabstand."
"Der Garten in Graz wirkt ordentlich, saubere Linie bei den Stützmauern."
Fassade / Verputz:
"Die Fassade in Salzburg schaut top aus, gleichmäßige Deckung über alles."
"Der Verputz in Wien ist ordentlich, saubere Kantenprofile."
Maler:
"Die Fassade in Linz ist sauber geworden, gleichmäßige Deckung quer durch."
"Der Innenanstrich in Graz schaut ordentlich aus, saubere Kanten rund um die Stöcke."

TON: Locker, österreichisch, handwerklich. Persönlichkeit. Wie eine Nachricht von jemandem, der gute Arbeit wirklich schätzt. Verwende Wörter wie: sauber, top, passt, leiwand, ordentlich, sitzt, sauber geworden, schaut top aus.

WANN DEN FALLBACK VERWENDEN (verpflichtend, keine Ausnahmen):
- Du siehst keine fertige Handwerksarbeit im Foto
- Du siehst nur Text, ein Logo, ein Dokument, ein Zertifikat oder ein KI-Bild
- Das Foto ist ein beschrifteter Lieferwagen / Fahrzeug, ein Schild / Banner, eine Visitenkarte oder ein Team- / Porträtfoto — nichts davon ist eine fertige Arbeit. Beschreib NICHT den Lieferwagen oder die Beschriftung ("das silberne Auto schaut top aus" ist falsch). → Fallback
- Du kannst den Ort nicht erkennen
- Du KANNST kein konkretes Material, keine Marke und kein Fachdetail benennen — der Ort allein reicht NICHT
- Allgemeine Wörter wie "Tapete", "Anstrich", "Fliesen", "Dach" ALLEINE reichen nicht — sie brauchen eine Farbe, Technik, Marke oder ein konkretes Detail dazu. Wenn nichts da ist → Fallback.
Fallback: "Bin auf eure Seite in [Ort] gestoßen." — nur Ort oder Stadt.

STRIKTE REGELN:
- Gib NUR die erste Zeile zurück. Keinen anderen Text.
- Kein Emoji, kein Rufzeichen, keine Gedankenstriche.
- Kopier die Beispiele nicht Wort für Wort.
- Maximal 10-16 Wörter.
- JEDER Satz beginnt mit dem Artikel der Arbeit (Die / Der / Das [Sache]) — die einzige Ausnahme ist der Fallback.
- VERBOTENER Schlagzeilen-Stil: "Schöne Arbeit beim...", "Tolle Arbeit beim...", "Schönes Finish beim..." — verboten.
- VERBOTENE passive Endungen: "schön fertiggestellt", "gut ausgeführt", "sauber verlegt", "schön gemacht" — verboten. Beschreib einfach in einem konkreten zweiten Teil, was du wirklich siehst.
- VARIIERE DAS VERB: nimm NICHT immer "schaut top aus" — wechsle natürlich zwischen "ist sauber geworden", "schaut ordentlich aus", "passt", "sitzt top", "ist top geworden", "schaut leiwand aus", "passt perfekt". Dieselbe Wendung bei jeder Arbeit liest sich wie eine Vorlage und killt es.
- BENENNE DIE KONKRETE SACHE, die du siehst, und nimm das, was in DIESEM Foto wirklich heraussticht. Lass kein Detail zum Standard über eine ganze Reihe werden — benenne den tatsächlichen Blickfang DIESES Bildes.
  - Bad-Menü (breit): die Wanne oder freistehende Wanne, die Dusche oder Duschwand, das WC, der Boden, das Fliesenmuster (Metro, Fischgrät, Verband), der Waschtisch, die Nische, der Handtuchheizkörper, der Spiegel, das Becken, die Armaturen. ÜBERSTRAPAZIERT, vermeide als Standard: "Fugen", "Becken / Armatur".
  - Therme / Heizung-Menü (breit): die Abgasführung, die Kondensatleitung, die Gasleitung, die Absperrhähne, der Magnetfilter, das Thermenmodell & wo sie sitzt, die Wandmontage, der Speicher, der Verteiler. ÜBERSTRAPAZIERT, vermeide als Standard: "Leitung an der Wand", "saubere Leitung".
  - Klima / HVAC-Menü (breit): das Außengerät, das Innengerät oder die Kassette, die gedämmte Leitung, die Wandhalterung, der Kabelkanal, der Kondensatablauf, das Gitter, der Standort, die Marke. ÜBERSTRAPAZIERT, vermeide als Standard: "saubere Leitungsführung", "Daikin-Gerät ... Leitungsführung".
- ABER beschreib nur, was wirklich sichtbar ist — Genauigkeit schlägt Abwechslung. Wenn du klar nur schlichte Fliesen und Fugen siehst, sag das ehrlich, oder nimm den Fallback. ERFINDE NIE ein Objekt, Muster, eine Marke oder Farbe, nur um konkret zu klingen oder dich nicht zu wiederholen.
- VERBOTENE allgemeine Endungen: "schaut gut gemacht aus", "schaut solide aus", "schaut nach viel Arbeit aus", "schöne Arbeit", "gutes Finish", "Fugen", "enge Fugen", "gerade Fugen", "saubere Fugen", "Anschlüsse an der Wand", "Leitung an der Wand", "an der Wand", "zur Wand", "über die Wand", "zum System", "saubere Anschlüsse", "klare Anschlüsse", "saubere Linien" — verboten. Das Verb braucht IMMER ein konkret benanntes Element (Marke, Bauteil, Anschluss, Material) — keine Richtung und keine allgemeine Fläche.
- VERPFLICHTEND: jede erste Zeile braucht einen mit KOMMA getrennten zweiten Teil mit einer konkreten Beobachtung. z.B. "...ist sauber geworden, sauberes Detail rund um den Rahmen." — der Satz endet nie nur auf dem Ort oder dem Namen des Objekts. Der zweite Teil muss ein PHYSISCHES Detail der Arbeit sein (eine Passung, eine Führung, eine Ebene, ein Finish, eine Fuge) — NIE ein Ort oder eine Wendung wie "[Stadt] Montage" / "[Stadt] Auftrag". Wenn der Ort alles ist, was du hast, ist das ein Fallback, kein zweiter Teil.
- Kein "Dach-Auftrag" — nenn den Typ: Flachdach, Steildach, Blechdach, Foliendach, Neueindeckung.
- Keine Straßennamen, Hausnummern oder Postleitzahlen — nur Ort oder Stadt.
- Kein "im Foto" oder "im Beitrag". Kein "anscheinend" — bezieh dich nie auf die Bildunterschrift. Keine Kommentare zu ihrem Feed oder der Beitragszahl.
- SCHREIB NIE den Firmen- / Betriebsnamen in die Zeile. Beschreib die Arbeit, die Marke oder den Ort — nicht den Firmennamen. FALSCH: "Die Müller Gas-Service Heizkörper...", "Die Huber Heizungsbau Montage...". RICHTIG: "Der Heizkörper schaut ordentlich aus, saubere Anschlüsse.", "Der Thermentausch ist sauber geworden, Speicher schön gesetzt." Wenn du keine Marke und keinen Ort hast, beschreib die Arbeit schlicht oder nimm den Fallback.
- Keine erfundenen Angaben — beschreib nur, was du WIRKLICH im Foto SIEHST. Diese Regel ÜBERSTEUERT die "variiere es"- und "sei konkret"-Regeln oben: wenn konkret sein Raten bedeuten würde, lass es — beschreib nur, was wirklich da ist, oder nimm den Fallback. Eine schlichte, korrekte Zeile schlägt immer eine konkret erfundene.
- Bei Fenstern schreib NIE "Plastikfenster" — schreib "Kunststofffenster" oder den Markennamen.
- Wenn du einen Markennamen im Foto siehst, MUSST du ihn einbauen — das ist nicht optional.
- MARKE-NISCHE-MATCH: Klima- / Kühlmarken (Daikin, Mitsubishi, Fujitsu, Midea, Gree, Toshiba) nur für Klimabetriebe. Thermen- / Heizungsmarken (Vaillant, Viessmann, Buderus, Junkers, Wolf) nur für Heizungsbetriebe. Nicht mischen.
- SCHREIB NIE "schönes Finish" — benenne IMMER GENAU, was gut ist: die Fugenbreite, die Passung, das Gefälle, die Leitungsführung, der Dielenabstand, die Firstlinie usw.
- Wenn eine Farbe oder Materialfarbe sichtbar ist (anthrazit, graphitgrau, natur, weiß, braun), nenn sie auch."""

# ── DM body copy — one per trade/sub-trade, no em dashes ─────────────────────

DM_BODY_ROOFING     = "hör von ein paar Dachdeckern, dass sie Aufträge verlieren, nur weil wer anderer schneller zurückgeschrieben hat. kennst du das?"
DM_BODY_DRIVEWAY    = "hör von ein paar Pflasterern, dass Anfragen reinkommen und dann nach dem Angebot Funkstille ist. bei dir auch?"
DM_BODY_PATIO       = "hör von ein paar Terrassenbauern, dass Anfragen reinkommen und dann nach dem Angebot nichts mehr kommt. bei dir auch?"
DM_BODY_LANDSCAPERS = "hör von ein paar Gärtnern, dass sie Anfragen kriegen und dann nach dem Angebot nichts mehr. bei dir auch?"
DM_BODY_BUILDERS    = "hör von ein paar Baumeistern, dass die Hälfte der Anfragen nur einen Richtpreis will und dann verschwindet. bei dir auch?"
DM_BODY_RENOVATION  = "hör von ein paar Sanierern, dass sie einen Auftrag durchrechnen und dann einfach nichts mehr hören. bei dir auch?"
DM_BODY_QUOTES      = "hör von ein paar Betrieben, dass Anfragen reinkommen und dann nach dem Angebot Funkstille ist. bei dir auch?"

# ── Follow-up sequences — 3 touches per trade ────────────────────────────────
# Touch 1 (Day 3):  new angle / peer curiosity — sounds like genuine interest
# Touch 2 (Day 10): social proof — concrete result from a similar business
# Touch 3 (Day 21): break-up / loss aversion — walk away, highest reply rate

FOLLOWUPS = {
    "roofers": [
        "Hab letzte Woche mit einem Dachdecker in Graz geredet, der gesagt hat, er verliert die Hälfte seiner Anfragen, nur weil er nicht am selben Tag zurückschreiben kann. Kommt das bei dir oft vor?",
        "Hab einem Dachdecker in Linz geholfen, von verpassten Anrufen zu 3-4 zusätzlichen Aufträgen pro Woche zu kommen, einfach weil seine Anfragen erledigt wurden, während er am Dach war. Dachte, das passt vielleicht.",
        "Klingt, als wär gerade nicht der richtige Zeitpunkt. Kein Problem, ich lass es dabei. Falls du mal reden willst, wie du mehr aus deinen Anfragen rausholst, weißt du, wo du mich findest.",
    ],
    "driveways": [
        "Hab mit einem Pflasterer in Salzburg geredet, der gemeint hat, er verliert 2 Aufträge pro Woche an Leute, die einfach billiger anbieten. Hast du das Gefühl, dass das Preisvergleichen schlimmer geworden ist?",
        "Hab einem Pflasterbetrieb in Wien geholfen, aufzuhören, toten Angeboten nachzulaufen, und stattdessen die zu schließen, die zählen. 6 zusätzliche Aufträge im ersten Monat. Vielleicht einen Blick wert.",
        "Klingt, als passt der Zeitpunkt nicht. Ich lass es dabei. Falls es mal relevant wird, schreib mir einfach.",
    ],
    "patio": [
        "Hab mit einem Terrassenbauer in Innsbruck geredet, der gesagt hat, er verbringt die Hälfte seiner Zeit mit Angeboten, aus denen nichts wird. Hast du viel davon bei dir?",
        "Hab einem Gartenbetrieb in Graz geholfen, die Zeitfresser auszusortieren und nur mehr Zeit in ernste Anfragen zu stecken. Ihre Abschlussquote ist von 1 von 5 auf 1 von 2 gestiegen. Dachte, das ist vielleicht nützlich.",
        "Klingt, als wär gerade nicht der richtige Zeitpunkt. Kein Problem, ich lass es dabei und du kannst jederzeit darauf zurückkommen.",
    ],
    "landscapers": [
        "Hab mit einem Gärtner in Klagenfurt geredet, der gemeint hat, seine besten Aufträge kommen immer noch über Mundpropaganda, aber er kriegt es nicht beständig hin. Ist das bei dir genauso?",
        "Hab einem Gartengestalter in der Steiermark geholfen, seine Facebook-Seite in einen beständigen Strom an Anfragen zu verwandeln, von sporadischer Mundpropaganda auf 8-10 Anfragen pro Woche. Einen kurzen Austausch wert?",
        "Klingt, als wär jetzt nicht der Zeitpunkt. Ich lass es dabei. Falls sich das ändert, weißt du, wo du mich findest.",
    ],
    "builders": [
        "Hab letzte Woche mit einem Baumeister in Linz geredet, der gesagt hat, 3 von 5 Anfragen sind Leute, die nur eine grobe Idee wollen und dann verschwinden. Hast du das Gefühl, dass das in letzter Zeit schlimmer geworden ist?",
        "Hab einem Baubetrieb in Wien geholfen, die Zeitfresser auszusortieren und sich auf ernste Anfragen zu konzentrieren. Sie sind von 10 Anfragen pro Woche auf 6 ordentliche gekommen, aus denen wirklich was wird. Vielleicht relevant.",
        "Klingt, als wär gerade nicht der richtige Zeitpunkt. Kein Problem, ich lass es dabei. Meld dich gerne, falls sich was ändert.",
    ],
    "renovation contractors": [
        "Hab mit einem Sanierer in Wien geredet, der gesagt hat, er verbringt Tage damit, Aufträge durchzurechnen, die nach dem Angebot still werden. Kommt das bei dir oft vor?",
        "Hab einem Sanierungsbetrieb in Niederösterreich geholfen, die Zeit vom Angebot bis zum Auftrag zu halbieren, indem sie geändert haben, wie sie nachfassen, kein Verlieren von Aufträgen mehr durch Funkstille. Dachte, das passt vielleicht.",
        "Klingt, als wär jetzt nicht der richtige Zeitpunkt. Ich lass es dabei. Falls du mal darüber reden willst, gerne.",
    ],
    "quotes": [
        "Hab letzte Woche mit einem Betrieb geredet, der gesagt hat, seine Angebote werden in etwa 60% der Fälle kalt, keine Antwort, kein Grund. Hast du das bei dir auch?",
        "Hab einem Handwerksbetrieb in der Steiermark geholfen, vom Nachlaufen hinter kalten Angeboten dahin zu kommen, dass die Kunden ihm nachlaufen, indem sie eine Sache geändert haben, wie sie nachfassen. Teil ich gern, wenn's nützlich ist.",
        "Klingt, als passt der Zeitpunkt nicht. Ich lass es dabei. Du weißt, wo du mich findest, falls sich was ändert.",
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
    # Drop trailing legal suffixes (UK + AT/DE). Run twice so "GmbH & Co KG" fully
    # collapses (first pass strips the "& Co KG" tail, second strips "GmbH").
    for _ in range(2):
        n = re.sub(
            r"\s*(?:&\s*co\.?)?\s*\b(ltd\.?|limited|llp|llc|plc|inc\.?|co\.?|company|"
            r"gmbh|gesmbh|ges\.?\s?m\.?\s?b\.?\s?h\.?|mbh|kg|keg|og|ag|gen|"
            r"e\.?\s?u\.?|e\.?\s?k\.?)\b\.?\s*$",
            "", n, flags=re.I).strip()
    if not re.search(r"[A-Za-z]", n):
        return ""
    words = n.split()
    if len(words) > 6:           # still descriptor-laden — cap so it stays readable
        n = " ".join(words[:4])
    return n.strip()


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
    # ── AT/DE trade + filler words ───────────────────────────────────────────
    "der", "die", "das", "und", "dein", "deine", "ihr", "ihre", "neu", "neue",
    "profi", "meister", "technik", "betrieb", "handwerk", "gmbh", "kg", "og", "ag",
    "bad", "bäder", "bader", "badsanierung", "badezimmer", "fliesen", "fliesenleger",
    "heizung", "gas", "installateur", "sanitär", "sanitar", "haustechnik",
    "dach", "dachdecker", "spengler", "spenglerei", "fenster", "türen", "turen",
    "garten", "gartenbau", "landschaftsbau", "galabau", "maler", "malerei",
    "elektro", "elektriker", "bau", "baumeister", "sanierung", "renovierung",
    "klima", "klimatechnik", "solar", "photovoltaik", "terrasse", "terrassenbau",
    "pflaster", "boden", "bodenleger", "tischler", "tischlerei", "schreiner",
    "küche", "kuche", "küchen", "kuchen", "möbel", "mobel", "fassade", "verputz",
}
_PITCH_PLACES = {
    "glasgow", "edinburgh", "manchester", "london", "leeds", "bristol", "cardiff",
    "liverpool", "birmingham", "sheffield", "newcastle", "worcester", "carlisle",
    "nottingham", "leicester", "coventry", "hull", "preston", "wales", "scotland",
    "yorkshire", "kent", "essex", "surrey", "midlands", "lancashire", "cheshire",
    # ── Austrian cities, states & common Vienna districts ────────────────────
    "wien", "graz", "linz", "salzburg", "innsbruck", "klagenfurt", "villach",
    "wels", "sankt", "pölten", "polten", "dornbirn", "bregenz", "steyr",
    "feldkirch", "leoben", "wolfsberg", "krems", "amstetten",
    "steiermark", "tirol", "kärnten", "karnten", "burgenland", "vorarlberg",
    "oberösterreich", "oberosterreich", "niederösterreich", "niederosterreich",
    "österreich", "osterreich", "favoriten", "donaustadt", "floridsdorf",
}


def _who_for_pitch(company_name: str) -> str:
    """How to address them in 'I've set one up for ___ too'. Most trade names read
    badly there ('for Bathroom Fitting too', 'for New Bathrooms Glasgow too'), so
    default to 'you' and only use the real name when it's a short, genuinely
    distinctive brand ('Ava Bathrooms', 'Downs Brothers')."""
    clean = _clean_company_name(company_name)
    if not clean:
        return "dich"
    # Domain / handle ("7bathrooms.uk", "acme.co.uk") reads like a URL, not a name.
    if re.search(r"\.(uk|com|co|net|org|io|biz)\b", clean.lower()) or "www." in clean.lower() or "@" in clean:
        return "dich"
    words = [w for w in clean.split() if w != "&"]
    if len(words) > 2:                      # long names sound like a descriptor here
        return "dich"
    distinctive = [w for w in words
                   if w.lower().strip(".&") not in _PITCH_GENERIC
                   and w.lower().strip(".&") not in _PITCH_PLACES]
    return clean if distinctive else "dich"  # all-generic/location -> "dich"

def _followup_text(touch: int, company_name: str = "", niche: str = "") -> str:
    """Niche-specialised nudge — references the actual trade/job and the sample,
    in the same voice as the opener, so it reads tailored rather than generic.
    Two wordings per touch for spam-safety."""
    trade_noun, job, examples = _pitch_for(niche)
    who = _clean_company_name(company_name) or "deinen Betrieb"
    t = max(1, min(touch, FOLLOWUP_TOUCHES))
    if t == 1:
        return random.choice([
            f"Nochmal hallo, gar kein Stress. Das Muster, das ich für {who} gebaut hab, gibt jemandem in etwa 30 Sekunden einen groben Preis für {job} und schickt dir dann seine Daten. Schick's dir gern jederzeit rüber, sag einfach Bescheid.",
            f"Hey, gar kein Druck. Das Muster, das ich für {who} gemacht hab, zeigt jemandem in Sekunden einen Richtpreis für {job} und schickt dir seine Infos. Schick's dir gern rüber, wann's dir passt.",
        ])
    if t == 2:
        return random.choice([
            f"Meld mich nur einmal nochmal. Es stellt ein paar kurze Fragen ({examples}), gibt einen Richtpreis und schickt dir nur die ernsten weiter. Soll ich dir das Muster rüberschicken?",
            f"Komm noch ein letztes Mal darauf zurück. Der ganze Sinn ist, dass es die Zeitfresser von den echten trennt, bevor sie bei dir landen, damit du nur die Aufträge anbietest, die's wert sind. Einen kurzen Blick wert?",
        ])
    return random.choice([
        "Alles gut, wenn's nichts für dich ist, ich lass es dabei. Falls du mal einen kurzen Blick aufs Muster werfen willst, meld dich einfach. LG.",
        "Kein Problem, wenn's nicht dein Ding ist, ich lass es dabei. Das Angebot steht, falls du mal schauen willst. LG.",
    ])

VOICE_NOTE_PROMPT = """\
Du schreibst ein kurzes WhatsApp-Sprachnachricht-Skript für jemanden, der einen Handwerksbetrieb (AT/DE) kalt anschreibt. In derselben Nachricht wie diese Sprachnachricht wird ein Screenshot mitgeschickt. Gib NUR das Skript zurück — keine Labels, keine Erklärung, keine Anführungszeichen. Schreib es als natürliche gesprochene Worte, auf Deutsch.

Das ist die ERSTE Nachricht. Der Empfänger sieht den Betriebsnamen des Absenders ohnehin schon auf WhatsApp, also musst du dich nicht vorstellen.

TON: Locker, österreichisch, wie ein echter Mensch. Kein Verkäufer. Verwende natürliche Ausdrücke. Klinge wie ein Kollege, der was Interessantes entdeckt hat, nicht wie eine Agentur, die ein Produkt verkauft.

AUFBAU — 5 Teile, insgesamt unter 90 Wörter (muss in unter 45 Sekunden sprechbar sein):

1. BEGRÜSSUNG + AUFHÄNGER
Beginn mit "Servus [Name]," — finde den Vornamen des Inhabers aus dem Info-Bereich ("Inhaber", "Gründer", "Geschäftsführer"), aus dem Seitennamen, wenn ein Personenname drin ist, oder aus Bildunterschriften. Nimm nie den Firmennamen als Vornamen.
Dann beziehe dich auf EINEN konkreten Beitrag — nenn das tatsächliche Material, die Art der Arbeit und ein sichtbares Detail oder den Ort aus dem Beitrag. So konkret wie möglich.
GUT: "hab grad die Granit-Einfahrt mit dem dunklen Randstein gesehen, die ihr in Graz fertig gemacht habt, schaut sauber aus" / "hab die Terrasse mit den WPC-Dielen und den Stützmauern gesehen, die ihr gepostet habt"
SCHLECHT: "hab gesehen, ihr macht Gartenbau in der Steiermark" / "hab gesehen, ihr arbeitet in ganz Tirol" — das liest die Bio, nicht die Beiträge.
Beende den Aufhänger natürlich: "schaut top aus" oder ähnlich. Halt es kurz.

2. SCREENSHOT-HINWEIS
Weis sie auf den Screenshot hin. Ein Satz.
Beispiel: "Schau dir kurz den Screenshot an, den ich geschickt hab, das ist im Grunde, wie eine Anfrage aussieht, wenn sie auf deinem Handy landet."

3. WAS SIE BEKOMMEN
Erklär beide Seiten in einem Satz — was der Kunde bekommt und was der Inhaber bekommt.
Der Kunde bekommt sofort einen groben Preis. Der Inhaber bekommt alle Daten — Budget, Ort, Art der Arbeit — bevor er zum Hörer greift.

4. WAS ES MACHT
Ein Satz. Der Sinn davon ist, die Zeitfresser auszusortieren und ihnen Ärger zu ersparen.
Verwende natürliche Ausdrücke: "Zeitfresser", "spart dir den Ärger", "nimmt dir das Lästige ab".
Verwende NICHT: Leads, Lösung, System, Software, KI, Agent, skalieren, wachsen, qualifiziert.

5. DISQUALIFIZIEREN + CTA
"Ist vielleicht nichts für dich, aber wenn dich interessiert, wie's funktioniert, schreib einfach zurück und wir reden kurz. LG."

REGELN:
- Insgesamt unter 90 Wörter
- Der Aufhänger muss konkret auf DIESE Seite passen — nicht etwas Allgemeines, das du jedem schicken könntest
- Durchgehend lockere Sprache — "Servus", "passt", "Zeitfresser", "echt", "Ärger" sind alle in Ordnung
- Keine Rufzeichen. Keine Emojis. Keine übertriebenen Komplimente.
- Klinge wie ein echter Mensch, nicht wie ein Skript

GOLDSTANDARD-BEISPIEL — triff genau diesen Ton:
"Servus Haydan. Hab grad die Gartenumgestaltung gesehen, die ihr in der Steiermark gepostet habt, schaut top aus. Schau dir kurz den Screenshot an, den ich geschickt hab, das ist im Grunde, wie eine Anfrage aussieht, wenn sie auf deinem Handy landet. Der Kunde bekommt sofort einen groben Preis, und du kriegst alle Daten, Budget, Ort, alles, bevor du überhaupt zum Hörer greifst. Ich bau die für Handwerker, um die Zeitfresser auszusortieren und dir den Ärger zu ersparen. Ist vielleicht nichts für dich, aber wenn dich interessiert, wie's funktioniert, schreib einfach zurück und wir reden kurz. LG."

Bad examples — never write these:
- "I have an AI solution that generates qualified leads" (jargon, wrong)
- "Would you be open to a quick chat?" (wrong CTA structure)
- "Saw you do landscaping in Manchester..." (no specific detail, too generic)
- "It asks questions and calculates a price..." (too specific, kills curiosity)"""

app = Flask(__name__)

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
                    ['script','style','noscript','svg','iframe']
                        .forEach(t => document.querySelectorAll(t).forEach(el => el.remove()));
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
                # Scroll to where the first post typically lives on a Facebook Page
                target_page.evaluate("window.scrollTo(0, 850)")
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
        if len(loc.split()) <= 4:
            return loc

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

    trade_signals = [
        (["roofing", "roofer", "flat roof", "pitched roof", "epdm", "felt roof",
          "fascia", "soffit", "guttering", "lead flashing", "chimney", "velux"], "roofers"),
        (["drone", "aerial", "uav", "uas", "thermal imaging", "aerial survey",
          "aerial photography", "aerial footage", "roof inspection drone"], "drone operators"),
        (["renovation", "refurb", "kitchen reno", "bathroom reno", "full house",
          "house renovation", "property renovation", "interior renovation"], "renovation contractors"),
        (["builder", "building contractor", "extension", "loft conversion",
          "groundwork", "brickwork", "blockwork", "underpinning", "footings",
          "new build", "steel frame", "construction"], "builders"),
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
        return ("Handwerksbetriebe", "einen Auftrag", "Umfang des Auftrags, was dazugehört")
    has = lambda *words: any(w in n for w in words)
    # Order matters: bathroom first (a "Bathroom Renovation" value also contains
    # "renovation"); air con before heating; etc. Each check matches BOTH the
    # English canonical label (from detect_niche) AND a raw German Excel value.
    if has("bathroom", "wet room", "ensuite", "bad", "nasszelle", "feuchtraum"):
        return ("Badbetriebe", "ein neues Bad", "Komplettsanierung, Nasszelle, Neuverfliesung")
    if has("hvac", "air con", "aircon", "air conditioning", "refrigeration", "cooling",
           "klima", "wärmepump", "warmepump", "kälte", "kalte"):
        return ("Klimatechniker", "eine neue Anlage", "ein Raum, Multisplit oder das ganze Haus")
    if has("boiler", "gas", "heating", "central heat", "therme", "heizung", "kessel"):
        return ("Heizungsbetriebe", "eine neue Therme", "Thermentausch, Anlagen-Upgrade oder nur eine Reparatur")
    if has("plumb", "installateur", "sanitär", "sanitar", "klempner"):
        return ("Installateure", "einen Auftrag", "eine Reparatur, eine Neuinstallation oder ein ganzer Auftrag")
    if has("window", "glazing", "conservatory", "upvc", "fenster", "verglas", "wintergart", "tür", "haustür"):
        return ("Fensterbetriebe", "neue Fenster", "Fenstertausch, Haustür oder Wintergarten")
    if has("roof", "dach", "spengler"):
        return ("Dachdecker", "ein neues Dach", "Neueindeckung, Flachdach oder nur eine Reparatur")
    if has("solar", "photovolt", "pv"):
        return ("Solarbetriebe", "eine PV-Anlage", "Aufdach, Speicher oder die ganze Anlage")
    if has("paving", "driveway", "block pav", "resin", "pflaster", "einfahrt", "asphalt"):
        return ("Pflasterer", "eine neue Einfahrt", "Einfahrt, Hof oder Gehweg")
    if has("tiling", "tiler", "fliesen"):
        return ("Fliesenleger", "einen Fliesenauftrag", "Bad, Boden oder die ganze Fläche")
    if has("render", "ewi", "k rend", "fassade", "vollwärme", "vollwarme", "wärmedämm", "warmedamm", "wdvs"):
        return ("Fassadenbetriebe", "eine Fassade", "Vollwärmeschutz, Verputz oder Sanierung")
    if has("landscap", "garden", "garten", "landschaft", "galabau"):
        return ("Gartenbetriebe", "einen Gartenauftrag", "Gestaltung, Pflasterung oder die ganze Anlage")
    if has("decking", "terrass"):
        return ("Terrassenbauer", "eine neue Terrasse", "Holz, WPC oder die ganze Terrasse")
    if has("joinery", "joiner", "carpenter", "kitchen", "tischler", "schreiner", "möbel", "mobel", "küche", "kuche"):
        return ("Tischler", "einen Auftrag", "Küche, Einbaumöbel oder Innenausbau")
    if has("electric", "elektr"):
        return ("Elektrobetriebe", "einen Auftrag", "Neuverkabelung, Verteiler oder eine Installation")
    if has("bricklay", "brickwork", "maurer", "mauerwerk"):
        return ("Maurer", "einen Auftrag", "Rohbau, Mauerwerk oder eine Reparatur")
    if has("plaster", "verputz", "stuckateur", "trockenbau"):
        return ("Verputzer", "einen Auftrag", "Innenputz, Fassade oder Trockenbau")
    if has("paint", "decorat", "maler", "anstrich"):
        return ("Maler", "einen Auftrag", "Innen, Fassade oder die ganze Wohnung")
    if has("floor", "boden", "parkett", "laminat", "vinyl"):
        return ("Bodenleger", "einen Auftrag", "Parkett, Laminat oder Vinyl")
    if has("pool", "schwimmbad", "schwimmbecken"):
        return ("Poolbauer", "einen Pool", "Neubau, Sanierung oder die ganze Anlage")
    if has("garage door", "garagentor", "rolltor", "markise", "beschattung", "sonnenschutz"):
        return ("Beschattungsbetriebe", "einen Auftrag", "Garagentor, Markise oder Pergola")
    if has("renovation", "refurb", "general build", "builder", "extension", "remodel",
           "sanierung", "renovierung", "baumeister", "umbau", "neubau", "zubau", "generalunternehm"):
        return ("Sanierungsbetriebe", "eine Sanierung", "wie groß der Auftrag ist, was dazugehört")
    # any other recognised trade — neutral German default
    return ("Handwerksbetriebe", "einen Auftrag", "Umfang des Auftrags, was dazugehört")


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
        f"Ich bin Student und studier KI, ich bau Angebots-Chatbots für {trade_noun}. {sp}, "
        f"und für {who} hab ich auch schon eins eingerichtet, gratis.",
        f"Ich bin Student und steig grad in KI ein, ich bau Angebots-Bots für {trade_noun}. {sp}, "
        f"und für {who} hab ich auch schon eins zusammengestellt, gratis.",
        f"Ich bin Student und lern KI, ich mach Angebots-Chatbots für {trade_noun}. {sp}, "
        f"und für {who} hab ich auch schon eins gemacht, kostenlos.",
    ])
    # Disarm: your real differentiator an agency can't fake. Answers "why free?".
    disarm = random.choice([
        "Kein Haken, ich bin Student, keine Agentur, ich brauch die echten Beispiele gerade mehr als das Geld.",
        "Kein Haken, ich bin Student, keine Agentur, ich brauch die Beispiele aus der Praxis gerade mehr als alles andere.",
    ])
    # Sell the BENEFIT, not the mechanics. Chill, spoken-out-loud tone (contractions,
    # "basically", "gonna"): what's in it for them, not how the bot works.
    # ONE self-contained middle (not a how×benefit cross-product) so the two
    # sentences never echo each other ("every site visit ... every site visit",
    # "wasted trip ... wasted trips"). Lead-QUALITY angle: they already have leads.
    middle = random.choice([
        "Es geht nicht um mehr Anfragen, die hast du eh schon. Es sorgt nur dafür, dass die, die bei dir landen, deine Zeit auch wert sind. Weniger Hin und Her, ernste Aufträge gebucht.",
        f"Die Leute kriegen vorab einen Richtpreis für {job}, also sind die, die dir schreiben, schon ernst, Budget und Details schon dabei. Jeder Vor-Ort-Termin wirklich der Mühe wert.",
        "Es trennt die guten Anfragen von den Zeitfressern, bevor sie bei dir landen, also ist jeder Vor-Ort-Termin ein echter Auftrag, keine vergebene Fahrt.",
        f"Sie kriegen zuerst einen groben Richtpreis für {job}, also schreiben dir nur die, die's ernst meinen und loslegen wollen. Gute Anfragen, keine Zeitfresser.",
    ])
    # Take-away close: removing the pressure lifts replies. Easy channel choice.
    close = random.choice([
        "Ist vielleicht gar nicht dein Ding, aber wenn du schauen willst, schick ich's dir rüber, hier oder WhatsApp?",
        "Ist vielleicht nichts für dich, aber wenn du schauen willst, schick ich's dir rüber, hier oder WhatsApp?",
    ])
    pitch = f"{intro} {disarm}\n\n{middle}\n\n{close}\n\nLG"

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
    who = _clean_company_name(company_name) or "deinen Betrieb"
    # Short niche word for a relevant variant: "ein neues Bad" -> "Bad".
    _, job, _ex = _pitch_for(niche)
    thing = re.sub(r"^(ein neues |eine neue |einen neuen |neue |neuen |ein |eine |einen |einer )",
                   "", (job or "")).strip() or "Auftrag"
    if first:
        f = first.lower()   # all-lowercase reads like a colleague, not marketing
        return random.choice([
            f"{f}, kurze sache",          # casual, 3 words, no hype
            f"hab dir was gebaut, {f}",   # honest curiosity, ties to the email's hook
            f"für {f}",                   # 2-word pattern interrupt
            f"{f}, einen blick wert?",    # question framing (~+10% opens)
            f"{f}, deine {thing}-angebote",  # niche-relevant, still curiosity not pitch
        ])
    return random.choice([
        f"hab dir was gebaut",
        f"kurze sache für {who}",
        f"{who}, einen blick wert?",
        f"deine {thing}-angebote",
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

    company = _clean_fb_title(page_title) or "deinen Betrieb"
    niche_resolved = niche or _en_detect_niche(td, company)
    _, job, _examples = _pitch_for(niche_resolved)   # e.g. "a new bathroom"

    first = _email_first_name(owner)
    greet = f"Hallo {first}," if first else "Hallo,"

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
        ["Ich bin Student und bau Sofort-Angebots-Tools für Handwerksbetriebe.",
         "Ich will dir nichts andrehen, ein paar Betriebe nutzen schon eins, ich brauch nur noch eine solide Referenz für meine Arbeit."],
        ["Ich bin Student und mach Sofort-Angebots-Tools für Handwerksbetriebe.",
         "Ich will nichts verkaufen, ein paar Betriebe nutzen schon eins, ich brauch nur noch eine gute Referenz für meine Arbeit."],
    ])

    # The "what it is" — ONE short line that names the mechanism concretely (an
    # estimate on THEIR site), niche-aware via {job}. No second clause. "Estimate"
    # not "price" — it's an honest ballpark range, and over-promising a firm
    # "price" then showing a range reads worse.
    what_it_is = random.choice([
        f"Es setzt einen Sofort-Richtpreis für {job} direkt auf deine Seite.",
        f"Es gibt den Leuten einen Sofort-Richtpreis für {job} auf deiner Seite.",
        f"Es rechnet {job} für die Leute direkt auf deiner Seite durch.",
    ])

    # The payoff — a short filter line, then the punchline broken across two lines
    # for a dramatic beat (reads as one sentence, lands harder). The build-up line
    # above ("a price on your site") earns the punchline so it isn't vague.
    payoff = random.choice([
        ["Die Zeitfresser fallen weg.",
         "Jeder Vor-Ort-Termin",
         "wird zum Auftrag."],
        ["Kein Durchrechnen von Aufträgen mehr, aus denen nie was wird.",
         "Jeder Vor-Ort-Termin",
         "wird zum Auftrag."],
        ["Die Zeitfresser sortieren sich von selbst aus.",
         "Jedes Angebot",
         "wird zum Auftrag."],
    ])

    # Believable proof (a real thing exists, on their branding) + the ask, with a
    # channel choice. Two short lines.
    close = random.choice([
        ["Ich hab schon eins mit deinem Branding gebaut.",
         "Willst du's hier oder über WhatsApp?"],
        ["Hab schon eins mit deinem Branding gebaut.",
         "Soll ich's dir hier oder über WhatsApp schicken?"],
    ])

    blocks = [greet]
    if op:
        blocks.append(op)
    blocks += frame
    blocks.append(what_it_is)
    blocks += payoff
    blocks += close
    blocks.append(f"LG,\n{EMAIL_SIGNOFF}")
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
    greet = f"Servus {name_hint}, " if name_hint else "Servus, "
    who = _clean_company_name(company_name) or "deinen Betrieb"
    if name_hint and who.lower() == name_hint.lower():
        who = "dich"   # sole trader: name == company, avoid "Servus Pierce ... für Pierce"

    # Rotate natural wordings of the middle so no two messages are textually
    # identical (kills Facebook's duplicate-content spam flag) while keeping the
    # exact meaning and tone. The "I'm a student" anchor and the "Can I send it
    # over?" closer never rotate out.
    build = random.choice([
        f"ich bin Student und bau KI-Angebots-Bots für {trade_noun}",
        f"ich bin Student und steig grad in KI ein, bau Angebots-Bots für {trade_noun}",
        f"ich bin Student und lern KI, ich bau Angebots-Bots für {trade_noun}",
    ])
    sample = random.choice([
        f"und für {who} hab ich schon ein Muster zusammengestellt",
        f"und für {who} hab ich schon ein Muster gebaut",
        f"und für {who} hab ich schon ein Muster gemacht",
    ])
    show = random.choice([
        "Kein Druck, ich dachte, ich zeig dir einfach, was es macht, statt es zu erklären.",
        "Kein Druck, ich find's einfacher, es dir zu zeigen, als es zu erklären.",
        "Kein Druck, dachte, es ist einfacher, dir einfach zu zeigen, was es macht.",
    ])
    return f"{greet}{build} {sample}. {show} Kann ich's dir rüberschicken?"


def _clean_dm_location(loc: str) -> str:
    """Tidy a scraped location so it reads like a human typed it mid-sentence.
        'Wien & Umgebung'              -> 'Wien'
        'Graz und Umgebung'            -> 'Graz'
        'Großraum Linz'                -> 'Linz'
        'Salzburg, Hallein & Umgebung' -> 'Salzburg'
    A real person names ONE town. Returns '' if nothing usable is left."""
    n = (loc or "").strip()
    if not n:
        return ""
    # Cut at the first connector — keep only the lead place.
    for sep in (" & ", " und ", ",", "/", " - "):
        i = n.lower().find(sep.lower())
        if i > 0:
            n = n[:i]
            break
    # Drop scraped filler words.
    n = re.sub(r"\b(umgebung|umkreis|grossraum|großraum|raum|der|die|das|region|bezirk|ganz|umliegende?n?)\b", "", n, flags=re.I)
    n = re.sub(r"\s{2,}", " ", n).strip(" .,-&")
    return n


def pick_british_message(opener: str, text_data: str, name_hint: str = "", company_name: str = "") -> str:
    """Casual British variant — short, punchy, no student framing.
    Structure: [personalised opener] + niche line + timewaster hook + sample close.
    No vision call needed beyond the opener (which is passed in already generated).
    """
    excel_name = _td_section(text_data, "COMPANY NAME")
    if excel_name:
        company_name = excel_name
    excel_niche = _td_section(text_data, "TRADE TYPE")
    niche = excel_niche or _en_detect_niche(text_data, company_name)
    trade_noun, _job, _ex = _pitch_for(niche)

    # ALWAYS use the real town in the niche line. The personalised opener, the
    # location and the niche always stay specific to this lead — "deiner Gegend"
    # is only a last-resort fallback when no location could be found at all.
    # The opener/page scan (extract_prospect_location) is UK-place based, so for
    # AT/DE leads it usually finds nothing — fall back to the Excel city column
    # (--- LISTED CITY ---) so the real Austrian town still lands in the line.
    prospect_loc = (_clean_dm_location(extract_prospect_location(opener, text_data))
                    or _clean_dm_location(_td_section(text_data, "LISTED CITY")))
    loc = prospect_loc or "deiner Gegend"
    # Rotate the surface wording so 150 messages aren't byte-identical (Facebook's
    # duplicate-content spam flag, the same reason pick_body/pick_stripped rotate).
    # Meaning, voice and the "estimator chatbots ... timewasters" anchor stay fixed;
    # only the phrasing moves. The user's locked wording is the first option.
    niche_line = random.choice([
        f"bau Angebots-Chatbots für {trade_noun} in {loc}, nimmt dir alle Zeitfresser ab.",
        f"ich bau Angebots-Chatbots für {trade_noun} in {loc}, hält dir die Zeitfresser vom Hals.",
        f"bau Angebots-Chatbots für {trade_noun} in {loc}, sortiert die Zeitfresser aus, bevor sie bei dir landen.",
    ])
    close = random.choice([
        "hab schon ein Muster mit deinem Branding zusammengestellt. soll ich's dir rüberschicken, hier oder whatsapp?",
        "hab schon ein Muster mit deinem Branding zusammengestellt. soll ich's dir schicken, hier oder whatsapp?",
        "hab schon ein Muster mit deinem Branding gemacht. soll ich's dir rüberschicken, hier oder whatsapp?",
    ])

    op = (opener or "").strip().rstrip(".")
    # If the vision model couldn't find a job photo to react to, it returns the
    # generic "came across your page in [town]" fallback. That reads like every
    # other cold DM. Swap it for an HONEST opener that names the elephant in the
    # room — owners can smell a mail-merge, so admitting it disarms them. Keeps
    # the greeting prefix the model already prepended.
    low = op.lower()
    for marker in ("bin auf eure seite", "auf eure seite"):
        i = low.find(marker)
        if i != -1:
            prefix = op[:i]  # keep the "servus, " greeting
            honest = random.choice([
                "ich weiß, du kriegst wahrscheinlich einen Haufen von denen, also halt ich's kurz",
                "ich weiß, du kriegst massig solche Nachrichten, also gleich auf den Punkt",
                "ich bin ehrlich, ich weiß, du wirst mit denen zugeschüttet, also kurz",
            ])
            op = (prefix + honest).strip()
            break
    return f"{op}\n\n{niche_line}\n\n{close}"


EXTRACT_PROMPT = """\
Schau dir den Seitentext dieses Handwerksbetriebs genau an. Liste NUR Details auf, die du wirklich lesen kannst. Rate nicht.

Gib eine kurze Stichpunktliste mit NUR diesen, falls vorhanden, aus:
- Gewerk (Gärtner / Dachdecker / Baumeister / Sanierer / Installateur usw.)
- NEUESTER BEITRAG: Art der Arbeit, Material, Farbe/Finish, jeder Ort in der Bildunterschrift (voller Ortsname, keine Abkürzungen oder Postleitzahlen)
- Andere Orte in Beiträgen: alle Orts-/Stadtnamen in Bildunterschriften (nur volle Namen)
- Profil-Ort: der Ort aus dem Info-/Bio-Bereich (für die Nachricht)
- Vorname des Inhabers (NUR wenn im Firmennamen, in der E-Mail oder im Info-Bereich — nicht geraten)
- Alles Ungewöhnliche: riesiger Auftrag, besonderes Material, auffälliges Vorher/Nachher, ungewöhnliche Leistung

Nimm nichts auf, was nicht klar im Text steht. Nur Stichpunktliste, nichts anderes."""


_FALLBACK_TEMPLATES = {
    "roofers": [
        "Bin auf eure Seite in {loc} gestoßen, einiges an Dacharbeit dabei.",
        "Hab eure Seite in {loc} entdeckt, schaut aus, als hättet ihr beim Dach viel zu tun.",
    ],
    "landscapers": [
        "Bin auf eure Seite in {loc} gestoßen, einiges an Gartenarbeit dabei.",
        "Hab eure Seite in {loc} entdeckt, schaut aus, als hättet ihr im Garten viel zu tun.",
    ],
    "builders": [
        "Bin auf eure Seite in {loc} gestoßen, einiges an Bauarbeit dabei.",
        "Hab eure Seite in {loc} entdeckt, solide Bauarbeit dabei.",
    ],
    "renovation contractors": [
        "Bin auf eure Seite in {loc} gestoßen, ordentliche Sanierungsarbeit dabei.",
    ],
    "drone operators": [
        "Bin auf eure Seite in {loc} gestoßen, ordentliche Luftaufnahmen dabei.",
    ],
}
_FALLBACK_LOC_ONLY = [
    "Bin auf eure Seite in {loc} gestoßen, solide Arbeit dabei.",
    "Hab eure Seite in {loc} entdeckt, schaut aus, als hättet ihr viel zu tun.",
]
_FALLBACK_NOTHING = "Bin auf eure Seite gestoßen, solide Arbeit dabei."


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


def patch_fallback_opener(opener: str, text_data: str, city_hint: str = "") -> str:
    """
    If the model used the generic fallback, normalise it to the plain
    'Came across your page in [location].' form (mirror of the HU fallback).
    Runs AFTER call_gemini() — Python safety net.
    """
    # Only touch it if it's the generic fallback
    if "auf eure seite" not in opener.lower() and "eure seite entdeckt" not in opener.lower():
        return opener

    loc = (_td_section(text_data, "LISTED CITY")
           or extract_prospect_location("", text_data)
           or city_hint)
    result = f"Bin auf eure Seite in {loc} gestoßen." if loc else "Bin auf eure Seite gestoßen."
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
    if has("bad", "nasszelle", "feuchtraum", "dusche"):   return "bathroom renovation"
    if has("klima", "wärmepump", "warmepump", "kälte", "kalte", "split"):
        return "air conditioning"
    if has("therme", "heizung", "gas", "kessel"):         return "heating"
    if has("installateur", "sanitär", "sanitar", "klempner"): return "plumbing"
    if has("fenster", "verglas", "wintergart", "haustür", "schiebetür"): return "windows and doors"
    if has("dach", "spengler"):                           return "roofing"
    if has("solar", "photovolt", "pv"):                   return "solar"
    if has("pflaster", "einfahrt", "asphalt"):            return "paving"
    if has("fliesen"):                                    return "tiling"
    if has("fassade", "vollwärme", "vollwarme", "wärmedämm", "warmedamm", "wdvs"):
        return "external wall insulation"
    if has("garten", "landschaft", "galabau"):            return "landscaping"
    if has("terrass"):                                    return "decking"
    if has("tischler", "schreiner", "möbel", "mobel", "küche", "kuche"): return "joinery"
    if has("elektr"):                                     return "electrical"
    if has("sanierung", "renovierung", "baumeister", "umbau", "neubau", "zubau", "generalunternehm"):
        return "general building"
    return n


def _resolve_niche(text_data: str) -> str:
    """
    Resolve the niche the SAME way the body does: the Excel TRADE TYPE column is
    authoritative (so the opener and the pitch never disagree), then name-first,
    then full text-scoring. Result is a canonical key.
    """
    excel_niche = _td_section(text_data, "TRADE TYPE")
    if excel_niche:
        return _canon_niche(excel_niche)
    company_name = _td_section(text_data, "COMPANY NAME")
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


def _en_opener_postprocess(opener: str, text_data: str) -> str:
    """Reject opener if it contradicts the detected niche OR embeds the company
    name; fall back to the city line."""
    niche = _resolve_niche(text_data)
    company = _td_section(text_data, "COMPANY NAME")
    reason = ""
    if not _en_opener_matches_niche(opener, niche):
        reason = f"contradicts niche '{niche}'"
    elif _opener_repeats_company_name(opener, company):
        reason = "embeds the company name"
    if reason:
        city = _td_section(text_data, "LISTED CITY")
        fallback = (f"Bin auf eure Seite in {city} gestoßen."
                    if city else "Bin auf eure Seite gestoßen.")
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
    "schaut top aus", "ist top geworden", "passt perfekt", "sitzt top",
    "schaut ordentlich aus", "schaut leiwand aus", "passt", "schaut richtig gut aus",
    "ist gut geworden", "sitzt ordentlich", "ist sauber geworden", "schaut sauber aus",
]
# Descriptors for the SECOND clause, so it isn't always "sauber ___". These must NOT
# overlap with _OPENER_LEADS words (top/ordentlich/passt) or the line doubles up.
# Kept to physical-observation words.
_OPENER_DETAILS = [
    "gleichmäßig", "schnurgerade", "eben", "ausgerichtet", "bündig",
    "eng", "knackig", "präzise", "dicht", "glatt", "exakt",
]


def call_gemini(text_data: str = "", images: list = None, prompt: str = None) -> str:
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
        f"\n\nFÜR DIESE EINE: verwende '{_lead_phrase}' als Lob-Wendung und keine andere. "
        f"Stütz den zweiten Teil auf '{_detail_word}' statt auf 'sauber' (vermeide im Detail "
        f"das Wort 'sauber' und wiederhol die Lob-Wendung nicht im Detail). "
        f"Bau es so: Die [die tatsächliche Arbeit] {_lead_phrase}, {_detail_word} [Detail].")

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
    for model, content in write_attempts:
        for attempt in range(3):
            try:
                import time
                result = _run_groq(client, model, content, max_tokens=400, temperature=0.92)
                if re.search(r"\b\d[\d,]*\s*(followers?|likes?|reviews?|ratings?|fans?)", result, re.I):
                    print(f"Output contained forbidden metric, retrying...")
                    continue
                if CUSTOM_OPENER_POSTPROCESS:
                    result = CUSTOM_OPENER_POSTPROCESS(result, text_data) or result
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
@app.route("/")
def index():
    return render_template("dm_de.html", default_excel=DEFAULT_EXCEL, variant=MESSAGE_VARIANT)

@app.route("/toggle_variant", methods=["POST"])
def toggle_variant():
    """Cycle the message variant: full -> stripped -> british -> full."""
    global MESSAGE_VARIANT
    cycle = {"full": "stripped", "stripped": "british", "british": "full"}
    MESSAGE_VARIANT = cycle.get(MESSAGE_VARIANT, "full")
    _save_variant(MESSAGE_VARIANT)
    print(f"Message variant -> {MESSAGE_VARIANT}")
    return jsonify({"variant": MESSAGE_VARIANT})

@app.route("/generate", methods=["POST"])
def generate():
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

        # Capture from Chrome (primary method — only if no manual image)
        if not images:
            hint_url = request.form.get("url", "").strip()
            captured_images, captured_text = capture_profile_tab(hint_url)
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
                    "Kein Facebook-/Instagram-Profil-Tab in Chrome gefunden. "
                    "Öffne zuerst die Profilseite in Chrome, dann klick auf Generieren."
                )
                if not text_data.strip():
                    text_data = "--- KEINE SEITENDATEN --- Erfinde KEINE Details."

        if not images and not text_data.strip() and not notes:
            return jsonify({
                "error": "Öffne zuerst das Profil in Chrome, dann klick auf Generieren. "
                         "Oder füg einen Screenshot ein / tipp Profil-Infos ein."
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
                    raw = call_gemini(text_data=text_data, images=images if images else None)
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
                    raw = call_gemini(text_data=text_data, images=images if images else None)
                except Exception:
                    break
                key = re.sub(r"\s+", " ", raw.strip().lower())
                if key in seen:
                    continue
                seen.add(key)
                dms.append(_assemble(raw))
            dm = dms[0]
        clipboard = copy_to_clipboard(dm)
        return jsonify({"dm": dm, "dms": dms, "warning": warning, "clipboard": clipboard})

    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500

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
                    "Kein Facebook-/Instagram-Profil-Tab in Chrome gefunden. "
                    "Öffne zuerst die Profilseite in Chrome, dann klick auf Generieren."
                )
                if not text_data.strip():
                    text_data = "--- KEINE SEITENDATEN --- Erfinde KEINE Details."

        if not images and not notes:
            return jsonify({
                "error": "Öffne zuerst das Profil in Chrome, dann klick auf Generieren. "
                         "Oder füg einen Screenshot ein / tipp Profil-Infos ein."
            }), 400

        script = call_gemini(text_data=text_data, images=images if images else None, prompt=VOICE_NOTE_PROMPT)
        return jsonify({"script": script, "warning": warning})

    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500

# ── Auto Sender state ──────────────────────────────────────────────────────────
import openpyxl as _xl
from pathlib import Path as _Path

_auto_state = {
    "excel_path":    None,
    "current_row":   None,
    "sent_today":    0,
    "wb":            None,
    "ws":            None,
    "cols":          {},
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
    def _find_or_create(match, label):
        """Return the index of the first header matching `match(header)`,
        creating a new column titled `label` at the end if none exists."""
        idx = next((i for i, h in enumerate(headers) if match(h)), None)
        if idx is None:
            ws.cell(row=1, column=len(headers) + 1, value=label)
            headers.append(label.lower())
            idx = len(headers) - 1
            created.append(label)
        return idx

    # Follow-up tracking columns are created on first load if absent.
    # Match the SENT-tracking column only — NOT the lead-contact column. On the
    # newer list "Contact (Email or Personal FB)" also contains "contact", so the
    # old loose "contact" match hijacked the email column and overwrote it with
    # "Skip"/"Yes". Match "contacted"/"outreach status" instead.
    contacted_idx = _find_or_create(
        lambda h: "contacted" in h or "outreach" in h or h.strip() == "status",
        "Contacted?")
    sent_date_idx = _find_or_create(lambda h: "date" in h, "Sent Date")
    replied_idx   = _find_or_create(lambda h: "repl" in h, "Replied?")
    followups_idx = _find_or_create(lambda h: "follow" in h, "Follow-Ups")
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
        "Kann nicht in die Excel-Datei schreiben — sie ist in Excel geöffnet (oder "
        "synchronisiert noch in OneDrive). Schließ die Tabelle in Excel, warte, bis "
        "OneDrive fertig ist, dann klick nochmal."
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
            return jsonify({"error": f"Excel-Datei nicht gefunden: {path}"}), 400

        wb, ws, cols = _load_excel(path)
        _auto_state.update({"excel_path": path, "wb": wb, "ws": ws, "cols": cols})

        lead = _next_lead(ws, cols)
        if not lead:
            return jsonify({"error": "Keine weiteren Leads — alles erledigt!"}), 200

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
            return jsonify({"error": "Lad zuerst deine Excel-Datei."}), 400

        # If we still have a current_row that was never marked sent/skip,
        # mark it as Skip now so it never resurfaces as a duplicate.
        prev_row = _auto_state["current_row"]
        if prev_row is not None and cols["contacted"] is not None:
            ws.cell(row=prev_row, column=cols["contacted"] + 1, value="Skip")
            _save_wb()

        lead = _next_lead(ws, cols, skip_row=prev_row)
        if not lead:
            return jsonify({"error": "Keine weiteren Leads — alles erledigt!"}), 200

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
            return jsonify({"error": "Chrome nicht gefunden. Schließ alle Chrome-Fenster, doppelklick auf START DE.bat, warte, bis Chrome ganz offen ist, dann probier's nochmal."}), 500

        with sync_playwright() as p:
            browser = p.chromium.connect_over_cdp(f"http://localhost:{CDP_PORT}")
            ctx     = browser.contexts[0]

            # Close any old Facebook profile tabs (not Messenger, not the app)
            for _pg in list(ctx.pages):
                _u = _pg.url.lower()
                if ("facebook.com/" in _u and
                        "messenger.com" not in _u and
                        "/messages" not in _u and
                        "localhost" not in _u):
                    try:
                        _pg.close()
                        print(f"  Closed old tab: {_u[:70]}")
                    except Exception:
                        pass

            page = ctx.new_page()
            import time as _perf
            _t_load = _perf.time()
            try:
                page.goto(lead["url"], wait_until="domcontentloaded", timeout=45000)
            except Exception:
                try:
                    page.goto(lead["url"], wait_until="commit", timeout=30000)
                except Exception as e:
                    return jsonify({"error": f"Seite konnte nicht geladen werden: {e}"}), 500
            print(f"  [timing] page load: {_perf.time() - _t_load:.1f}s")
            page.wait_for_timeout(500)   # settle after load (image-render wait below protects the screenshot)

            # Strip whatever overlay is dimming the page (a dialog or a full-screen
            # scrim). It dims the screenshot the vision model reads and can sit
            # invisibly on top of the Message button. Also logs every full-viewport
            # fixed/absolute layer to server.log so we can target it if it persists.
            try:
                page.keyboard.press("Escape")
            except Exception:
                pass
            try:
                page.evaluate(r"""() => {
                    // 1) Persistent CSS so the login/cookie dialog stays hidden even
                    //    when React re-adds it (the one-time removal loses that race).
                    if (!document.getElementById('dm-overlay-killer')) {
                        const st = document.createElement('style');
                        st.id = 'dm-overlay-killer';
                        st.textContent =
                            '[role="dialog"]{display:none!important;}' +
                            'html,body{overflow:auto!important;}';
                        (document.head || document.documentElement).appendChild(st);
                    }
                    // 2) Strip full-viewport dark/semi-transparent backdrop scrims, and
                    //    keep doing so for a few seconds as Facebook re-renders.
                    const nuke = () => {
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
                    };
                    nuke();
                    const obs = new MutationObserver(nuke);
                    obs.observe(document.documentElement, { childList: true, subtree: true });
                    setTimeout(() => obs.disconnect(), 6000);
                }""")
            except Exception:
                pass

            # Extract text
            try:
                text = page.evaluate("""() => {
                    ['script','style','noscript','svg','iframe']
                        .forEach(t => document.querySelectorAll(t).forEach(el => el.remove()));
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

            if MESSAGE_VARIANT == "stripped":
                # Direct one-line variant — no opener, so skip the screenshot AND
                # the vision call entirely (faster, and the photo isn't used).
                dm = _dm_case(tidy_message(pick_stripped_message(
                    text_data, name_hint,
                    lead.get("name", "") or extract_company_name(text_data))))
            else:
                # Screenshot the first-post area so the vision model can see the actual job photos
                images = []
                try:
                    import time as _tw
                    page.evaluate("window.scrollTo(0, 850)")
                    _tw.sleep(0.8)   # wait for images to render
                    shot = page.screenshot(type="jpeg", quality=72)
                    images = [(shot, "image/jpeg")]
                    print(f"  First-post screenshot: {len(shot)} bytes")
                    page.evaluate("window.scrollTo(0, 0)")
                except Exception as _se:
                    print(f"  Screenshot skipped: {_se}")

                # Generate DM — pass screenshot so vision model sees the actual job
                import time as _perf2
                _t_ai = _perf2.time()
                opener = call_gemini(text_data=text_data, images=images if images else None)
                print(f"  [timing] AI opener: {_perf2.time() - _t_ai:.1f}s")
                opener = patch_fallback_opener(opener, text_data, lead.get("city", ""))
                opener = _apply_greeting(opener, name_hint)
                if MESSAGE_VARIANT == "british":
                    company_name = lead.get("name", "") or extract_company_name(text_data)
                    dm = _dm_case(tidy_message(pick_british_message(
                        opener, text_data, name_hint, company_name)))
                else:
                    opener = opener.rstrip(".!,") + "."
                    trade  = detect_trade(text_data)
                    body   = pick_body(trade, text_data, lead.get("city", ""), lead.get("name", "") or extract_company_name(text_data))
                    dm     = _dm_case(tidy_message((opener + " " + body).strip()))

            # Copy DM to clipboard — always reliable
            clipboard = copy_to_clipboard(dm)

            # Click Message button to open the chat — user just Ctrl+V to paste
            msg_opened = _click_message_btn(page)
            if msg_opened:
                print("  Message chat opened. User can Ctrl+V to paste.")
            else:
                print("  Could not click Message button — user can open chat manually.")

        return jsonify({
            "dm":         dm,
            "lead":       lead,
            "clipboard":  clipboard,
            "msg_opened": msg_opened,
            "sent_today": _auto_state["sent_today"],
        })

    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/auto_skip", methods=["POST"])
def auto_skip():
    """Skip current lead (mark as Skip) and return next lead info without opening browser."""
    try:
        wb   = _auto_state["wb"]
        ws   = _auto_state["ws"]
        cols = _auto_state["cols"]
        row  = _auto_state["current_row"]

        if ws is None:
            return jsonify({"error": "Keine Excel-Datei geladen."}), 400

        if row is not None and cols["contacted"] is not None:
            ws.cell(row=row, column=cols["contacted"] + 1, value="Skip")
            _save_wb()

        _auto_state["current_row"] = None

        lead = _next_lead(ws, cols)
        if not lead:
            return jsonify({"error": "Keine weiteren Leads — alles erledigt!"}), 200

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
            return jsonify({"ok": False, "error": "Nichts zum Markieren — klick zuerst auf „Profil öffnen & DM generieren“."}), 400

        # Without a 'Contacted?' column we can't record the send, so _next_lead would
        # keep returning this same lead forever. Surface that clearly instead of
        # silently looping on the same company.
        if cols["contacted"] is None:
            return jsonify({"ok": False, "error": "Keine „Contacted?“-Spalte in der Excel-Datei — der Lead kann nicht als gesendet markiert werden, deshalb kommt immer derselbe. Schließ die Datei in Excel, lad die Leads neu, dann probier's nochmal."}), 400
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
            return jsonify({"error": f"Excel-Datei nicht gefunden: {path}"}), 400
        wb, ws, cols = _load_excel(path)
        _email_state.update({"excel_path": path, "wb": wb, "ws": ws, "cols": cols, "current_row": None})
        lead = _next_email_lead(ws, cols)
        if not lead:
            return jsonify({"error": "Keine E-Mail-Leads in dieser Datei gefunden."}), 200
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
            return jsonify({"error": "Lad zuerst deine Excel-Datei."}), 400

        prev_row = _email_state["current_row"]
        if prev_row is not None and cols["contacted"] is not None:
            ws.cell(row=prev_row, column=cols["contacted"] + 1, value="Skip")
            _save_email_wb()

        lead = _next_email_lead(ws, cols, skip_row=prev_row)
        if not lead:
            return jsonify({"error": "Keine weiteren E-Mail-Leads — alles erledigt!"}), 200
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
            return jsonify({"error": "Chrome nicht gefunden. Doppelklick auf START DE.bat, warte, bis Chrome offen ist, dann probier's nochmal."}), 500

        with sync_playwright() as p:
            browser = p.chromium.connect_over_cdp(f"http://localhost:{CDP_PORT}")
            ctx     = browser.contexts[0]
            for _pg in list(ctx.pages):
                _u = _pg.url.lower()
                if ("facebook.com/" in _u and "messenger.com" not in _u and
                        "/messages" not in _u and "localhost" not in _u):
                    try: _pg.close()
                    except Exception: pass
            page = ctx.new_page()
            try:
                page.goto(scan_url, wait_until="domcontentloaded", timeout=45000)
            except Exception:
                try:
                    page.goto(scan_url, wait_until="commit", timeout=30000)
                except Exception as e:
                    return jsonify({"error": f"Seite konnte nicht geladen werden: {e}"}), 500
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
                    ['script','style','noscript','svg','iframe'].forEach(t => document.querySelectorAll(t).forEach(el => el.remove()));
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
                page.evaluate("window.scrollTo(0, 850)")
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
            return jsonify({"error": "Keine Excel-Datei geladen."}), 400
        if row is not None and cols["contacted"] is not None:
            ws.cell(row=row, column=cols["contacted"] + 1, value="Skip"); _save_email_wb()
        _email_state["current_row"] = None
        lead = _next_email_lead(ws, cols)
        if not lead:
            return jsonify({"error": "Keine weiteren E-Mail-Leads — alles erledigt!"}), 200
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
            return jsonify({"error": "Lad zuerst deine Leads im Auto-Sender-Tab."}), 400

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
            return jsonify({"error": "Lad zuerst deine Leads."}), 400
        r = int(request.get_json(force=True).get("row", 0))
        if r < 2:
            return jsonify({"error": "Ungültige Zeile."}), 400

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

@app.route("/mark_replied", methods=["POST"])
def mark_replied():
    """Mark a lead as replied so it drops off the follow-up list for good."""
    try:
        ws   = _auto_state["ws"]
        cols = _auto_state["cols"]
        if ws is None:
            return jsonify({"error": "Lad zuerst deine Leads."}), 400
        r = int(request.get_json(force=True).get("row", 0))
        if r < 2 or cols.get("replied") is None:
            return jsonify({"error": "Diese Zeile kann nicht markiert werden."}), 400
        ws.cell(row=r, column=cols["replied"] + 1, value="Yes")
        _save_wb()
        return jsonify({"ok": True})
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(debug=False, port=5003)

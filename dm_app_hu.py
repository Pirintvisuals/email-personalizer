"""
Hungarian mode - imports the full English DM app and overrides
PROMPT + body copy with Hungarian versions. Runs on port 5002.
Do NOT modify dm_app.py - this is a standalone wrapper.
"""

# START_HU.bat runs this with stdout/stderr redirected to server_hu.log
# ("python dm_app_hu.py >> server_hu.log 2>&1"), and a redirected (non-TTY)
# stream on Windows defaults to the cp1252 codepage, not UTF-8. This codebase
# prints plenty of Hungarian accents and "→" arrows in log lines (e.g. the
# opener-rejection log below), and any one of those raised a
# UnicodeEncodeError that got caught by call_gemini's retry loop and treated
# as "the model failed" — burning every remaining retry attempt and returning
# a 500 for what was actually a successful generation, just an unprintable
# log line. Reconfigure both streams to UTF-8 process-wide, before anything
# else prints, so no log message can ever crash a request again. Also forces
# line buffering so log lines show up immediately instead of sitting in an
# unflushed buffer during a long-running dev server process.
import sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
sys.stderr.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)

import os
os.environ.setdefault("CDP_PORT", "9223")
os.environ.setdefault("DEFAULT_EXCEL", r"C:\Users\pirin\Downloads\hungary_leads_3niches 08 25.xlsx")

import random
import re as _re
import dm_app
from hu_niche import (
    _hu_detect_niche,
    _hu_detect_niche_from_name,
    _opener_matches_niche,
    _NICHE_VOCAB,
    _TEXT_KEYWORDS,
    _norm,
    NICHE_OPENER_INSTRUCTIONS,
)
# SECOND round of this exact failure (2026-09-17) — see the ORIGINAL note
# below this one for the first round (llama-4-scout/maverick dying in
# 2026-07). This time ALL THREE models this pipeline depended on died at
# once: qwen/qwen3.6-27b, llama-3.3-70b-versatile, AND llama-3.1-8b-instant
# all 404 as "model does not exist". Confirmed live in server_email_hu.log —
# every single opener call for a whole batch fell through every model to the
# last-resort keyword-based fallback (_hu_text_opener's "Látom, hogy Önnél is
# szerepel a/az X." templated line), which is why dozens of kertépítő leads
# in a row read as near-identical ("...szerepel az öntözőrendszer." repeated
# verbatim): the REAL personalization step was silently failing on every
# lead, not producing duplicate leads or fabricated content, just falling
# back to the same generic template over and over. It also explains the
# slowness: every lead was burning through 3 dead-model attempts (each with
# its own retry/backoff) before ever reaching that fallback.
#
# Queried this account's live model list directly (client.models.list())
# rather than guessing: qwen/qwen3.6-27b was bumped to qwen/qwen3.8-27b (a
# straight version bump, same vision capability, confirmed working with
# reasoning_effort="none" exactly like its predecessor). llama-3.3-70b-
# versatile and llama-3.1-8b-instant have no direct successor on this
# account; openai/gpt-oss-120b/20b are the closest live general-purpose text
# models and are verified working below.
#
# ORIGINAL note (2026-07-21): llama-4-maverick-17b-128e-instruct was shut
# down 2026-03-09, and llama-4-scout-17b-16e-instruct followed on 2026-07-17
# — four days before that fix. Confirmed 27/27 in server_hu.log: every
# single call 404s with "model does not exist or you do not have access to
# it". That meant the vision step had been silently a no-op on every lead
# since the 17th — every opener fell straight through to the TEXT-ONLY
# models, guessing from scraped page text with no actual look at the
# screenshot (this is why openers sometimes described things not in the
# photo at all, e.g. a bedroom "compliment" for a general contractor). It
# also funneled ALL traffic across every language variant sharing this Groq
# key onto the two remaining TEXT_MODELS, which is why llama-3.3-70b-
# versatile started hitting 429s and burning 20-60s+ per lead in
# retry/backoff — confirmed in the same log.
dm_app.VISION_MODELS = ["qwen/qwen3.8-27b"]

# TEXT_MODELS lives on dm_app.py itself (shared with every language variant
# in THEORY, but each variant is its own process, so this reassignment only
# touches the HU process — dm_app_de.py/dm_app.py's own EN run need this
# same fix applied there separately if they still point at the dead llama
# models). Verified live with this account's key before wiring in.
dm_app.TEXT_MODELS = ["openai/gpt-oss-120b", "openai/gpt-oss-20b"]

# qwen/qwen3.6-27b is a reasoning model that defaults to emitting a <think>...
# block ahead of its real answer — verified live: without reasoning_effort,
# raw output starts "<think>\nHere's a thinking process...". Passing
# reasoning_effort="none" gives a clean direct answer. But passing that same
# kwarg to the existing TEXT_MODELS (llama-3.3-70b-versatile,
# llama-3.1-8b-instant) is a hard 400 ("reasoning_effort is not supported with
# this model") — verified live — so it can only be added for qwen, not
# dm_app.py's shared _run_groq signature itself, hence the wrap rather than a
# plain reassignment.
_dm_run_groq = dm_app._run_groq

# qwen/qwen3.6-27b's own latency is highly variable (measured in server_hu.log
# from 1.5s up to 66s on this account, no correlation with image size — likely
# provider-side load, since Groq just migrated everyone off llama-4-scout onto
# this model 4 days before this fix). Uncapped, one slow vision call could
# block an entire lead for over a minute. A 12s per-attempt timeout raises
# groq.APITimeoutError (confirmed a plain Exception subclass, so it's caught
# by call_gemini's existing `except Exception` in dm_app.py without needing
# any change there) — the retry loop then tries once more (2s sleep + another
# 12s) before falling through to the fast, reliable TEXT_MODELS, capping the
# worst case at roughly 26s instead of 66s+, in exchange for occasionally
# trading a slow-but-real vision opener for a faster text-only guess.
_QWEN_TIMEOUT_S = 12


def _hu_run_groq(client, model, content, max_tokens=500, temperature=0.7):
    if "qwen" in model:
        resp = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": content}],
            max_tokens=max_tokens,
            temperature=temperature,
            reasoning_effort="none",
            timeout=_QWEN_TIMEOUT_S,
        )
        return resp.choices[0].message.content.strip()
    if "gpt-oss" in model:
        # Unlike qwen, gpt-oss has no "none" reasoning_effort — verified
        # live, only low/medium/high are accepted, and it hard-400s on
        # "none". It ALWAYS spends some real completion tokens on a
        # reasoning field before the visible answer (confirmed: a 20-token
        # budget produced zero visible content, all 20 spent on reasoning).
        # "low" keeps that overhead small, and this call site's max_tokens
        # (500 default) already leaves enough headroom for reasoning +
        # a one-sentence opener to both fit — verified live.
        resp = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": content}],
            max_tokens=max_tokens,
            temperature=temperature,
            reasoning_effort="low",
        )
        return resp.choices[0].message.content.strip()
    return _dm_run_groq(client, model, content, max_tokens, temperature)


dm_app._run_groq = _hu_run_groq

# LOWERCASE_DM lowercases the first letter of every sentence (an English
# "casual DM" convention) via _dm_case()/_casual_lowercase() in dm_app.py,
# applied to every language's output at send-time. It's why HU messages came
# out "szia. az a ..." instead of "Szia. Az a ...". Off for HU — proper
# capitalisation reads as more professional, which was the whole point of
# this change.
dm_app.LOWERCASE_DM = False

dm_app.GRAMMAR_CHECK = False
dm_app.STRICT_NAME_DETECTION = True
# MESSAGE_VARIANT is persisted in the shared settings.json (also read by
# dm_app.py/dm_app_de.py), so a "british" toggle left on from the English
# app would silently route HU sends through pick_british_message (English,
# not niche-Hungarian) instead of _hu_pick_body. Pin it here without saving,
# so the English app's toggle is untouched.
dm_app.MESSAGE_VARIANT = "full"
dm_app.ALWAYS_GREET = True
# Magázás/önözés throughout (was "Szia, {name}. " / "Szia. ").
# "Kedves {name}!" rather than "Tisztelt X Úr!" because formal Hungarian
# address takes the SURNAME, and _hu_extract_first_name() deliberately returns
# the GIVEN name (it takes the last token, Hungarian name order being family
# name first) — there is no reliable surname or gender to build "Tisztelt
# Kovács Úr" from, and getting that wrong is worse than not using it.
# "Kedves + given name" + önözés is the standard Hungarian business register
# and is exactly what hu_email_gen.py:447 already emits, so the DM and email
# channels now greet identically.
# Both still end in "!" so _hu_apply_greeting's keep_cap check leaves the
# opener capitalised.
dm_app.GREETING_NAME_FMT = "Kedves {name}! "
dm_app.GREETING_SOLO_FMT = "Jó napot kívánok! "   # no name fallback

# In "full" mode dm_app pastes _auto_state["last_screenshot"] — a shot of the
# prospect's OWN Facebook page — into the message box. Sending someone a
# picture of their own page reads as "I scraped you" and is the most obviously
# automated part of the message; the opener already proves a human looked at
# the work. (The UK "funny" variant is different: there the joke photo IS the
# message.) No-op the paste so HU sends stay text-only. The "Copy picture"
# button in the UI still works if a screenshot is ever wanted manually.
# Note: dm_app then logs "Could not attach picture automatically" — that's the
# no-op reporting False, not a real failure.
HU_ATTACH_SCREENSHOT = False
if not HU_ATTACH_SCREENSHOT:
    dm_app._paste_image_into_box = lambda page, image_bytes: False

# dm_app.py runs detect_brand_color() on every single lead's profile
# screenshot (dm_app.py:3501), unconditionally, before pick_body is even
# called. Its result (brand_color) is only ever read by pick_british_message
# (dm_app.py:1756) — the EN "british" variant — never by _hu_pick_body below,
# so on HU it's a second qwen/qwen3.6-27b vision call per lead for a value
# nothing uses. Confirmed in server_hu.log: this account's Groq TPM budget
# (8000/min) was getting exhausted by opener-writing + this call combined,
# causing 429 rate-limit errors that turned a ~1s opener into 19-38s per lead
# (retry loop's 2s sleep-and-retry on every 429). No-opping this for HU cuts
# vision-model traffic per lead roughly in half for zero loss of message
# content.
dm_app.detect_brand_color = lambda *a, **k: ""

# dm_app._apply_greeting() hardcodes English "Hey"/"Hi" and never reads
# GREETING_NAME_FMT/GREETING_SOLO_FMT (dead constants in dm_app.py itself) —
# so without this override every HU send started with "Hey"/"Hi" instead of
# "Szia". Reimplement it here using the two format constants above.
def _hu_apply_greeting(opener: str, name_hint: str, market: str = "US") -> str:
    # `market` is unused — it only picks US vs UK wording in dm_app.py, and the
    # HU greeting is always the magázás one above — but auto_next() passes it
    # positionally, so the signature has to match or every /auto_next 500s.
    g = dm_app.GREETING_NAME_FMT.format(name=name_hint) if name_hint else dm_app.GREETING_SOLO_FMT
    keep_cap = g.rstrip()[-1:] in ".!?"
    return g + (opener if keep_cap else opener[:1].lower() + opener[1:])
dm_app._apply_greeting = _hu_apply_greeting

# Common Hungarian given names — male and female
_HU_GIVEN_NAMES = {
    "Gábor","László","István","Péter","Zoltán","Attila","Tibor","Ferenc","Sándor",
    "János","Gyula","Ernő","Béla","Imre","Ákos","Norbert","Tamás","Balázs","Zsolt",
    "Roland","Richárd","Dávid","András","Márton","Bence","Máté","Ádám","Dániel",
    "Levente","Krisztián","Róbert","Gergő","Gergely","Csaba","Szabolcs","Lajos",
    "Dezső","Árpád","Mihály","Miklós","Boldizsár","Benedek","Barnabás","Botond",
    "Dominik","Erik","Endre","Frigyes","Géza","Győző","Hunor","Jenő","Kálmán",
    "Károly","Kornél","Kristóf","Lóránd","Luca","Marcell","Milán","Nándor",
    "Olivér","Ottó","Patrik","Pál","Rezső","Simon","Szilveszter","Tivadar",
    "Valter","Viktor","Vilmos","Zalán",
    # Female
    "Erzsébet","Katalin","Mária","Anna","Éva","Zsuzsanna","Ildikó","Judit",
    "Márta","Ágnes","Edit","Erika","Gabriella","Henrietta","Irén","Klára",
    "Krisztina","Magdolna","Monika","Nikolett","Noémi","Orsolya","Réka",
    "Szilvia","Tímea","Veronika","Viktória","Virág","Zsuzsa","Zsófia",
    "Boglárka","Csilla","Emőke","Fanni","Hajnalka","Kinga","Nóra","Petra",
    "Piroska","Renáta","Sarolta","Tünde","Beáta","Anita","Barbara","Dorina",
}

_HU_NAV_WORDS = {
    "facebook", "bejelentkezés", "regisztráció", "hirdetés", "hírek", "marketplace",
    "watch", "groups", "gaming", "home", "messenger", "notifications", "menu",
    "like", "follow", "share", "comment", "more", "search", "create",
    "főoldal", "videók", "fotók", "névjegy", "közösség", "vélemények", "események",
    "kedvelés", "követés", "megosztás", "keresés", "értesítések", "posts", "reels",
}

def _hu_extract_name(text: str) -> str:
    """Find a Hungarian given name only in the FB page title area (first ~8 lines)."""
    if not text:
        return ""
    lines = [l.strip() for l in text.strip().split("\n") if l.strip()]
    for line in lines[:8]:
        if line.lower() in _HU_NAV_WORDS:
            continue
        if len(line) > 50:          # page titles are short; longer = content/post
            continue
        if line == line.upper():    # ALL-CAPS = heading, label, abbreviation
            continue
        if line[0].isdigit():       # starts with digit = phone, address, stat
            continue
        for word in _re.split(r"[\s\-–—|·•,/()[\]]+", line):
            w = word.strip(".,!?:")
            if w in _HU_GIVEN_NAMES:
                return w
    return ""

dm_app.CUSTOM_NAME_EXTRACTOR = _hu_extract_name

# -- Hungarian opener prompt --------------------------------------------------
dm_app.PROMPT = """\
Te egy hideg Facebook DM nyitósorát írod. CSAK a nyitósort add vissza - semmi mást. Semmi magyarázat, semmi emoji, semmi extra szöveg.

Látsz egy képernyőképet egy magyar szakipari Facebook oldalról. Nézd meg alaposan a fotót. Reagálj egy befejezett munkára - úgy mintha havernak írnál, aki ért is hozzá.

ELSŐBBSÉG - SZAKMAI RÉSZLET:
Ha látod bármelyiket, reagálj arra - ezek sokkal meggyőzőbbek mint az általános "szép lett":
- Márkanév (lásd lista lentebb)
- Anyagfajta (cserép, pala, EPDM, porcelán, mészkő, gránit, WPC, kompozit, resin stb.)
- Elhelyezés / kivitelezési minőség (fugaköz, dőlésszög, csőelvezetés, szegélykő, mintaillesztés)
- Méret / projekt nagysága (hány panel, mekkora terület, nagy rendszer)
- Részlet amit csak szakember vesz észre (szellőzés, vízvezetés lejtése, tartószerkezet, szigetelés, gerinccsempe illesztés, oromszegély vonala, ereszcsatorna lejtése, kéménykörüli villásszeg, korcok távolsága, alátétfólia látható rétege)

ABLAK / NYÍLÁSZÁRÓ FOTÓNÁL — ezeket keresd konkrétan:
- Profilszín: antracit, sötétbarna, aranytölgy, fehér — hogyan illik a homlokzathoz
- Nyitástípus: bukó-nyíló, toló/emelt-toló (lift-slide), harmonika szárny — ez rögtön szakmai
- Külső párkány dőlésszöge: látható-e a lejtés a vízvezetéshez (lapos párkány = hiba)
- Tokreveal mélysége: mélyen beépített tok vs. lapos beépítés — csak szakember veszi észre
- Fal-tok csatlakozás: szalaggal-e vagy csak vakolva — a szalagos csatlakozás a prémium
- Bejárati ajtónál: küszöb illesztése a padlóburkolattal, tömítőcsík préselése
- Tetőtéri ablaknál (Velux/Fakro): villogópajzs lefektetése, cserép visszavezetése a keret körül
- Integrált redőnynél: a redőnytok beépítve van-e a falba vagy kiszögell

MÁRKANEVEK NICHES SZERINT - ha látod, használd:
- Klíma / split / hőszivattyú: Daikin, Mitsubishi, LG, Samsung, Panasonic, KAISAI, Midea, Gree, Toshiba
- Kazán / gázfűtés: Vaillant, Viessmann, Bosch, Nibe, Stiebel Eltron, Ariston, Baxi, Ferroli, Buderus, Junkers
- Napelem: LONGi, Jinko, Canadian Solar, Risen, Trina, Huawei (inverter), SMA, Fronius
- Tető: Bramac, Tondach, Eternit, Velux, Fakro, Creaton, Lindab, Tegola, Zenit, Icopal
- Ablak / ajtó: Schüco, Rehau, Veka, Internorm, Gealan, KBE
- Térkő / burkolat: Semmelrock, Bradstone, Marshalls, Tobermore
- Terasz / deszka: WPC, bangkirai, teak, kompozit deszka

SZAKMAI PÉLDÁK - így néz ki ha valaki ért hozzá:
Klíma / split:
"Az a Daikin egység Szombathelyen nagyon jól néz ki, szép a csőelvezetés."
"Az a Mitsubishi split Pécsett jó helyre lett bepakolva, rendesen van a kültéri egység."
Kazán / gázfűtés:
"Az a Vaillant kazán Debrecenben nagyon rendesen néz ki, szép a csőelvezetés az aljzaton."
"Az a Viessmann rendszer Győrben jól néz ki, látszik a gondos bekötés."
Napelem:
"Az a LONGi rendszer Győrben nagyon jól néz ki, egyforma a panelek dőlésszöge."
"Az a napelem tető Egerben komolyan néz ki, látszik a szimmetria."
Tetőfedés:
"Az a Bramac cserép Debrecenben nagyon jól néz ki, egyforma a fedési minta."
"Az a korcolt fémtető Győrben klassz lett, szép az oromszegély vonala."
"Az a EPDM lapostető Sopronban rendesen néz ki, precíz a szegélyezés."
"Az a Velux ablak Miskolcon jól néz ki, pontos a keret körüli szigetelés."
"Az a cserepeslemez Egerben nagyon jól néz ki, egyforma a korcok távolsága."
"Az a Lindab ereszcsatorna Kecskeméten rendesen néz ki, egyenletes a lejtés."
"Az a kéményszigetelés Miskolcon nagyon pontosan van megcsinálva, szép a körüli villásszeg."
Térkő / felhajtó:
"Az a gránit térkő Székesfehérváron nagyon jól néz ki, egyforma a fugaköz."
"Az a Semmelrock felhajtó Győrben klassz lett, látszik a pontos aljzat."
Burkolat / fürdőszoba:
"Az a porcelán burkolat Miskolcon nagyon jól néz ki, egyenes a fugaköz."
"Az a fürdőszoba Kecskeméten klassz lett, jó a padló-fal átmenet."
Ablak / nyílászáró:
"Az a Schüco antracit ablak Győrben nagyon jól néz ki, szép a mélyen beépített tok."
"Az a Rehau bukó-nyíló ablak Pécsett jól néz ki, pontos a külső párkány dőlésszöge."
"Az a Veka nyílászáró Debrecenben rendesen néz ki, egyforma az üvegtáblák hézaga."
"Az a bejárati ajtó Miskolcon jól néz ki, szép a küszöb és a padlóburkolat illesztése."
"Az a Velux ablak Szentendrén nagyon jól néz ki, precízen van a villogópajzs visszavezetve."
"Az a toló nyílászáró Szegeden klassz lett, látszik a vékony tokprofil."
"Az a harmonika ajtó Kecskeméten jól néz ki, egyforma a szárnyak közti hézag."
"Az a beépített redőnytokos ablak Kaposváron jól néz ki, nem lóg ki a falból."
"Az a sötétbarna PVC ablak Egerben rendesen néz ki, szép az aranytölgy szín a homlokzaton."
Kert / terasz:
"Az a WPC terasz Debrecenben nagyon jól néz ki, egyforma a deszkatávolság."
"Az a kert Győrben komolyan néz ki, látszik a tereplépcsők vonala."
Homlokzat:
"Az a homlokzat Budapesten nagyon jól néz ki, egyenesek a sarkok élei."
"Az a dryvit homlokzat Egerben rendesen néz ki, egyforma a vakolat vastagsága."
Festés / mázolás:
"Az a homlokzatfestés Debrecenben nagyon jól néz ki, egyenletes a szín lefedése."
"Az a beltéri festés Miskolcon rendesen néz ki, szép az ajtókeretek körüli elvágás."

MIKOR HASZNÁLD A FALLBACKET (kötelező, kivétel nélkül):
- Nem látsz befejezett szakipari munkát a fotón
- Csak szöveget, logót, dokumentumot, tanúsítványt, AI képet látsz
- Nem tudod meghatározni a várost
- NEM tudsz konkrét anyagnevet, márkanevet vagy szakmai részletet megnevezni — a városnév egyedül NEM ELÉG
- Általános anyagnevek mint "tapéta", "festés", "burkolat", "tető" ÖNMAGUKBAN nem elégségesek — kell mellé szín, technika, márka vagy konkrét részlet. Ha nincs ilyen → fallback.
- A bejegyzés NEM a cég saját, elkészült munkája — pl. "megosztotta", "megosztás", "közösen [valakivel]", "megjelölve" felirat látszik a bejegyzés tetején, vagy a szöveg egy VÉLEMÉNY/ÉRTÉKELÉS valaki mástól (nem a cég sajét hangja), vagy egy beszállító/partner reklámja. Ha nem egyértelmű, hogy EZ A CÉG csinálta a munkát → fallback. Ez fontosabb, mint bármilyen szakmai részlet — egy rossz céghez kötött munka rosszabb, mint egy általános nyitósor, mert a tulajdonos emlékszik, mit csinált.
Fallback: "Rátaláltam a Facebook-oldalukra [városban]." — a városnév KÖTELEZŐEN ragos: "Kaposváron", "Győrben", "Budapesten" — SOHA nem alanyeset ("Kaposvár", "Győr").

SZIGORÚ SZABÁLYOK:
- CSAK a nyitósort add vissza. Semmi más szöveg.
- Semmi emoji, semmi felkiáltójel, semmi nagykötőjel.
- Ne másold szó szerint a példákat.
- A városnév mindig ragos alakban legyen.
- Mindig használj helyes magyar ékezeteket (á, é, í, ó, ö, ő, ú, ü, ű).
- Max 14 szó.
- MINDEN mondat "Az a" szavakkal kezdődjön - kivétel csak a fallback.
- TILOS hírcím stílus: "Szép lett a X...", "Szép munka lett a...", "Elkészült a..." - TILOSAK.
- TILOS passzív befejezés: "szép lett kivitelezve", "szépen lett kivitelezve", "szép lett fedve", "szép lett kialakítva", "szép lett megtervezve" — TILOSAK.
- A legtermészetesebb forma: "jól néz ki [a részlet]" vagy egyszerűen leírod amit látsz. Példa: "Az a Bramac cserép Debrecenben nagyon jól néz ki, egyforma a fedési minta." — ez természetesebb mint a "lett felfektetve" forma. Használd ezt előnyben.
- TILOS általános befejezés: "jól lett megcsinálva", "jól lett bekötve", "jól lett elvégezve", "jól lett felszerelve", "szép a fedése", "szép a munkája", "szép munka", "szép a kialakítás", "szép a rendszer kiépítése", "klasszul lett kialakítva", "klassz lett kialakítva" — TILOSAK. Az ige mellé MINDIG kell konkrét anyagnév, márkanév vagy szakmai terminus.
- KÖTELEZŐ: minden nyitósornak kell egy VESSZŐVEL elválasztott második tagmondat egy konkrét megfigyeléssel. Pl. "...jól néz ki, pontos a tok körüli simítás." — a mondat soha nem végződhet csak a városnévnél vagy a tárgy nevénél.
- TILOS "Budapestben" — mindig "Budapesten". Helyes ragozás: Budapesten, Miskolcon, Pécsen, Győrben, Szegeden, Debrecenben.
- Ha az egyetlen konkrét szó a városnév — ez NEM elég, használd a fallbacket.
- Ablak esetén SOHA ne írj "műanyag ablak" - írj "PVC ablak" vagy a márkanevet.
- Ha márkanevet látsz a fotón, azt KÖTELEZŐ beépíteni — nem opcionális, nem hagyható ki.
- MÁRKA-NICHE EGYEZÉS: Klíma/hűtés márkák (Daikin, Mitsubishi, KAISAI, Midea, Gree, Toshiba) csak klíma cégeknek valók. Kazán/fűtés márkák (Vaillant, Viessmann, Bosch, Buderus, Junkers) csak fűtésszerelő cégeknek. Ne keverd: ha klíma cégről van szó, ne írj kazánról, ha fűtéscégről van szó, ne írj klímáról.
- SOHA ne írj "szép a kivitelezés" — mindig nevezd meg PONTOSAN mi szép: fugaköz, illesztés, lejtés, csőelvezetés, korcok távolsága, gerincvonal, stb.
- Ha a szín vagy anyagszín látható (antracit, grafitszürke, natúr, fehér, barna), azt is nevezd meg."""

# -- Hungarian a/az helper ----------------------------------------------------
_HU_VOWELS  = set("aáeéiíoóöőuúüűAÁEÉIÍOÓÖŐUÚÜŰ")
_HU_FRONT   = set("eéiíöőüűEÉIÍÖŐÜŰ")
_HU_BACK    = set("aáoóuúAÁOÓUÚ")

def _az(name: str) -> str:
    return "az" if name and name[0] in _HU_VOWELS else "a"

def _nak_nek(name: str) -> str:
    """Return 'nak' or 'nek' based on vowel harmony of the last vowel in name."""
    for ch in reversed(name):
        if ch in _HU_FRONT:
            return "nek"
        if ch in _HU_BACK:
            return "nak"
    return "nak"

# -- Hungarian body -----------------------------------------------------------
_HU_TRADE_LABELS = [
    # "fűtés-hűtés" removed — every heating company writes this, causes fűtésszerelő→klíma misclassification
    (["klíma", "légkondicionáló", "hőszivattyú", "heat pump", "hvac",
      "daikin", "mitsubishi", "kaisai", "midea", "gree", "toshiba", "panasonic", "lg klíma", "samsung klíma",
      "klímaszerelő", "klímaszerelés", "hűtés", "split", "klíma telepítés", "klíma beépítés"], "klíma és hőszivattyú"),
    (["napelem", "napelemrendszer", "napelemes", "pv rendszer", "longi", "jinko", "canadian solar",
      "risen", "trina", "huawei inverter", "sma inverter", "fronius", "napkollektor",
      "szolár", "solar"], "napelemes"),
    (["tetőfedő", "tetőfedés", "cserép", "pala", "epdm", "lapostető", "bramac", "tondach",
      "eternit", "velux", "fakro", "creaton", "lindab", "tegola", "zenit", "icopal",
      "korcolt", "oromszegély", "ereszcsatorna", "kémény", "tetőszerkezet", "héjazat"], "tetőfedő"),
    (["villanyszerelő", "villanyszerelés", "elektromos szerelés", "áramkör",
      "kismegszakító", "elosztó", "elektromos hálózat", "villanyszerelő mester"], "villanyszerelő"),
    (["vízvezetékszerelő", "vízszerelő", "vízvezeték", "csőszerelő", "csőszerelés",
      "szaniter", "vízszereléс", "szivattyú", "vízlágyító"], "vízvezetékszerelő"),
    (["kőműves", "kőműves mester", "falazó mester", "kőművesmunka",
      "téglafalazás", "blokk falazás"], "kőműves"),
    (["festő", "mázoló", "festés", "mázolás", "homlokzatfestés", "beltéri festés",
      "szobafestő", "szobafestés", "szoba festés", "fal festés", "festési munkák",
      "belső festés", "külső festés", "festő-mázoló", "festő mázoló",
      "diszperzit", "latex", "falfesték", "színezés"], "festő"),
    (["gipszkarton", "szárazépítés", "álmennyezet", "válaszfal", "knauf", "rigips",
      "ásványgyapot", "hangszigetelés"], "gipszkartonos"),
    (["burkolat", "burkoló", "burkolás", "csempe", "járólap", "gránit", "mészkő",
      "porcelán", "mozaik", "padlólap", "falilap", "fürdőszoba burk"], "burkoló"),
    (["parketta", "laminált", "padló", "padlóburkoló", "padlóburkolás", "vinyl",
      "padlólap lerakás", "fa padló"], "padlóburkoló"),
    (["fürdőszoba felújítás", "fürdőszoba átalakítás", "fürdő felújítás",
      "komplett fürdő", "fürdőszoba csere"], "fürdőszoba felújító"),
    (["asztalos", "bútorasztalos", "egyedi bútor", "belsőajtó", "fa lépcső",
      "beépített bútor", "konyhabútor", "gardrób"], "asztalos"),
    (["medence", "úszómedence", "medenceépítés", "medencefelújítás",
      "jacuzzi", "szauна", "wellness"], "medenceépítő"),
    # garázskapu: only unambiguous garage/shading terms — "redőny","árnyékoló","roló","napvédő"
    # removed because they appear on every ablak/nyílászáró company page and cause misclassification
    (["garázskapu", "kapuautomatika", "tolókapu", "szekcionált kapu",
      "zip árnyékoló", "zip rendszer", "reluxa", "pergola", "somfy", "raffstore",
      "redőny szerelő", "árnyékoló rendszer", "napvédő rendszer"], "garázskapu és árnyékoló"),
    (["térkő", "térkövezés", "felhajtó", "járda", "semmelrock", "bradstone",
      "marshalls", "kocsibeálló", "parkoló burkolat"], "térkövező"),
    (["kertépítés", "kertépítő", "kertrendezés", "kertészeti", "öntözőrendszer",
      "öntözés", "kerti", "gyep", "fűnyírás", "sövény", "kerítésépítés", "kerítés"], "kertépítő"),
    (["terasz", "teraszburkolat", "deszka", "wpc", "bangkirai", "teak", "kompozit deszka",
      "faburkolat"], "terasz és deszka"),
    (["ablak", "nyílászáró", "ablakcsere", "ablakbeépítés", "schüco", "rehau", "veka",
      "internorm", "gealan", "kbe", "műanyag ablak", "pvc ablak", "fa ablak"], "nyílászáró"),
    (["homlokzat", "homlokzatszigetelés", "külső szigetelés", "hőszigetelés",
      "nemesvakolat", "dryvit", "styroport", "homlokzatfelújítás"], "homlokzatszigetelő"),
    (["generálkivitelező", "általános kivitelező", "komplex felújítás",
      "komplett felújítás", "teljes körű felújítás", "általános felújítás",
      "építési vállalkozó", "több szakmát"], "felújító kivitelező"),
    # épületgépész/épületgépészet added — Zengl-type companies were falling through to felújító
    (["kazán", "gázkazán", "fűtésszerelő", "fűtésszerelés", "radiátor", "padlófűtés",
      "gázszerelő", "gázszerelés", "gázvezeték", "gázbekötés", "víz-gáz", "vízgáz",
      "épületgépész", "épületgépészet", "épületgépészeti", "fűtés-hűtés", "hőközpont",
      "vaillant", "viessmann", "nibe", "bosch", "stiebel", "worcester", "ariston",
      "baxi", "ferroli", "immergas", "junkers", "buderus"], "fűtésszerelő"),
]

# _hu_detect_niche imported from hu_niche.py (name-first + weighted text scoring)

# History, because this has been flipped twice: originally off ("Diák vagyok"
# alone reads as cheap/inexperienced and invites haggling before a price is
# named), then on after real rejections cited age/trust ("you're too young")
# rather than price, reframed so it sat next to the proof line instead of
# standing alone. Now off again on request. The True branch below is left
# intact so this is a one-word revert, not a rewrite.
HU_MENTION_STUDENT = False

# The one piece of proof that matters: a fill that became a real job. Pointing
# at the Facebook post rather than a URL keeps it verifiable without tripping
# Meta's link filter on a first DM. Update the figure as more clients convert.
# Compressed from two sentences to one to cut length.
# Deliberately says "az egyik ügyfelemnél" and NOT "egy másik költöztető
# cégnél": naming the prospect's own trade here would read better, but it
# would be a claim about having a client in THAT trade, which isn't true for
# every niche in the list. Vague-but-true beats specific-but-invented.
HU_PROOF_LINE = (
    "Az egyik ügyfelemnél az első héten be is jött rajta egy 6 milliós munka, "
    "van róla egy poszt a profilomon."
)


# Per-niche question list — now the ONLY place the message signals which trade
# it was built for. The pitch used to also name the audience outright
# ("árkalkulátor chatbotokat készítek költöztető cégeknek"), but that sentence
# and the closing "Önnek is összeraktam egyet" were two sentences doing one
# job; the opening now leads with the finished build instead, and this list
# proves it was made for their trade more concretely than a label did.
# Defaulted to trade wording, which is fine for the 20 construction niches and
# wrong for the non-trade ones — "mekkora felület, milyen anyag" means nothing
# to a moving company or a dentist, and a question list that obviously doesn't
# apply to you is the fastest possible tell that the message went to a
# thousand people.
# Phrased from the END CUSTOMER's side (that's who the bot talks to). Formal
# for fogászat because a patient is addressed formally in Hungarian, informal
# for the rest to match how these trades actually talk to their customers.
_HU_NICHE_QUESTIONS = {
    "költöztető":  "hány szobás lakás, hányadik emeletről hova, van-e lift",
    "fogászat":    "milyen kezelés érdekli, hány fogról van szó, mikorra tervezi",
    "raktárbérlő": "mekkora helyre lenne szüksége, mennyi időre, mikortól",
    # Landscaping stays on the PHOTO path (finished gardens have real, visible
    # craft), but the default "mekkora felület, milyen anyag" is a burkoló
    # question, not a kertépítő one.
    "kertépítő":   "mekkora a kert, mit szeretne benne, mikorra kellene",
    # Lookup-table niches: these are the actual pricing variables, which is
    # the whole reason they fit a quoting bot. Getting them right is the one
    # place the message proves it was built for their business and not
    # pasted from a template.
    # klíma fell to the default, which asks "mekkora felület, milyen anyag" —
    # a burkoló question, and worse, the kind a layperson cannot answer for an
    # AC install. Keresztesi Károly (k-klima) made the point directly: nobody
    # knows how many metres of pipe, how many idom, or where the power comes
    # from, and a wrong answer poisons the whole number. Ask only what the
    # customer actually knows.
    "klíma":                     "hány szobába kellene, mekkorák a szobák, hová kerülne a kültéri egység",
    "autóüveg":                  "milyen autó, javítás vagy csere kell, van-e rá casco",
    "háztartásigép-szerviz":     "milyen készülék, melyik márka, mi a hibajelenség",
    "telefon- és laptopszerviz": "milyen készülék, mi a hiba, mennyire sürgős",
    "fordítóiroda":              "milyen nyelvpár, hány oldal, mikorra kell",
    "autókozmetika":             "milyen autó, melyik csomag érdekli, mikorra szeretné",
    "hajbeültetés":              "melyik területet érinti, mit szeretne elérni, mikorra tervezi",
    # "mekkora területről van szó" was the phrasing you would use to price a
    # FLOOR. Hair transplants are priced per graft off the affected area and
    # the patient's goal, so "melyik területet érinti" is both the clinic's
    # own language and the question they actually ask.
}
_HU_DEFAULT_QUESTIONS = "mekkora felület, milyen anyag, mikorra kellene"


# Per-niche named proof. Where there is a real client in that trade, the
# message names them and the problem the build solved; every other niche falls
# back to HU_PROOF_LINE above (anonymous, with the 6M figure).
#
# WHY IT IS PHRASED THIS WAY — the problem is stated as the TRADE's general
# problem, not as a private fact about that client. "A gazszerelesnel az a
# legnagyobb gond, hogy..." is an observation about the industry: true, and
# the reader agrees before finishing the sentence. The only thing asserted
# about the named client is WHAT WAS BUILT FOR THEM, which is live on their
# own website and findable in one search. Inventing a measured outcome for a
# named company would be the one claim a prospect could disprove by picking
# up the phone.
#
# Names must match landscale.agency/hu/case-studies exactly, or the
# verification this whole mechanism depends on breaks.
_NM_BAU_PROOF = (
    "A felújításban az a legnagyobb gond, hogy mindenki az árral kezdi, "
    "pedig felmérés nélkül nem lehet pontosat mondani. Az NM Baunak is pont "
    "ezt építettem, ott él most is."
)
_HU_NICHE_PROOF = {
    "fürdőszoba felújító": _NM_BAU_PROOF,
    "felújító kivitelező": _NM_BAU_PROOF,
    "fűtésszerelő":
        "A gázszerelésnél az a legnagyobb gond, hogy a hívások fele csak "
        "árat kérdez, aztán soha nem hallani róluk többet. A kecskeméti Hírös "
        "Kft-nek is pont ezt építettem, ott él most is.",
    "klíma":
        "A klímaszerelésnél az a legnagyobb gond, hogy szezonban rengeteg a "
        "megkeresés, de a fele csak nézelődik. A Klímaszerelő Kecskemétnek is "
        "pont ezt építettem, ott él most is.",
}


# Facebook shows a Page's website in its intro block, and auto_next() already
# scrapes the whole page innerText — so whether a lead has a site is knowable
# for free, with no extra request, from text we are holding anyway.
#
# Platform domains are excluded: every FB page mentions facebook.com, and half
# of them link a Messenger or WhatsApp shortlink. Counting those as "has a
# website" would make the detector always true and therefore useless.
_HU_PLATFORM_DOMAINS = (
    "facebook.com", "fb.com", "fb.me", "m.me", "messenger.com", "instagram.com",
    "wa.me", "whatsapp.com", "youtube.com", "youtu.be", "tiktok.com",
    "linkedin.com", "google.com", "goo.gl", "maps.app", "gmail.com",
    "twitter.com", "x.com", "pinterest.com", "w3.org", "schema.org",
    # Directories and marketplaces. A Page's own text often cites its listing
    # on one of these, and it is never the company's own site.
    "cylex.hu", "telefonkonyv.hu", "nemzeticegtar.hu", "ceginformacio.hu",
    "ceginfo.hu", "opten.hu", "aranyoldalak.hu", "arukereso.hu", "jofogas.hu",
    "hasznaltauto.hu", "ingatlan.com", "startlap.hu", "tripadvisor.com",
    "booking.com", "yelp.com", "foursquare.com", "bing.com", "wikipedia.org",
)
_HU_DOMAIN_RE = _re.compile(
    r"\b([a-z0-9][a-z0-9\-]{1,62}(?:\.[a-z0-9\-]{2,63})*"
    r"\.(?:hu|com|net|eu|org|info|biz|shop|online|site|website|pro|club|xyz"
    r"|digital|studio|agency|services|solutions|expert))\b",
    _re.IGNORECASE)


# Facebook's own intro panel headings, English and Hungarian UI. The Links /
# Contact info block is where a Page states ITS OWN website — everywhere else
# on the page a domain might belong to a supplier, a shared post, or a partner.
# Facebook's intro panel headings, English and Hungarian UI. Split in two,
# because the two sections carry different levels of certainty.
#
# A "Links" row is the Page DECLARING its own website — unambiguous.
# A "Contact info" row holds an email, whose domain is only probably theirs,
# and in a short page the window after it can run on into the feed. So a
# domain found there must also look like the company's name before it counts.
_HU_LINK_MARKERS = ("links", "hivatkozások", "hivatkozasok", "linkek",
                    "weboldal", "website", "web oldal", "honlap", "weblap",
                    "webhely", "web site")
_HU_CONTACT_MARKERS = ("contact info", "elérhetőség", "elerhetoseg",
                       "elérhetőségek", "kapcsolatfelvételi adatok",
                       "névjegy", "nevjegy", "about", "információ", "informacio")
_HU_PANEL_WINDOW = 220       # panel rows are short, but the URL is rarely first


def _hu_all_domains(chunk: str):
    """Every non-platform, non-directory domain in `chunk`, in order, deduped."""
    seen = set()
    for m in _HU_DOMAIN_RE.finditer(chunk):
        dom = m.group(1).lower()
        if dom.startswith("www."):
            dom = dom[4:]
        if any(dom == p or dom.endswith("." + p) for p in _HU_PLATFORM_DOMAINS):
            continue
        if dom in seen:
            continue
        seen.add(dom)
        yield dom


def _hu_first_domain(chunk: str) -> str:
    return next(_hu_all_domains(chunk), "")


def _hu_windows_after(page: str, markers) -> list[str]:
    """Every window following every occurrence of every marker, in page order.

    The old version took min() over the FIRST hit of each marker, so a post
    saying "nezze meg a weboldalunkat" anchored the search near the top of the
    feed and the window never reached the intro panel further down.
    """
    low = page.lower()
    starts = []
    for mk in markers:
        i = low.find(mk)
        while i != -1:
            starts.append(i)
            i = low.find(mk, i + 1)
    return [page[i:i + _HU_PANEL_WINDOW] for i in sorted(set(starts))]


def _hu_domain_in(page: str, markers) -> str:
    for w in _hu_windows_after(page, markers):
        dom = _hu_first_domain(w)
        if dom:
            return dom
    return ""


# Tokens that identify a TRADE or a PLACE, not a company. "klimaszerelo.hu"
# shares "klima" with every klima company in the country, so a name match on
# one of these alone proves nothing — it is exactly how a competitor's or a
# supplier's domain gets claimed as the lead's own.
_HU_GENERIC_NAME_TOKENS = {
    "klima", "klimaszerelo", "kert", "kertepites", "kertepito", "teto",
    "tetofedo", "tetofedes", "auto", "autos", "autoszerviz", "futes",
    "futesszerelo", "gazszerelo", "vizszerelo", "epito", "epitoipari",
    "epuletgepeszet", "burkolo", "burkolas", "festo", "festes", "szerelo",
    "szerviz", "javitas", "csere", "ablak", "ajto", "uveg", "szelvedo",
    "autouveg", "felujitas", "felujito", "kivitelezo", "csaladi",
    "bau", "kft", "bt", "zrt", "kkt", "nyrt", "team", "group", "profi",
    "mester", "szaki", "centrum", "center", "hungary", "magyar", "trade",
    # Cities: real domains do contain them, but a city alone is not identity.
    "budapest", "debrecen", "szeged", "miskolc", "pecs", "gyor", "kecskemet",
    "nyiregyhaza", "szekesfehervar", "szombathely", "sopron", "eger",
    "tatabanya", "kaposvar", "veszprem", "zalaegerszeg", "szolnok",
    "bekescsaba", "dunaujvaros", "salgotarjan", "erd", "hodmezovasarhely",
}


def _hu_domain_matches_name(domain: str, company_name: str,
                            strict: bool = False) -> bool:
    """Does the domain plausibly belong to THIS company?

    Guards the weaker Contact-info path. Without it, a shared supplier post
    sitting just below the panel ("Megosztotta: a beszallito-alkatresz.hu
    akciója") got read as the lead's own site, which would then have fetched a
    stranger's website and told this lead theirs needed work.
    """
    dom = _norm(domain.split(".")[0]).replace("-", "")
    for tok in _re.split(r"[^0-9a-záéíóöőúüű]+", _norm(company_name)):
        if len(tok) < 4 or tok not in dom:
            continue
        # strict=True is used by the whole-page fallback, where the domain has
        # no panel heading vouching for it and a generic trade word would let
        # any competitor's URL through.
        if strict and tok in _HU_GENERIC_NAME_TOKENS:
            continue
        return True
    return False


def _hu_has_website(text_data: str) -> str:
    """The lead's own domain if Facebook's own panel states one, else "".

    Deliberately NO whole-page scan. The feed is full of other people's
    domains, and the two errors here are not symmetrical: a missed website
    yields "láttam, hogy weboldala nincs", which at worst gets corrected — and
    a correction is a REPLY, which is the point. A wrongly-claimed website
    yields a confident sentence about a site that isn't theirs, with nothing
    to say when they ask what needs improving.
    """
    page = text_data
    for hdr in ("--- COMPANY NAME ---", "--- TRADE TYPE ---", "--- LISTED CITY ---"):
        page = page.replace(hdr, " ")
    name = _td_section(text_data, "COMPANY NAME")

    # 1. A declared link: trusted on its own.
    dom = _hu_domain_in(page, _HU_LINK_MARKERS)
    if dom:
        return dom

    # 2. An email domain: only if it looks like the company.
    dom = _hu_domain_in(page, _HU_CONTACT_MARKERS)
    if dom and _hu_domain_matches_name(dom, name):
        return dom

    # 3. Any domain anywhere on the page. A name-matched one wins if there is
    #    one, because it is the likeliest to be theirs, but ANY non-platform,
    #    non-directory domain counts as "they have a website".
    page_domains = list(_hu_all_domains(page))
    for dom in page_domains:
        if _hu_domain_matches_name(dom, name, strict=True):
            return dom
    return page_domains[0] if page_domains else ""


def _hu_site_is_theirs(domain: str, text_data: str) -> bool:
    """Is `domain` confirmed to belong to THIS lead, rather than merely present
    on their page?

    Detection is deliberately loose (see above) so the "weboldala nincs" line
    never fires wrongly. This is the separate, strict question, and it gates
    the only sentence that asserts something about the site's contents. Getting
    it wrong means telling someone their website is slow when the URL belonged
    to a supplier, and having nothing to say when they ask what needs fixing.
    """
    if not domain:
        return False
    page = text_data
    for hdr in ("--- COMPANY NAME ---", "--- TRADE TYPE ---", "--- LISTED CITY ---"):
        page = page.replace(hdr, " ")
    if domain == _hu_domain_in(page, _HU_LINK_MARKERS):
        return True            # the Page declared it under its own Links row
    return _hu_domain_matches_name(domain, _td_section(text_data, "COMPANY NAME"),
                                   strict=True)


# Words a Hungarian trade site uses when it ALREADY offers an instant price.
# If none appear, the visitor's only route to a number is picking up the phone,
# which is exactly the gap this product fills and a specific, checkable thing
# to have noticed.
_HU_QUOTE_PATH_WORDS = (
    "árkalkulátor", "arkalkulator", "kalkulátor", "kalkulator",
    "azonnali ár", "azonnali arajanlat", "online ajánlat", "online ajanlat",
    "árajánlat kérése", "ajánlatkérő", "ajanlatkero", "árbecslés", "arbecsles",
)


def _hu_site_check(domain: str) -> dict:
    """Actually FETCH the lead's site, so "megnéztem a weboldalát" is a fact.

    Without this the improve-line is a bluff: nothing else in the DM pipeline
    ever loads their website, only their Facebook page. A prospect asking
    "mi az?" would get silence, which is worse than never mentioning it.

    Short timeout, fails open: a site that will not load in 6s yields no claim
    rather than delaying the lead or inventing one. Findings are printed so
    the log holds the answer to "mi az?" while the message stays short.
    """
    out = {"ok": False, "secs": 0.0, "quote_path": False, "mobile": True}
    try:
        import time as _t
        import requests as _rq
        hdrs = {
            "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                           "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"),
            "Accept-Language": "hu-HU,hu;q=0.9",
        }
        t0 = _t.time()
        try:
            r = _rq.get("https://" + domain, timeout=6, allow_redirects=True, headers=hdrs)
        except _rq.exceptions.SSLError:
            # Small Hungarian trade sites are very often on an expired or
            # mismatched certificate, which says nothing about whether the
            # business is real. aqua-system.hu — straight off a real lead's
            # page — fails cert validation on the apex domain but serves a
            # perfectly good 24KB site at www over https.
            #
            # Plain http does NOT rescue it (the server 302s back to https and
            # the same cert fails), so the retry has to be https with
            # verification off. Narrow and deliberate: this request only GETs
            # public marketing HTML to count keywords. Nothing is sent, no
            # credential is involved, and nothing from the response is trusted
            # beyond "does this page contain the word árkalkulátor".
            print(f"  {domain}: bad certificate, retrying without verification")
            import urllib3 as _u3
            _u3.disable_warnings(_u3.exceptions.InsecureRequestWarning)
            r = _rq.get("https://" + domain, timeout=6, allow_redirects=True,
                        headers=hdrs, verify=False)
        out["secs"] = round(_t.time() - t0, 1)
        html = (r.text or "").lower()
        out["ok"] = r.status_code < 400
        out["quote_path"] = any(w in html for w in _HU_QUOTE_PATH_WORDS)
        out["mobile"] = "viewport" in html
    except Exception as e:
        print(f"  Site check failed for {domain}: {type(e).__name__}")
    return out


# THREE REGISTERS, because these are three different businesses and the pitch
# was written for exactly one of them.
#
#   trade   (default) — low volume, high ticket, priced by site visit. Wants
#                       FEWER enquiries: every dead lead costs him a day.
#   medical           — priced only after examination. "6 milliós munka" is a
#                       construction number (a transplant is 1-2 million) and
#                       announces the text was reused; a clinic does kezelés,
#                       not munka.
#   volume            — high volume, low ticket, priced off a lookup (car +
#                       which glass, device + fault, word count + language).
#                       Wants MORE enquiries, not fewer, so "kivel érdemes
#                       foglalkozni" describes a problem they do not have —
#                       there is nobody to filter out. And hedging with
#                       "tájékoztató becslés, nem fix ár" undersells: for these
#                       trades the price really is knowable up front, which is
#                       the strongest thing that can be said about the bot.
_HU_MEDICAL_NICHES = {"fogászat", "hajbeültetés"}
_HU_VOLUME_NICHES = {
    "autóüveg", "autókozmetika", "fordítóiroda",
}

# These two used to sit in _HU_VOLUME_NICHES and get the árkalkulátor pitch
# with "volume" mech/notify copy ("majd rögtön meg is mondja, hogy mennyibe
# kerül"). That claim doesn't hold for a repair shop: nobody can price a
# phone or a fridge repair without physically seeing the device, so an
# instant-quote bot is the wrong product for this niche entirely (2026-09-03
# conversation — [[nmbau-austria-coverage]]'s sibling decision for this
# niche, not yet its own memory file).
#
# What actually fits a one-person repair shop is what stalls THEM, not their
# customers: a missed call while their hands are in a device, and — if they
# have no site — a lookup that finds nothing. _hu_pick_body/_hu_followup_text
# route these two straight to _hu_repair_pitch()/_HU_REPAIR_FOLLOWUP instead
# of the árkalkulátor copy every other niche gets.
_HU_REPAIR_NICHES = {"háztartásigép-szerviz", "telefon- és laptopszerviz"}

# Neither register cites a forint figure: there is no clinic or gyorsszerviz
# client to name, and inventing a plausible-sized number is the one claim a
# prospect can check.
#
# "Pár helyen már élesben fut" (used to read this way) sits right after
# "Készítettem Önnek egy árkalkulátort" with nothing marking whose "pár hely"
# it means — a real prospect could read it as their own instance already
# being live, which is exactly the claim [[outreach-no-unverified-claims]]
# rules out (never imply a demo is already live under the recipient's own
# name). "Más cégeknél" makes the subject explicit: the live examples are
# OTHER businesses, proof the product works, not a claim about theirs.
_HU_ALT_PROOF = "Más cégeknél már élesben fut, van róla egy poszt a profilomon."

_HU_REGISTER_COPY = {
    "medical": {
        # A clinic's first objection is not "do not pin me to a number" but
        # "you cannot price this without seeing the patient" — which is true,
        # and which the trade wording never answered. This is the medical
        # counterpart of "A felmérést nem váltja ki".
        "mech": ("majd megmondja, honnan indul az ár, és mi az, ami azon felül "
                 "jön. A konzultációt nem váltja ki. "),
        "notify": ("Ön pedig megkapja a nevét és a számát, így azokhoz jut el, "
                   "akik tényleg tervezik a beavatkozást, nem csak érdeklődnek."),
    },
    "volume": {
        # No hedge: the price genuinely is a lookup here, so say so.
        "mech": "majd rögtön meg is mondja, hogy mennyibe kerül, éjjel-nappal. ",
        # Even here the promise is better customers, not saved minutes: a
        # real installer (Keresztesi Károly, k-klima) pointed out that time
        # saved is worth nothing next to landing the job.
        "notify": ("Ön pedig megkapja a nevét és a számát, így nem mindenkivel "
                   "tölti az idejét, hanem azokkal, akik tényleg hozzák a munkát."),
    },
}
_HU_TRADE_MECH = ("majd megmondja az alapárat, és azt is, mi az, "
                  "ami azon felül jön. ")
# The pitch is deliberately NOT "fewer wasted callouts". For klíma and every
# other trade where the survey is where the sale and the margin happen, the
# visit is the asset, not the overhead — so the promise is better customers
# reaching that visit, never fewer visits.
_HU_TRADE_NOTIFY = ("Ön pedig megkapja a nevét és a számát, így nem több "
                    "érdeklődőt kap, hanem jobbakat.")


def _hu_register(niche: str) -> str:
    if niche in _HU_MEDICAL_NICHES:
        return "medical"
    if niche in _HU_VOLUME_NICHES:
        return "volume"
    return "trade"


# ─── Repair-niche pitch (háztartásigép-szerviz, telefon- és laptopszerviz) ──
# Framed entirely as money the OWNER is already losing, never as a nicer
# experience for his customers — same logic as [[outreach-sell-time-not-
# leads]]: he isn't short on people calling him, he's leaking jobs he
# already won because he can't reach the phone mid-repair. No pricing claim
# anywhere in here, unlike the árkalkulátor pitch every other niche gets.
#
# This is a real conversational AI VOICE agent (answers the phone, holds an
# actual conversation), not an SMS auto-reply — the SMS version was the
# first draft, replaced after 2026-09-03 research showed phone/appliance
# repair specifically has among the worst call-answer rates of any trade
# (74.1% of phone-repair calls, ~30% of appliance-repair calls go
# unanswered), and these customers are call-people, not text-people, so
# a voice answer fits how they actually behave. It also books straight into
# the calendar rather than requiring the owner to call back afterward —
# the callback-required framing was cut on explicit instruction, since
# making the owner do a manual follow-up step undersells what the system
# actually does.
#
# STYLE: no em dash anywhere in this block (standing rule, see the "No em
# dash" note further down in this file) and Ön/Önnek throughout, matching
# every other niche's register.
_HU_REPAIR_INTRO = (
    "Készítettem egy rendszert, ami felveszi Ön helyett a telefont, amikor "
    "épp szerelés közben nem tud válaszolni."
)

# Per-niche PAIN + MECH, not shared text with a device name swapped in — the
# two niches have genuinely different customer psychology, not just
# different nouns. A phone/laptop customer is usually without their only
# device and wants same-day help; a fridge/washing-machine customer is
# inconvenienced but rarely desperate the same way, so the urgency framing
# differs, not just "hűtő" vs "telefon".
_HU_REPAIR_COPY = {
    "háztartásigép-szerviz": {
        "pain": (
            "Ilyenkor a hívó fél jellemzően egyszerűen a következő "
            "szervizt hívja fel, és az a meló elvész, pedig már Öné "
            "lett volna."
        ),
        "mech": (
            "A rendszer csak akkor válaszol, ha Ön nem tudja felvenni: "
            "ilyenkor valódi beszélgetést folytat az ügyféllel, "
            "kikérdezi, milyen géppel van gond és mennyire sürgős, "
            "majd akár rögtön be is foglalja egy időpontot a naptárába, "
            "anélkül hogy Önnek vissza kellene hívnia."
        ),
    },
    "telefon- és laptopszerviz": {
        "pain": (
            "Ilyenkor a hívó fél, akinek gyakran pont a törött készülék "
            "miatt sürgős az ügy, egyszerűen a következő szervizt hívja "
            "fel, és az a meló elvész, pedig már Öné lett volna."
        ),
        "mech": (
            "A rendszer csak akkor válaszol, ha Ön nem tudja felvenni: "
            "ilyenkor valódi beszélgetést folytat az ügyféllel, "
            "kikérdezi, milyen telefonnal vagy laptoppal van gond és "
            "mennyire sürgős, majd akár rögtön be is foglalja egy "
            "időpontot a naptárába, anélkül hogy Önnek vissza kellene "
            "hívnia."
        ),
    },
    # Not a real niche in _resolve_niche's vocabulary (hu_niche.py) — a full
    # separate niche means its own _NICHE_VOCAB, _TEXT_KEYWORDS, opener
    # instructions and contradiction list, which is a lot of scaffolding for
    # what is a narrow slice of "háztartásigép-szerviz" leads. Instead this
    # is selected by _hu_repair_pitch's own coffee-machine keyword check
    # below, layered on top of the normal niche resolution. Real example that
    # prompted this: "Kávégép Szerviz Cegléd" — Delonghi/Philips/Jura service
    # for both home AND commercial (ipari kávégépek, "AtisCoffee" affiliation)
    # machines. "háztartási géppel van gond" reads oddly here on two counts:
    # a café's espresso machine breaking is not thought of as a "háztartási
    # gép" the way a fridge is, and a chunk of this shop's real customers are
    # BUSINESSES, not households, for whom a dead machine is lost revenue
    # every hour, not a household inconvenience — closer to the phone/laptop
    # urgency case than the generic appliance one.
    "kávégép-szerviz": {
        "pain": (
            "Ilyenkor a hívó fél, akinek gyakran egy kávézó vagy iroda "
            "működése áll vagy bukik a gépen, egyszerűen a következő "
            "szervizt hívja fel, és az a meló elvész, pedig már Öné "
            "lett volna."
        ),
        "mech": (
            "A rendszer csak akkor válaszol, ha Ön nem tudja felvenni: "
            "ilyenkor valódi beszélgetést folytat az ügyféllel, "
            "kikérdezi, milyen kávégéppel van gond és mennyire sürgős, "
            "majd akár rögtön be is foglalja egy időpontot a "
            "naptárába, anélkül hogy Önnek vissza kellene hívnia."
        ),
    },
}

# Keyword check for the kávégép-szerviz carve-out above. Matches company
# name and scraped page text alike (both already live in text_data by the
# time this runs), covers home espresso machines and industrial/commercial
# ones ("ipari kávégép", "presszógép") under the same copy, since the
# business-urgency framing applies to either.
_HU_COFFEE_MACHINE_RE = _re.compile(
    r"káv[eé](gép|automata)|presszógép|eszpresszógép", _re.IGNORECASE)


def _hu_is_coffee_machine_repair(text_data: str) -> bool:
    return bool(_HU_COFFEE_MACHINE_RE.search(text_data or ""))


def _hu_repair_niche_corroborated(niche: str, text_data: str) -> bool:
    """True if the scraped page text itself contains real vocabulary for
    this repair niche (hu_niche.py's own _TEXT_KEYWORDS list), not just an
    Excel label or a name-based guess. Same corroboration idea already used
    for openers (_hu_corroboration_guard below), applied here to the PITCH
    BODY choice instead — a wrong opener is embarrassing, but a wrong
    product pitch (telling a solar installer an AI will answer their
    "repair" calls) is worse, so this gate exists specifically to stop
    _hu_pick_body from trusting a weak repair-niche label.

    The kávégép-szerviz carve-out is NOT in hu_niche.py's own vocabulary (see
    _HU_REPAIR_COPY's comment on why it's a keyword check here rather than a
    real niche), so a coffee-machine lead's text never matches
    "háztartásigép-szerviz"'s own _TEXT_KEYWORDS list and would otherwise be
    wrongly rejected as uncorroborated. The coffee-machine regex is checked
    as an equally valid alternate signal for that one niche.
    """
    t_norm = _norm(text_data)
    if any(_norm(kw) in t_norm for kw, _w in _TEXT_KEYWORDS.get(niche, [])):
        return True
    return niche == "háztartásigép-szerviz" and _hu_is_coffee_machine_repair(text_data)


def _hu_repair_pitch(niche: str, text_data: str, city_hint: str = "") -> str:
    """AI voice-agent pitch for the two repair niches. Deliberately NOT a
    variant of the árkalkulátor pitch — a repair shop cannot price a job
    without seeing the device, so an instant-quote bot doesn't fit here at
    all, unlike klíma/autóüveg-style lookup pricing.

    No website upsell here on explicit instruction: this pitch sells the
    voice agent only, one thing, not a menu of two offers.
    """
    copy_key = ("kávégép-szerviz"
                if niche == "háztartásigép-szerviz" and _hu_is_coffee_machine_repair(text_data)
                else niche)
    copy = _HU_REPAIR_COPY[copy_key]
    pitch = (f"{_HU_REPAIR_INTRO} "
             f"{copy['pain']}\n\n"
             f"{copy['mech']}")
    # Same low-friction close as the árkalkulátor pitch: no cold no-brainer to
    # respond to, just permission to send something already built.
    pitch += "\n\nCsak megmutatnám, mást nem kérek. Elküldhetem?"

    if "miskolc" in city_hint.lower():
        pitch += "\n\nMellesleg én is miskolci vagyok, szóval ha kérdése van, személyesen is megbeszélhetnénk."

    return "\n\n" + pitch


def _hu_pick_body(trade: str, text_data: str, city_hint: str = "", company_name: str = "") -> str:
    # Always prefer the clean Excel name injected into text_data
    # over whatever extract_company_name() scraped from the page
    excel_name = _td_section(text_data, "COMPANY NAME")
    if excel_name:
        company_name = excel_name
    # Was _hu_detect_niche() — i.e. name + scraped-text only, ignoring the
    # sheet's own Niche column, while _hu_opener_postprocess() right below used
    # _resolve_niche() and DID honour it. The two could therefore disagree on
    # the same lead. That mattered little while every niche was a trade (the
    # text scorer had plenty of signal), but a moving company's or a storage
    # firm's Facebook page contains almost no keyword the scorer knows, so
    # without the Excel hint they fall to the default and get pitched as
    # construction firms. Same resolver in both places now.
    niche = _resolve_niche(text_data)

    # Repair niches get a completely different product pitch — see
    # _hu_repair_pitch's docstring for why an árkalkulátor doesn't fit them.
    #
    # But only once the scraped text actually backs it up. _resolve_niche()
    # can land on "háztartásigép-szerviz"/"telefon- és laptopszerviz" from a
    # loose Excel TRADE TYPE label or a name-based guess alone, with nothing
    # in the real page text confirming it — a real case that surfaced this:
    # a Miskolc lead came back with "milyen háztartási géppel van gond" in
    # the pitch despite the business actually being something else (a solar
    # installer, in the case flagged). Sending "an AI answers your repair
    # calls" to a business that does not take repair calls at all is a much
    # worse miss than the árkalkulátor pitch just being the generic default
    # for them, so an uncorroborated repair niche falls through to the
    # regular pitch below instead of trusting the label.
    if niche in _HU_REPAIR_NICHES and _hu_repair_niche_corroborated(niche, text_data):
        return _hu_repair_pitch(niche, text_data, city_hint)

    questions = _HU_NICHE_QUESTIONS.get(niche, _HU_DEFAULT_QUESTIONS)

    # Leads with the finished build instead of "I make X for Y". The old
    # opening announced the product and the closing said "Önnek is összeraktam
    # egyet" — two sentences doing one job, and the reader had to reach the end
    # to learn something already existed for them. Now the first thing after
    # the personalised opener is that it's already built, and the close is a
    # bare "Elküldhetem?" with nothing to restate.
    #
    # company_name is still never inflected into the copy: the Excel COMPANY
    # NAME column often has the city glued on ("Generálkivitelező Kft. pécs"),
    # and Hungarian dative depends on whether the name ends in an abbreviation
    # ("Kft.-nek" but "Kovács Építőiparnak"), so "Készítettem a <cégnév>-nek"
    # would mangle names in the very first line. "Kedves <keresztnév>!" in the
    # greeting already carries the personalisation.
    # "meg szeretném", not "megszeretném": the igekötő is separated from its
    # verb by the auxiliary, so it is written as two words (AkH). The comma
    # before "és" is there because it joins two clauses, not two objects.
    # The explicit "and I'd like to show it to you" states the ask in the
    # second sentence rather than leaving it to the closing "Elküldhetem?" —
    # the reader knows what this message wants before deciding to keep reading.
    intro = (
        "Diák vagyok, mesterséges intelligenciával foglalkozom: készítettem Önnek "
        "egy árkalkulátort, és meg szeretném mutatni."
        if HU_MENTION_STUDENT
        else "Készítettem Önnek egy árkalkulátort, és meg szeretném mutatni."
    )
    # Singular Ön throughout ("Önnek", "Önnél", "tudja") so the whole message
    # addresses one person in one register. The per-niche question lists above
    # need no change — they are phrased from the END CUSTOMER's side and were
    # already impersonal or formal.
    # One line or the other, never a hedge covering both. Saying "ha nincs
    # weboldala" to someone whose site is linked on the page they are reading
    # this from reads as not having looked; saying "lenne rajta mit javítani"
    # to someone with no site is nonsense. The scrape already knows which.
    _reg = _hu_register(niche)
    _copy = _HU_REGISTER_COPY.get(_reg)
    _proof_line = (_HU_ALT_PROOF if _copy
                   else _HU_NICHE_PROOF.get(niche, HU_PROOF_LINE))
    _mech_line = _copy["mech"] if _copy else _HU_TRADE_MECH
    _notify_line = _copy["notify"] if _copy else _HU_TRADE_NOTIFY

    _site = _hu_has_website(text_data)
    # A domain that is merely PRESENT on the page is enough to know they have a
    # site, but not enough to say anything about it.
    _site_theirs = _hu_site_is_theirs(_site, text_data)
    if _site and not _site_theirs:
        print(f"  Site {_site} found but not confirmed theirs — no claim made")
        site_line = "Ha a weboldalán lenne mit javítani, azt is megcsinálom."
    elif _site:
        chk = _hu_site_check(_site)
        if chk["ok"]:
            # Earned, not claimed: the site was really fetched a moment ago.
            gaps = []
            if not chk["quote_path"]:
                gaps.append("nincs rajta azonnali árajánlat")
            if not chk["mobile"]:
                gaps.append("nincs mobilra optimalizálva")
            if chk["secs"] >= 3:
                gaps.append(f"{chk['secs']}s alatt tölt be")
            print(f"  Site {_site}: {chk['secs']}s, quote_path={chk['quote_path']}, "
                  f"mobile={chk['mobile']} — FINDINGS: {'; '.join(gaps) or 'none'}")
            site_line = ("Megnéztem a weboldalát is, van rajta pár dolog, "
                         "amit érdemes lenne javítani.")
        else:
            # Could not load it, so nothing was seen and nothing is claimed.
            print(f"  Site {_site} did not load — no website claim made")
            site_line = "Ha a weboldalán lenne mit javítani, azt is megcsinálom."
    else:
        print("  No website found on page — using the build line")
        site_line = ("Weboldalt nem találtam Önöknél. "
                     "Azt is meg tudom csinálni, az árkalkulátorral együtt.")

    pitch = (f"{intro} "
             f"{_proof_line}\n\n"
             # The "nem fix ár" clause is the important one: a tradesman's
             # first fear is being pinned to a number he didn't quote, and
             # preempting it is worth more than another benefit. The "stb." +
             # "Önre lennének szabva" pair matters too — without it the example
             # questions read as the fixed script rather than as illustrations,
             # and a mover who also does office jobs would discount the whole
             # thing on a question list that doesn't cover them.
             # No em dash (per explicit instruction) — colon + comma list
             # instead of the em-dash-wrapped parenthetical this used to be.
             # "a bot" dropped (the subject carries from the previous
             # sentence), and "Nem fix ár, a végleges ajánlat mindig Önnél
             # marad." collapsed into "nem fix árat": the same claim in two
             # words instead of nine.
             f"Ha valaki árat kér, rákérdez a részletekre: {questions}, stb., "
             f"{_mech_line}"
             f"A kérdések nyilván Önre lennének szabva.\n\n"
             f"{_notify_line}\n\n"
             # DM ONLY. These leads come from Facebook pages and the
             # sheet has no website column, so a large share of them may
             # have no site at all — in which case the whole offer reads
             # as "install this where?". One line turns that from a dead
             # end into the bigger job, and someone answering "nincs is
             # weboldalam" has self-identified as wanting a website.
             # Deliberately NOT added to the email channel: those leads
             # were sourced by website, so all of them already have one.
             f"{site_line}\n\n"
             # The ask, deliberately made small. Standard sales logic closes
             # on the value; this closes on how little is being asked for,
             # because what stalls this funnel is not persuasion (people DO
             # say yes) but the size of the next step. Nothing is being sold
             # here — the only request is permission to send a link, and
             # saying so plainly is both true and the lowest-friction ask
             # available.
             f"Csak megmutatnám, mást nem kérek. Elküldhetem?")

    if "miskolc" in city_hint.lower():
        pitch += "\n\nMellesleg én is miskolci vagyok, szóval ha kérdése van, személyesen is megbeszélhetnénk."

    # "\n\n" not "\n": the caller builds the DM as `opener + " " + body`, so a
    # single newline left the pitch hard against the opener with no blank line
    # between them. tidy_message() rstrips each line and preserves blank ones,
    # so the trailing space on the opener line disappears cleanly.
    return "\n\n" + pitch

dm_app.pick_body = _hu_pick_body


# ── Hungarian follow-up sequence ──────────────────────────────────────────────
# dm_app.py already has a real, working follow-up system — FOLLOWUP_TOUCHES=3,
# a day 3/10/21 cadence tracked automatically via /followups_due and
# /followup_send (dm_app.py), and _followup_text() generating the message per
# touch. It was never overridden here, so the app's automatic follow-up button
# would have sent ENGLISH text to Hungarian leads. This replaces it with the
# same signature (touch, company_name, niche), so the existing routes pick it
# up transparently — same override pattern as pick_body/_apply_greeting above.
# Two explicit decisions from the user, different from the EN version:
# - No repeat of the proof-line stat — already seen once in the cold DM,
#   repeating it verbatim reads as a script rather than a real reply.
# - Touch 3 keeps pushing toward a call rather than the EN version's graceful
#   "no worries, I'll leave it here" exit.
# Repair-niche follow-ups: same 3-touch cadence, but never mention
# "árkalkulátor"/"chatbot" — those two got a different first message
# (_hu_repair_pitch) and a follow-up naming the wrong product would out the
# whole sequence as a script. Same shape as the trade version otherwise:
# touch 1 nudges on the thing already built, touch 2 restates the benefit,
# touch 3 pushes for a short call.
_HU_REPAIR_FOLLOWUP = {
    1: [
        "Jó napot! Csak visszajelzek, nem sürgős. A rendszer, amit Önnek "
        "készítettem, továbbra is áll, csak szóljon, és megmutatom.",
        "Jó napot! Egy gyors emlékeztető csak, nem akarok tolakodni. A rendszer, "
        "ami felveszi a kihagyott hívásokat, megvan Önnek, csak egy szavába kerül.",
        "Jó napot! Nem sürgős, csak visszajelzek. A telefonos rendszer kész van, "
        "csak annyi kell, hogy szóljon, és megmutatom.",
    ],
    2: [
        "Jó napot, még egyszer visszatérek erre. A lényeg, hogy egy kihagyott "
        "hívás se vesszen el szerelés közben, mert a rendszer azonnal felveszi "
        "helyette. Küldjem, hogy lássa élőben?",
        "Jó napot! Csak még egyszer megkérdezem. Nem azért, hogy több hívást "
        "kapjon, hanem hogy egyik se vesszen el, amíg nem ér oda a telefonhoz. "
        "Van most öt perce megnézni?",
    ],
    # Touch 3 — keeps pushing toward a call, per explicit instruction, rather
    # than a graceful exit.
    3: [
        "Jó napot! Utoljára írok emiatt. Ha van öt perce, szívesen megmutatom élőben. "
        "Mikor érne rá egy rövid hívásra?",
        "Jó napot! Még egyszer, utoljára. Nem köt semmire, csak egy gyors bemutató "
        "lenne, hogy lássa, mire gondolok. Ma vagy holnap ráérne öt percre?",
    ],
}


def _hu_followup_text(touch: int, company_name: str = "", niche: str = "") -> str:
    t = max(1, min(touch, dm_app.FOLLOWUP_TOUCHES))
    if niche in _HU_REPAIR_NICHES:
        return random.choice(_HU_REPAIR_FOLLOWUP[t])

    # `niche` used to be turned into a niche_label that nothing ever read.
    # It feeds the touch-2 question list instead now: that variant hardcoded
    # "mekkora a munka, milyen anyag, mikorra kellene", which is the same
    # wrong-trade tell the cold DM had — a moving company or a dentist reading
    # a materials question knows the follow-up is a script.
    questions = _HU_NICHE_QUESTIONS.get(niche, _HU_DEFAULT_QUESTIONS)
    if t == 1:
        # Leads with an already-finished fact (már összeraktam / elkészítettem
        # / van egy kész...), never a conditional ("I'll send you a sample" if
        # you want one). The whole point is removing the friction of them
        # having to ask for something — it already exists, waiting for a yes.
        # Wording kept in step with the cold DM's new opening ("készítettem
        # Önnek egy árkalkulátor chatbotot, és meg szeretném mutatni") — the
        # nudge should read as the same person continuing the same thought,
        # not a different script. "A chatbot" / "az árkalkulátor", never the
        # vaguer "minta", so touch 1 still names the actual thing.
        return random.choice([
            "Jó napot! Csak visszajelzek, nem sürgős. A chatbot, amit Önnek készítettem, továbbra "
            "is áll, csak szóljon, és átküldöm.",
            "Jó napot! Egy gyors emlékeztető csak, nem akarok tolakodni. Az árkalkulátor megvan "
            "Önnek, csak egy szavába kerül, és küldöm is.",
            "Jó napot! Nem sürgős, csak visszajelzek. A chatbot kész van, egy élő linken meg tudom "
            "mutatni, csak annyi kell, hogy szóljon.",
            "Jó napot! Csak egy gyors emlékeztető. Az árkalkulátor, amit Önnek raktam össze, kész, "
            "csak szólnia kell, és elküldöm.",
        ])
    if t == 2:
        return random.choice([
            "Jó napot, még egyszer visszatérek erre. Pár kérdést tesz fel az érdeklődőnek, "
            f"{questions}, megmondja az alapárat, és azok jutnak el Önig, "
            "akik tényleg meg akarják csináltatni. Küldjem a mintát, hogy lássa élőben?",
            "Jó napot! Csak még egyszer megkérdezem. A lényeg, hogy ne több érdeklődőt kapjon, "
            "hanem jobbakat, olyanokat, akik tényleg meg akarják csináltatni. Van most öt perce megnézni?",
            "Jó napot, még egyszer visszatérek erre. Ön egy értesítést kap a nevével és a számával, "
            "mielőtt felhívná, így azokhoz jut el, akik tényleg meg akarják csináltatni. Küldjem a "
            "mintát, hogy lássa élőben?",
        ])
    # Touch 3 — keeps pushing toward a call, per explicit instruction, rather
    # than the EN version's graceful exit.
    return random.choice([
        "Jó napot! Utoljára írok emiatt. Ha van öt perce, szívesen megmutatom élőben, hogyan nézne ki "
        "Önöknél, és utána eldöntheti, hogy érdekli-e. Mikor érne rá egy rövid hívásra?",
        "Jó napot! Még egyszer, utoljára. Nem köt semmire, csak egy gyors bemutató lenne, hogy lássa, "
        "mire gondolok. Ma vagy holnap ráérne öt percre?",
        "Jó napot! Utoljára írok emiatt. Ha van öt perce, szívesen megmutatom élőben, hogyan nézne ki "
        "Önöknél, a munkamenetében semmit nem változtat, csak abban segít, hogy a jó ügyfelekhez "
        "jusson el. "
        "Mikor érne rá egy rövid hívásra?",
    ])


dm_app._followup_text = _hu_followup_text

# -- Hungarian greeting prefix (replaces "Cheers") ----------------------------
# Dead in practice — dm_app.py only DEFINES GREETING_PREFIX (line 37) and never
# reads it; the real greeting comes from GREETING_NAME_FMT/GREETING_SOLO_FMT
# via _hu_apply_greeting above. Kept in the magázás register anyway so it can't
# reintroduce tegezés if something ever starts consuming it.
dm_app.GREETING_PREFIX = "Jó napot"

# -- Hungarian name order: family name first, given name second ---------------
_orig_extract = dm_app.extract_first_name
def _hu_extract_first_name(text: str, company_name: str = "") -> str:
    name = _orig_extract(text, company_name)
    if not name:
        return ""
    parts = name.strip().split()
    # If two words detected, take the second (given name in Hungarian order)
    return parts[-1] if len(parts) >= 2 else parts[0]
dm_app.extract_first_name = _hu_extract_first_name

# ── City ragos form helper ────────────────────────────────────────────────────
_CITY_RAGOS: dict[str, str] = {
    # Budapest and major cities
    "Budapest": "Budapesten", "Miskolc": "Miskolcon", "Pécs": "Pécsen",
    "Debrecen": "Debrecenben", "Győr": "Győrben", "Szeged": "Szegeden",
    "Nyíregyháza": "Nyíregyházán", "Kecskemét": "Kecskeméten",
    "Székesfehérvár": "Székesfehérváron", "Eger": "Egerben",
    "Szombathely": "Szombathelyen", "Sopron": "Sopronban",
    "Kaposvár": "Kaposváron", "Veszprém": "Veszprémben",
    "Szolnok": "Szolnokon", "Tatabánya": "Tatabányán",
    "Zalaegerszeg": "Zalaegerszegen", "Dunaújváros": "Dunaújvárosban",
    "Érd": "Érden", "Hódmezővásárhely": "Hódmezővásárhelyen",
    "Nagykanizsa": "Nagykanizsán", "Győrszentiván": "Győrszentivánon",
    "Budaörs": "Budaörsön", "Gödöllő": "Gödöllőn",
    "Gyöngyös": "Gyöngyösön", "Esztergom": "Esztergomban",
    "Ajka": "Ajkán", "Oroszlány": "Oroszlányban",
    "Pápa": "Pápán", "Mosonmagyaróvár": "Mosonmagyaróváron",
    "Győr-Moson-Sopron": "Győr-Moson-Sopron megyében",
    "Baja": "Baján", "Cegléd": "Cegléden", "Szentendre": "Szentendrén",
    "Vác": "Vácon", "Dunakeszi": "Dunakeszin", "Szigetszentmiklós": "Szigetszentmiklóson",
    "Pomáz": "Pomázon", "Vecsés": "Vecsésen", "Monor": "Monoron",
    "Gyula": "Gyulán", "Békéscsaba": "Békéscsabán", "Orosháza": "Orosházán",
    "Makó": "Makón", "Szentes": "Szentesen",
}

def _city_to_ragos(city: str) -> str:
    """Return the locative (ragos) form of a Hungarian city name, or the raw name if unknown."""
    return _CITY_RAGOS.get(city.strip(), city.strip())


def _td_section(text_data: str, header: str) -> str:
    """Pull the first non-empty line after a --- HEADER --- marker."""
    marker = f"--- {header} ---"
    for line in text_data.split("\n"):
        if line.startswith(marker):
            idx = text_data.find(line) + len(line)
            rest = text_data[idx:].strip().split("\n")
            return rest[0].strip() if rest else ""
    return ""


# dm_app.py injects the lead sheet's own "Niche" column into text_data as
# "--- TRADE TYPE ---" for BOTH apps (dm_app.py:3392/3843, unconditional on
# lead["niche"]) — the EN app reads it (_en_prompt_builder / _resolve_niche in
# dm_app.py all check TRADE TYPE first), but this HU wrapper never did, so
# hand-curated ground truth from hun_leads_new_100.xlsx ("Roofing", "Bathroom",
# "Construction", "Renovation", "Boiler/HVAC") sat unused while every lead was
# re-derived from noisy scraped Facebook text instead.
# "Construction"/"Renovation" map straight to the default niche — that's
# literally what those labels mean, not a fallback. "Boiler/HVAC" is
# deliberately NOT a direct mapping: it spans two niches (klíma /
# fűtésszerelő) whose _NICHE_VOCAB contradiction lists actively conflict
# (klíma contradicts "kazán", fűtésszerelő contradicts "daikin") — forcing the
# wrong half would make the opener validator reject correct openers. Instead
# it narrows the text scorer to just those two candidates rather than the
# full niche set, and only picks one if there's real keyword evidence.
_EXCEL_NICHE_DIRECT_MAP = {
    "roofing": "tetőfedő",
    "bathroom": "fürdőszoba felújító",
    "construction": "felújító kivitelező",
    "renovation": "felújító kivitelező",
    # Newer niches. Several spellings each because the sheet's Niche column is
    # hand-typed — an unrecognised label silently falls through to text scoring,
    # which for a non-trade niche means the default "felújító kivitelező" and a
    # message that pitches a moving company as a construction firm.
    "moving": "költöztető",
    "removals": "költöztető",
    "moving/removals": "költöztető",
    "dental": "fogászat",
    "dentist": "fogászat",
    "dental (implants/veneers)": "fogászat",
    "implants": "fogászat",
    "solar": "napelemes",
    "solar installers": "napelemes",
    "storage": "raktárbérlő",
    "warehouse": "raktárbérlő",
    "storage/warehouse rental": "raktárbérlő",
    # Lookup-table pricing niches.
    "windscreen": "autóüveg",
    "windshield": "autóüveg",
    "windscreen repair": "autóüveg",
    "windscreen repair/replacement": "autóüveg",
    "auto glass": "autóüveg",
    "appliance repair": "háztartásigép-szerviz",
    "appliance": "háztartásigép-szerviz",
    "white goods": "háztartásigép-szerviz",
    "phone repair": "telefon- és laptopszerviz",
    "laptop repair": "telefon- és laptopszerviz",
    "phone/laptop repair": "telefon- és laptopszerviz",
    "device repair": "telefon- és laptopszerviz",
    "translation": "fordítóiroda",
    "translation/interpretation": "fordítóiroda",
    "translation agency": "fordítóiroda",
    "interpretation": "fordítóiroda",
    "car detailing": "autókozmetika",
    "detailing": "autókozmetika",
    "valeting": "autókozmetika",
    "car detailing/valeting": "autókozmetika",
    "hair transplant": "hajbeültetés",
    "hair transplants": "hajbeültetés",
}

# Substring fallback, tried in order when the exact map above misses. Every
# sheet is hand-labelled fresh, so the wording drifts — "Storage/warehouse
# rental" one week, "Storage unit rental" the next, "Hair transplants" then
# "Hair transplant clinic". Exact-match alone silently dropped 13 of 100 leads
# on hungary_leads 08 14 to the default niche, which means a storage company
# pitched as a builder. Matching on the distinctive word instead makes that
# whole class of failure go away without another edit here every time.
# ORDER MATTERS: first hit wins, so anything that is a substring of another
# label must come after it.
_EXCEL_NICHE_FUZZY = [
    ("hair transplant", "hajbeültetés"),
    ("hajbeültet",      "hajbeültetés"),
    ("windscreen",      "autóüveg"),
    ("windshield",      "autóüveg"),
    ("auto glass",      "autóüveg"),
    ("autóüveg",        "autóüveg"),
    ("appliance",       "háztartásigép-szerviz"),
    ("white goods",     "háztartásigép-szerviz"),
    ("phone",           "telefon- és laptopszerviz"),
    ("laptop",          "telefon- és laptopszerviz"),
    ("device repair",   "telefon- és laptopszerviz"),
    ("translation",     "fordítóiroda"),
    ("interpret",       "fordítóiroda"),
    ("detailing",       "autókozmetika"),
    ("valeting",        "autókozmetika"),
    ("storage",         "raktárbérlő"),
    ("warehouse",       "raktárbérlő"),
    ("moving",          "költöztető"),
    ("removal",         "költöztető"),
    ("dental",          "fogászat"),
    ("dentist",         "fogászat"),
    ("implant",         "fogászat"),
    ("solar",           "napelemes"),
    ("roof",            "tetőfedő"),
    ("bathroom",        "fürdőszoba felújító"),
    # Trade labels. "boiler"/"heating" MUST stay ahead of "hvac": a label like
    # "Gas boiler installation/service" contains neither word exclusively, and
    # heating is the safer read when both could apply. A label naming BOTH is
    # handled in _excel_niche() instead — see there.
    # Hungarian-language sheet labels. hungary_leads_3niches 08 25 arrived with
    # "Autóüveg / szélvédő", "Autókozmetika / detailing" and "Háztartási gép
    # szervíz" — the first two happened to hit existing needles, the third hit
    # nothing, so three appliance shops fell through to the company-name rule
    # and came back as klíma (their names mention hűtő/klíma work). A curated
    # sheet label should always outrank a guess from the business name.
    # Both accented and unaccented spellings: _excel_niche lowercases but does
    # not strip accents, and "szervíz" is a very common misspelling of
    # "szerviz" on these sheets.
    ("háztartási gép",  "háztartásigép-szerviz"),
    ("haztartasi gep",  "háztartásigép-szerviz"),
    ("gépszervíz",      "háztartásigép-szerviz"),
    ("gépszerviz",      "háztartásigép-szerviz"),
    ("gepszerviz",      "háztartásigép-szerviz"),
    ("szélvédő",        "autóüveg"),
    ("szelvedo",        "autóüveg"),
    ("autókozmetika",   "autókozmetika"),
    ("autokozmetika",   "autókozmetika"),
    ("landscap",        "kertépítő"),
    ("garden",          "kertépítő"),
    ("kertépít",        "kertépítő"),
    ("boiler",          "fűtésszerelő"),
    ("heating",         "fűtésszerelő"),
    ("hvac",            "klíma"),
    ("air con",         "klíma"),
    ("klíma",           "klíma"),
    ("plumb",           "vízvezetékszerelő"),
    ("electric",        "villanyszerelő"),
    ("paint",           "festő-mázoló"),
    ("tiling",          "burkoló"),
    ("tiler",           "burkoló"),
    ("window",          "nyílászáró"),
]


def _excel_niche(label: str) -> str:
    """Map a hand-typed sheet label to a niche. Exact first, then substring."""
    lab = (label or "").strip().lower()
    if not lab:
        return ""
    # A label naming BOTH boiler and HVAC is genuinely ambiguous: klíma and
    # fűtésszerelő have contradiction lists that actively exclude each other
    # ("kazán" disqualifies klíma, "daikin" disqualifies fűtésszerelő), so
    # picking one here would pre-empt the two-candidate text scoring in
    # _resolve_niche and could force the opener validator to reject correct
    # openers. Stay out of its way and let it decide on real evidence.
    if "boiler" in lab and "hvac" in lab:
        return ""
    if lab in _EXCEL_NICHE_DIRECT_MAP:
        return _EXCEL_NICHE_DIRECT_MAP[lab]
    for needle, niche in _EXCEL_NICHE_FUZZY:
        if needle in lab:
            return niche
    return ""
_HVAC_CANDIDATE_NICHES = ("klíma", "fűtésszerelő")


def _resolve_niche(text_data: str) -> str:
    """
    Name-first classifier (most reliable) → Excel TRADE TYPE hint → full text
    scoring as a last resort.
    """
    company_name = _td_section(text_data, "COMPANY NAME")
    excel_niche = _td_section(text_data, "TRADE TYPE").strip().lower()

    # A SPECIFIC sheet label now outranks the name heuristic. It's hand-curated
    # per sheet, whereas the name rule is a guess from a string — and on a
    # company whose name spans two niches the guess loses. Real case on
    # hungary_leads 08 14: "ZedStore Raktározás és Költöztetés" is labelled
    # "Storage unit rental", but the name rule fires on "költöztetés" and
    # pitched it as a moving company with moving questions.
    # Catch-all labels are deliberately excluded: "Construction"/"Renovation"
    # resolve to the DEFAULT niche and carry no information, so a specific name
    # like "Kovács Tetőfedő" must still beat them.
    _from_label = _excel_niche(excel_niche)
    if _from_label and _from_label != "felújító kivitelező":
        return _from_label

    # Name next: still the best signal when the label is generic or missing.
    if company_name:
        niche = _hu_detect_niche_from_name(company_name)
        if niche:
            return niche

    if excel_niche == "boiler/hvac":
        t_norm = _norm(text_data)
        scores = {
            n: sum(w for kw, w in _TEXT_KEYWORDS.get(n, []) if _norm(kw) in t_norm)
            for n in _HVAC_CANDIDATE_NICHES
        }
        scores = {n: s for n, s in scores.items() if s > 0}
        if scores:
            return max(scores, key=scores.get)
        # No corroborating keyword for either half — don't guess, fall through
    elif excel_niche in _EXCEL_NICHE_DIRECT_MAP:
        return _EXCEL_NICHE_DIRECT_MAP[excel_niche]

    # Last resort: full detector (text scoring) — ignores stale/unknown Excel labels
    return _hu_detect_niche(text_data, company_name)


# ── Hungarian praise phrases ──────────────────────────────────────────────────
# call_gemini() pins one _OPENER_LEADS phrase and one _OPENER_DETAILS word per
# call and appends "FOR THIS ONE: use '<phrase>' as the praise phrase" to the
# very end of the prompt — AFTER CUSTOM_PROMPT_BUILDER runs, so the builder
# can't prevent it. The stock lists are English ("looks cracking", "proper
# tidy", "flush"), and being the last and most specific instruction they beat
# the Hungarian PROMPT, producing half-English openers like "az a tetőfedés
# Szegeden looks cracking". Override both lists with Hungarian equivalents.
# Leads are verb phrases, details are adjectives that open the second clause,
# so they can't collide the way the English lists could.
# Vocabulary taken from real Hungarian tradesman reviews (joszaki, tudjukki,
# korrektmester), where "precíz" and "igényes" dominate and praise attaches to
# a named detail ("szép vágások, szép illesztések"). Deliberately NOT
# "klassz/szuper/komolyan néz ki" — that register reads as a kid writing to a
# 50-year-old kőműves, and "komolyan néz ki" is a calque of "looks proper"
# rather than anything a Hungarian says.
# Every lead must be a predicate that can follow "Az a [munka]" directly, and
# must avoid the passive "lett megcsinálva/kivitelezve" forms the PROMPT bans.
# "rendesen néz ki" was dropped: "rendes" is faint praise in Hungarian
# ("decent"), and lukewarm praise from a stranger lands worse than none.
# "gyönyörű"/"nagyon szép" are what the reviews actually use.
dm_app._OPENER_LEADS = [
    "gyönyörű munka", "nagyon szép munka", "nagyon igényes munka",
    "precíz munka", "gondos munka", "kifogástalan lett",
    "nagyon szépen néz ki", "nagyon jól néz ki",
]
# Adjectives opening the second clause. Stems here must not appear in the
# leads above, or the line doubles up ("precíz munka, precíz a szegélyezés").
dm_app._OPENER_DETAILS = [
    "egyenletes", "egyforma", "pontos", "tiszta",
    "szabályos", "egyenes", "arányos",
]


# ── Niche-aware prompt builder ────────────────────────────────────────────────
def _hu_prompt_builder(full_prompt: str, text_data: str) -> str:
    """Inject per-niche context before the vision call."""
    niche = _resolve_niche(text_data)
    instr = NICHE_OPENER_INSTRUCTIONS.get(niche)
    if instr:
        print(f"  Prompt builder: niche='{niche}' — injecting instructions")
        return full_prompt + f"\n\n{instr}"
    return full_prompt

dm_app.CUSTOM_PROMPT_BUILDER = _hu_prompt_builder


# ── Corroboration guard: is the specific detail actually backed by anything? ──
# Originally scoped to the default niche ("felújító kivitelező") only, on the
# reasoning that _opener_matches_niche()'s CONTRADICTION check already covers
# niches with a _NICHE_VOCAB entry (klíma, tetőfedő, burkoló, ...). But
# contradiction-checking only catches a claim that clashes with a DIFFERENT
# niche's vocabulary (e.g. "cserép" showing up on a klíma company) — it does
# nothing for a claim that's simply invented but doesn't contradict anything,
# e.g. a correctly-classified tetőfedő company getting "Bramac" invented when
# nothing on their page ever mentions Bramac. That left 18 of 19 niches with
# zero protection against exactly this — only the noisiest bucket had a
# corroboration check. Extended to run for every niche, since the actual risk
# (the model inventing a specific, checkable detail) exists everywhere the
# model makes a specific claim, not just in the default bucket.
# Real tradeoff, taken deliberately: a vision model can sometimes correctly
# read a brand directly off the photo even when the page's caption text never
# mentions it — that legitimate case gets rejected too under this rule. But
# two real, confirmed rejections in this conversation ("nem bízom az
# ilyesmiben" and "nem fűtésszereléssel foglalkozunk...") both traced back to
# a wrong specific claim, and the 2026 cold-outreach research is explicit that
# wrong personalisation converts worse than generic — erring toward more
# fallbacks and fewer wrong claims is the right side to be wrong on.
# The _NICHE_VOCAB-derived set alone missed most real brand names: "daikin"
# and "vaillant" happened to be in it only because they're cross-referenced as
# disqualifiers in OTHER niches' contradict lists (klíma/fűtésszerelő), but
# "tondach", "eternit", "creaton", "icopal", "fakro" — the other real roofing
# brands sitting right there in tetőfedő's own _TEXT_KEYWORDS — were entirely
# absent, since nothing cross-references them. Verified live: 'tondach' in
# _SPECIFIC_TRADE_TERMS was False before this fix. _TEXT_KEYWORDS' own weights
# (comment: "4 = uniquely niche-specific, 3 = very specific") are exactly the
# right signal for "specific enough to need corroboration" — pulling every
# weight >= 3 term from every niche gives real, comprehensive brand coverage
# instead of an incidental byproduct of a different mechanism.
_SPECIFIC_TRADE_TERMS = sorted({
    _norm(term)
    for vocab in _NICHE_VOCAB.values()
    for term in vocab.get("contradict", [])
} | {
    _norm(kw) for kw_list in _TEXT_KEYWORDS.values() for kw, weight in kw_list if weight >= 3
} | {_norm(t) for t in [
    "tetőfedés", "tetőfedő", "klímaszerelés", "kazánszerelés",
    "fűtésszerelés", "napelemszerelés", "térkövezés", "burkolás",
]})


def _hu_corroboration_guard(opener: str, text_data: str) -> bool:
    o, td = _norm(opener), _norm(text_data)
    for term in _SPECIFIC_TRADE_TERMS:
        if term in o and term not in td:
            print(f"  Opener REJECTED (uncorroborated '{term}'): '{opener[:60]}'")
            return False
    return True


# ── Attribution guard: is this actually the company's OWN job? ───────────────
# Different failure mode from the two guards above — those catch the WRONG
# TRADE; this catches the RIGHT trade attributed to the WRONG COMPANY. In a
# market this small, the owner remembers every real job they've done, so
# "I've seen that job you did — I love the X" lands as "no, I never did that"
# and kills the reply outright. This happens because Facebook renders a
# reshared supplier post, a partner's tagged photo, or a customer's own
# review-with-photo as an ordinary post on the company's page, and
# _scroll_to_freshest_post (dm_app.py) has no way to tell those apart from an
# original completed job — it just targets the first [role="article"] it
# finds. The PROMPT above now tells the model to check for this, but a model
# instruction alone isn't reliable (same reason the niche guards above exist
# as a backstop rather than trusting the prompt) — this is the same backstop
# for attribution specifically, checked against the scraped page text.
# False positives here (skipping personalisation when the post WAS actually
# theirs) cost a generic opener. False negatives cost a wrong, checkable claim
# to someone who'd know — confirmed worse than generic in 2026 cold-outreach
# research, so this deliberately ers toward rejecting.
_RESHARE_SIGNALS = [
    "megosztotta", "megosztott egy", "megosztás:",
    "közösen ", "megjelölve", "veled közösen",
    "értékelte:", "véleményt írt", "ajánlotta:",
]


def _hu_looks_like_reshare_or_review(text_data: str) -> bool:
    t = _norm(text_data)
    return any(_norm(sig) in t for sig in _RESHARE_SIGNALS)


# ── City guard: does the opener name the lead's actual city? ──────────────────
# Nothing in the vision PROMPT ever tells the model to source its city mention
# from LISTED CITY — that value is only referenced in the FALLBACK-generation
# code (below), never fed to the model as "this is the city, use this." So a
# successful (non-fallback) opener's city comes entirely from whatever the
# model reads or infers off the screenshot, with zero check against the real,
# spreadsheet-sourced location for that lead. Same failure class as the brand
# guard above, just for location instead of material: a company based in one
# county, pitched with a job "in" a city in a different county, is an
# instantly checkable wrong claim — this is plausibly part of what's behind
# the "csak Borsodban" rejection two turns ago (Debrecen named in the opener,
# but the company's real service area doesn't include it).
# Conservative by design: only rejects when the opener names a DIFFERENT
# recognized city than LISTED CITY. An opener that names no city, or an
# unrecognized one, is left alone rather than guessed at.
def _hu_city_guard(opener: str, text_data: str) -> bool:
    listed_city = _td_section(text_data, "LISTED CITY").strip()
    if not listed_city:
        return True
    o = _norm(opener)
    if _norm(listed_city) in o or _norm(_city_to_ragos(listed_city)) in o:
        return True
    for city, ragos in _CITY_RAGOS.items():
        if _norm(city) == _norm(listed_city):
            continue
        if _norm(city) in o or _norm(ragos) in o:
            print(f"  Opener REJECTED (wrong city '{city}', lead is in '{listed_city}'): '{opener[:60]}'")
            return False
    return True


# ── Text-sourced openers for the non-trade niches ─────────────────────────────
# The entire vision opener premise is "look at a photo of a FINISHED job and
# name a craft detail" — fugaköz, korcok távolsága, tokreveal. That premise has
# no equivalent here:
#   költöztető  — the photo is a van, boxes or a lifting crew. There is no
#                 workmanship to admire; "szép a furgonod" is a joke.
#   raktárbérlő — the photo is an empty unit, a gate or a fence. Same problem.
#   fogászat    — the photo is very often a real patient's face or teeth.
#                 Complimenting those in a cold DM is creepy, and it is the one
#                 failure here that would actively damage him.
# Left alone, all three would burn a vision call and then get kicked to the
# generic city fallback by the guards anyway, since nothing the model invented
# could be corroborated.
#
# So for these niches the opener is sourced from the page TEXT instead of the
# photo, and built deterministically rather than by a model. The rule is that a
# phrase is only ever emitted if its trigger word literally appears in the
# scraped text — which makes the claim unfalsifiable by construction, the same
# property _hu_corroboration_guard tries to enforce after the fact for trades.
# Deliberately NOT hooked on the generic trade word itself ("költöztetés" for a
# company called "X Költöztetés"): that produces a line that says nothing.
# Only distinctive, self-selected services count.
_HU_TEXT_OPENER_HOOKS: dict[str, list[tuple[str, str]]] = {
    # Frame: "Láttam, hogy {phrase} Győrben." — phrase is a full predicate,
    # in magázás (E/3, "Önök") to match the rest of the message.
    "költöztető": [
        ("zongoraszállítás",  "zongoraszállítást is vállalnak"),
        ("zongora",           "zongorát is szállítanak"),
        ("páncélszekrény",    "páncélszekrényt is szállítanak"),
        ("irodaköltöztetés",  "irodaköltöztetést is vállalnak"),
        ("iroda költöztetés", "irodaköltöztetést is vállalnak"),
        ("nemzetközi",        "nemzetközi költöztetést is vállalnak"),
        ("külföldi",          "külföldre is költöztetnek"),
        ("lomtalanítás",      "lomtalanítást is vállalnak"),
        ("bútorszerelés",     "a bútorok szét- és összeszerelését is vállalják"),
        ("csomagolóanyag",    "csomagolóanyagot is adnak"),
        ("csomagolás",        "a csomagolást is elvállalják"),
        ("raktározás",        "raktározást is vállalnak"),
        # No ("zongorát", ...) entry: "zongora" above is a substring of it, so
        # any text containing "zongorát" already matched and this could never
        # fire. Same reason there is no plain ("kamera", ...) in raktárbérlő.
    ],
    "raktárbérlő": [
        ("fűtött",           "fűtött raktárt is kiadnak"),
        ("önkiszolgáló",     "önkiszolgáló tárolóik is vannak"),
        ("0-24",             "0-24 be lehet menni a tárolókhoz"),
        ("24 órás",          "0-24 be lehet menni a tárolókhoz"),
        ("éjjel-nappal",     "éjjel-nappal hozzáférhetők a tárolók"),
        ("kamerás",          "kamerával őrzött a terület"),
        ("kamerával",        "kamerával őrzött a terület"),
        ("konténer",         "konténert is adnak bérbe"),
        ("dobozt",           "dobozokat is lehet Önöknél venni"),
        ("rakodólap",        "raklapos tárolást is vállalnak"),
        ("raklap",           "raklapos tárolást is vállalnak"),
        ("logisztik",        "logisztikai szolgáltatást is adnak"),
    ],
    # Dental was originally left out of this map on purpose, so every dental
    # lead fell through to the bare city line. The reason was register: a
    # formal opener would have clashed with a body that was still tegező. The
    # body is magázás now, so that objection is gone and these can run.
    # The photo ban is untouched and permanent — every phrase here comes from
    # the practice's own page TEXT (services they advertise), never from a
    # picture, so there is no route to commenting on a patient's face, teeth
    # or smile. Phrased E/3 formal, like the other two.
    "fogászat": [
        ("all-on-4",            "All-on-4 implantációval is foglalkoznak"),
        ("all on 4",            "All-on-4 implantációval is foglalkoznak"),
        ("azonnali terhelés",   "azonnali terhelésű implantátumokat is beültetnek"),
        ("szájsebész",          "szájsebészeti beavatkozásokat is végeznek"),
        ("fogtechnikai labor",  "saját fogtechnikai laborral dolgoznak"),
        ("saját labor",         "saját laborral dolgoznak"),
        ("altatás",             "altatásos kezelést is vállalnak"),
        ("cerec",               "CEREC-kel dolgoznak"),
        ("digitális lenyomat",  "digitális lenyomatvétellel dolgoznak"),
        ("cbct",                "saját CBCT-vel dolgoznak"),
        ("3d ct",               "saját 3D CT-vel dolgoznak"),
        ("invisalign",          "Invisalign fogszabályozást is csinálnak"),
        ("fogszabályoz",        "fogszabályozással is foglalkoznak"),
        ("parodont",            "parodontológiával is foglalkoznak"),
        ("gyermekfogász",       "gyermekfogászattal is foglalkoznak"),
        ("porcelán héj",        "porcelán héjakat is készítenek"),
        ("implantát",           "implantátumokkal is foglalkoznak"),
    ],
    # ── Lookup-table pricing niches ──────────────────────────────────────────
    # Same contract as above: the trigger word must literally appear in the
    # scraped page text, so the claim can't be wrong. Ordered most-distinctive
    # first, since the first match wins and a specific service reads far
    # better than a generic one.
    "autóüveg": [
        # Every phrase must end somewhere a place adverbial can follow, because
        # the frame appends " Győrben." — a trailing contrast clause ("..., nem
        # csak cserélnek") or a postposition ("... a csere után") strands the
        # city mid-thought.
        ("kőfelverődés",     "kőfelverődést is javítanak"),
        ("kofelverodes",     "kőfelverődést is javítanak"),
        ("szélvédőjavítás",  "szélvédőt javítanak is, nem csak cserélnek"),
        ("szélvédő javítás", "szélvédőt javítanak is, nem csak cserélnek"),
        ("szelvedojavitas",  "szélvédőt javítanak is, nem csak cserélnek"),
        ("repedés",          "a repedt szélvédőt is meg tudják javítani"),
        ("repedt",           "a repedt szélvédőt is meg tudják javítani"),
        ("kamerakalibrál",   "ADAS-kamerakalibrálást is végeznek"),
        ("adas",             "ADAS-kalibrálást is végeznek"),
        ("mobil szerviz",    "helyszínre is kimennek javítani"),
        ("helyszíni",        "helyszínre is kimennek javítani"),
        ("casco",            "a cascós ügyintézést is átvállalják"),
        ("kaszkó",           "a kaszkós ügyintézést is átvállalják"),
        ("biztosító",        "a biztosítós ügyintézést is átvállalják"),
        ("panorámatető",     "panorámatetőt is cserélnek"),
        ("oldalüveg",        "oldalüveget és hátsó szélvédőt is cserélnek"),
    ],
    "háztartásigép-szerviz": [
        ("beépíthető",       "beépíthető gépekhez is kiszállnak"),
        ("helyszíni",        "helyszínen is javítanak"),
        ("alkatrész",        "alkatrészt is rendelnek a javításhoz"),
        ("szárítógép",       "szárítógépet is javítanak"),
        ("mosogatógép",      "mosogatógépet is javítanak"),
        ("hűtő",             "hűtőket is javítanak"),
        ("mosógép",          "mosógépeket is javítanak"),
        ("garancia",         "garanciát is adnak a javításra"),
    ],
    "telefon- és laptopszerviz": [
        ("adatmentés",       "adatmentést is vállalnak"),
        ("alaplap",          "alaplapi javítást is vállalnak"),
        ("ázott",            "ázott készülékeket is megmentenek"),
        ("vízkár",           "vízkáros készülékeket is megmentenek"),
        ("macbook",          "MacBookot is javítanak"),
        ("iphone",           "iPhone-t is javítanak"),
        ("azonnal",          "van, amit azonnal, helyben megjavítanak"),
        ("kijelzőcsere",     "kijelzőcserét is csinálnak"),
        ("akkumulátorcsere", "akkumulátorcserét is vállalnak"),
    ],
    "fordítóiroda": [
        ("apostille",        "apostille-os ügyintézést is vállalnak"),
        ("hiteles",          "hiteles fordítást is vállalnak"),
        ("tolmácsolás",      "tolmácsolást is vállalnak"),
        ("lektorálás",       "lektorálást is vállalnak"),
        ("honlapfordítás",   "honlapfordítást is vállalnak"),
        ("jogi",             "jogi szakfordítást is vállalnak"),
        ("műszaki",          "műszaki szakfordítást is vállalnak"),
        ("orvosi",           "orvosi szakfordítást is vállalnak"),
        ("szakfordítás",     "szakfordítást is vállalnak"),
    ],
    "autókozmetika": [
        ("kerámia",          "kerámiabevonatot is csinálnak"),
        ("polírozás",        "polírozást is vállalnak"),
        ("ózon",             "ózonos fertőtlenítést is csinálnak"),
        ("kárpittisztítás",  "kárpittisztítást is vállalnak"),
        ("fóliáz",           "fóliázást is vállalnak"),
        ("motormosás",       "motormosást is vállalnak"),
        ("viasz",            "viaszolást is csinálnak"),
        ("belső tisztítás",  "belső tisztítást is vállalnak"),
    ],
    "hajbeültetés": [
        ("dhi",              "DHI technikával is dolgoznak"),
        ("fue",              "FUE technikával dolgoznak"),
        ("szakállbeültetés", "szakállbeültetést is végeznek"),
        ("szemöldök",        "szemöldökbeültetést is végeznek"),
        ("prp",              "PRP-kezelést is adnak"),
        ("női hajbeültetés", "női hajbeültetéssel is foglalkoznak"),
        ("ingyenes konzultáció", "ingyenes konzultációt is adnak"),
    ],
}

# Niches whose opener must never come from the vision model, regardless of what
# it returned. Anything not in _HU_TEXT_OPENER_HOOKS falls to the city line.
_HU_NON_PHOTO_NICHES = {
    "költöztető", "raktárbérlő", "fogászat",
    # Lookup-table niches. None has a finished craft job to photograph, and
    # two have photos that must never be described: hajbeültetés pages are
    # before/after shots of a real patient's scalp (same rule as fogászat),
    # and autóüveg/autókozmetika photos are customers' cars, number plates
    # included. Being in this set also means they skip the screenshots and
    # the vision call entirely — see _hu_skip_vision below.
    "autóüveg", "háztartásigép-szerviz", "telefon- és laptopszerviz",
    "fordítóiroda", "autókozmetika", "hajbeültetés",
}


def _hu_skip_vision(text_data: str) -> bool:
    """True when this lead's opener comes from page text, so dm_app.py can
    skip the screenshots and the vision call entirely.

    Until this hook existed, these leads still paid the full visual pipeline
    (scroll to freshest post, screenshot, vision call at 12s per attempt and
    up to 26s with the retry, then a second screenshot) and then
    _hu_opener_postprocess threw the model's answer away, because a photo of a
    van or a waiting room can't produce a corroboratable claim. On the current
    Hungarian sheet that is 38 of 47 leads, so it's most of the run.
    The opener stays personalised via _hu_text_opener — that's pure string
    matching over text already scraped, so it costs nothing and adds no wait.
    """
    return _resolve_niche(text_data) in _HU_NON_PHOTO_NICHES


dm_app.SKIP_VISION_FOR = _hu_skip_vision


def _hu_city_fallback(text_data: str) -> str:
    city = _td_section(text_data, "LISTED CITY")
    # Magázás: "oldalatokra" (tegező plural) → "oldalukra" (Önök). NOTE:
    # hu_email_gen.py detects this fallback by string prefix to know the AI
    # opener was rejected — it matches on "Rátaláltam" only, so both the old
    # and new wording are caught. Don't change the leading word without
    # checking there.
    return (f"Rátaláltam a Facebook-oldalukra {_city_to_ragos(city)}."
            if city else "Rátaláltam a Facebook-oldalukra.")


def _hu_text_opener(niche: str, text_data: str) -> str:
    """Build an opener from a distinctive service named in the page text.
    Returns "" when nothing specific is found — the caller then uses the safe
    city line rather than reaching for something vaguer."""
    hooks = _HU_TEXT_OPENER_HOOKS.get(niche)
    if not hooks:
        return ""
    t = _norm(text_data)
    city = _td_section(text_data, "LISTED CITY")
    for needle, phrase in hooks:
        if _norm(needle) in t:
            where = f" {_city_to_ragos(city)}" if city else ""
            return f"Láttam, hogy {phrase}{where}."
    return ""


# ── Niche-aware opener post-validator ─────────────────────────────────────────
def _hu_opener_postprocess(opener: str, text_data: str) -> str:
    """Reject opener if it contradicts the detected niche, the detail isn't
    corroborated, it names the wrong city, or the post looks reshared/tagged
    rather than the company's own job; fall back to the safe city line."""
    niche = _resolve_niche(text_data)

    # Non-trade niches: discard whatever the vision model produced — it was
    # looking at a van, an empty unit or a patient's face — and build the line
    # from the page text instead. Checked FIRST, before the reshare guard,
    # because none of the photo-based guards below are meaningful when the
    # photo was never the source.
    if niche in _HU_NON_PHOTO_NICHES:
        line = _hu_text_opener(niche, text_data)
        if line:
            print(f"  Opener TEXT-SOURCED (niche={niche}): '{line}'")
            return line
        fallback = _hu_city_fallback(text_data)
        # Log WHAT the page actually said when no hook matched. These niches
        # skip the screenshot and the vision call, so scraped text is the only
        # signal there is — and when it yields nothing, guessing at more hook
        # words is exactly the kind of unmeasured change that has never once
        # paid off here. Printing the page's own words turns every fallback
        # into evidence: after a batch, grep server_hu.log for "NO HOOK" and
        # the real vocabulary of these pages is sitting there to hook on.
        _page = text_data
        for _hdr in ("--- COMPANY NAME ---", "--- TRADE TYPE ---", "--- LISTED CITY ---"):
            _page = _page.replace(_hdr, " ")
        _page = " ".join(_page.split())[:240]
        print(f"  Opener → city fallback (niche={niche}, no distinctive service in text)")
        print(f"  NO HOOK [{niche}] page said: {_page!r}")
        return fallback

    if _hu_looks_like_reshare_or_review(text_data):
        fallback = _hu_city_fallback(text_data)
        print(f"  Opener REJECTED (looks reshared/tagged/review): '{opener[:60]}' → '{fallback}'")
        return fallback
    if not (_opener_matches_niche(opener, niche)
            and _hu_corroboration_guard(opener, text_data)
            and _hu_city_guard(opener, text_data)):
        fallback = _hu_city_fallback(text_data)
        print(f"  Opener REJECTED (niche={niche}): '{opener[:60]}' → '{fallback}'")
        return fallback
    return opener

dm_app.CUSTOM_OPENER_POSTPROCESS = _hu_opener_postprocess

# Autószerelő tab: a second sender with its own sheet and its own message,
# living next to Auto Sender. All of it is in hu_mech.py.
import hu_mech
hu_mech.register(dm_app)


# -- Run on port 5002 so it doesn't clash with the English app ----------------
if __name__ == "__main__":
    dm_app.app.run(debug=False, port=5002)

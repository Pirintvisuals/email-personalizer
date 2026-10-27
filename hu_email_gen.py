"""
Hungarian cold-EMAIL generator — batch, offline, parallel.

Excel in -> Excel out. No Facebook, no logged-in Chrome, no CDP port, no
Messenger typing. Personalisation comes from the lead's own WEBSITE, which
means (unlike the DM sender) this can run many leads at once: nothing here
touches an account that can be rate-limited or banned for concurrency.

DESIGN — why this file is short:
It imports dm_app_hu, which is itself a monkey-patch layer over dm_app. Simply
importing it applies every Hungarian patch to the shared engine: the HU vision
PROMPT, the niche-aware prompt builder, the four opener guards (niche
contradiction / corroboration / wrong city / reshare), the Hungarian praise
phrase lists, the qwen vision model + timeout wrapper, and the Groq model
ladder with its bounded retry budget. So the opener written here is generated
by EXACTLY the same code path as a DM opener, with zero duplication — fix a
guard in dm_app_hu.py and this file inherits the fix. Only the delivery format
(email blocks + subject line) is new, and that lives in _hu_email_body below.

Run:
    py hu_email_gen.py "C:\\path\\to\\leads.xlsx"
"""

import os
import re
import sys
import time
import random
import threading
import traceback
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

# Importing dm_app_hu applies all HU patches to dm_app as a side effect.
# It must come before dm_app is used, never after.
import dm_app_hu          # noqa: F401  (imported for its patching side effects)
import dm_app


# ─── CONFIG ────────────────────────────────────────────────────────────────────

# Groq's token-per-minute budget on this account — not the network — is the real
# ceiling here. Vision calls are token-heavy (a base64 JPEG per lead), and
# dm_app_hu.py's own comments record this account hitting 429s and 19-38s stalls
# when opener-writing competed with a second vision call per lead. 5 workers
# keeps throughput high while staying under that budget; raise it only if the
# log stays clean of 429s.
MAX_WORKERS      = 5
REQUEST_TIMEOUT  = 12
SCREENSHOT_W     = 1280
SCREENSHOT_H     = 900
# NOTE: there is deliberately no SIGNOFF/COMPANY constant any more. Gmail
# appends the account's configured signature to messages sent through the API
# too (confirmed against a real bounced message's body, which carried the
# "-- / Landscale / Milán / URL" block below the generated text), so the
# generator writing its own name and company produced each of them twice.
# The closing line in _hu_email_body() now hands off to that signature
# instead of competing with it.

HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"),
    "Accept-Language": "hu-HU,hu;q=0.9,en;q=0.8",
}

# Trade sites put their best photos behind one of these links. Screenshotting a
# gallery page instead of the homepage gives the vision model finished work to
# react to rather than a hero banner and a phone number — the single biggest
# lever on opener quality, since a homepage often has no job photo at all and
# every such lead falls through to the generic city fallback.
_GALLERY_WORDS = ("galéria", "galeria", "referencia", "referenciák", "munkáink",
                  "munkaink", "projektek", "elkészült", "képek", "gallery")


# ─── SITE SCRAPING ─────────────────────────────────────────────────────────────

def _clean(text: str) -> str:
    return re.sub(r"[ \t]+", " ", re.sub(r"\n{3,}", "\n\n", text or "")).strip()


def fetch_site_text(url: str) -> dict:
    """Plain-HTTP fetch of the lead's site. Returns title + visible text.

    Deliberately requests-only, no browser: this runs in a thread pool and a
    headless Chromium per lead would dominate the runtime. The screenshot step
    below is the only part that needs a browser, and it's optional."""
    out = {"ok": False, "url": url, "title": "", "text": "", "gallery_url": "", "error": ""}
    if not url or url.strip().lower() in ("", "n/a", "-", "nincs"):
        out["error"] = "no website"
        return out

    if not url.startswith(("http://", "https://")):
        url = "https://" + url.strip()
    out["url"] = url

    try:
        resp = requests.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT, allow_redirects=True)
        resp.raise_for_status()
    except requests.exceptions.SSLError:
        # Small Hungarian trade sites are frequently on expired or mismatched
        # certs. That says nothing about whether the business is real, so fall
        # back to http rather than dropping the lead.
        try:
            resp = requests.get(url.replace("https://", "http://"), headers=HEADERS,
                                timeout=REQUEST_TIMEOUT, allow_redirects=True)
            resp.raise_for_status()
        except Exception as e:
            out["error"] = f"ssl+http failed: {str(e)[:100]}"
            return out
    except Exception as e:
        out["error"] = f"{type(e).__name__}: {str(e)[:100]}"
        return out

    try:
        soup = BeautifulSoup(resp.text, "lxml")

        # Find the gallery BEFORE stripping nav — the link usually lives in the
        # main menu, which the noise-removal below deletes.
        for a in soup.find_all("a", href=True):
            label = f"{a.get_text(' ')} {a['href']}".lower()
            if any(w in label for w in _GALLERY_WORDS):
                out["gallery_url"] = urljoin(resp.url, a["href"])
                break

        for tag in soup(["script", "style", "head", "noscript", "iframe", "svg"]):
            tag.decompose()

        t = soup.find("title")
        out["title"] = _clean(t.get_text()) if t else ""

        body = soup.find("body")
        # 8k is well past what the facts-extraction step in call_gemini() reads
        # (it slices to 4000), but the niche scorer and the corroboration guard
        # both search the FULL string, and a brand name mentioned once on a long
        # page is exactly what the guard needs to see to let an opener through.
        out["text"] = _clean(body.get_text(separator="\n") if body else "")[:8000]
        out["ok"] = True
    except Exception as e:
        out["error"] = f"parse: {str(e)[:100]}"
    return out


# ─── GROQ THROTTLE ──────────────────────────────────────────────────────────
# The first real 95-lead run hit this hard: 208 separate 429s, and 39/95 leads
# (41%) got ZERO real opener attempt because all three models in the ladder
# were rate-limited simultaneously — dm_app_hu.py's own comments already flag
# this account's tight Groq TPM budget, and 5 workers firing vision calls at
# once blew straight through it. dm_app.call_gemini()'s own retry only sleeps
# 2s between attempts on the SAME model — nowhere near a TPM window reset — so
# the fix has to sit here, not there. Two layers: pace the START of each lead's
# Groq call-sequence so concurrent site-fetching/screenshotting (which costs no
# API quota) doesn't turn into concurrent API bursts, and if a lead still gets
# fully exhausted, cool down and give it one real second attempt before
# accepting the fallback — by then the rolling window has likely cleared.
GROQ_MIN_GAP_S   = 8.0
GROQ_COOLDOWN_S  = 30.0

_groq_lock = threading.Lock()
_groq_last_call = [0.0]


def _throttle_groq():
    with _groq_lock:
        wait = _groq_last_call[0] + GROQ_MIN_GAP_S - time.time()
        if wait > 0:
            time.sleep(wait)
        _groq_last_call[0] = time.time()


def _call_gemini_throttled(text_data: str, images):
    _throttle_groq()
    try:
        return dm_app.call_gemini(text_data=text_data, images=images).strip()
    except Exception as e:
        print(f"    opener failed ({str(e)[:90]}), cooling down {GROQ_COOLDOWN_S:.0f}s for one retry...")
        time.sleep(GROQ_COOLDOWN_S)
        _throttle_groq()
        return dm_app.call_gemini(text_data=text_data, images=images).strip()


_thread_local = threading.local()


def _get_browser():
    """One headless Chromium per worker thread, reused across leads.

    Playwright's sync API is bound to the thread that created it, so this can't
    be a single shared browser — but launching one per LEAD would cost ~1.5s
    every time. Per-thread is the right granularity: MAX_WORKERS browsers total,
    each handling a whole share of the list."""
    if getattr(_thread_local, "browser", None) is None:
        from playwright.sync_api import sync_playwright
        pw = sync_playwright().start()
        _thread_local.pw = pw
        _thread_local.browser = pw.chromium.launch(headless=True)
    return _thread_local.browser


def _close_browser():
    if getattr(_thread_local, "browser", None) is not None:
        try:
            _thread_local.browser.close()
            _thread_local.pw.stop()
        except Exception:
            pass
        _thread_local.browser = None


def screenshot_site(url: str, gallery_url: str = "") -> bytes | None:
    """Viewport screenshot of the gallery page if there is one, else the
    homepage. Viewport, not full_page: a full-page shot of a long trade site is
    a huge image that costs vision tokens without adding usable detail, and the
    model only needs one finished job to react to."""
    target = gallery_url or url
    if not target:
        return None
    try:
        page = _get_browser().new_page(viewport={"width": SCREENSHOT_W, "height": SCREENSHOT_H})
        try:
            page.goto(target, wait_until="domcontentloaded", timeout=25000)
            page.wait_for_timeout(1200)          # let hero images and lazy loads paint
            # Cookie banners cover the top third of most Hungarian sites and are
            # the single most common thing a vision model ends up describing.
            try:
                page.evaluate("""() => {
                    const kill = ['[id*="cookie" i]','[class*="cookie" i]',
                                  '[id*="consent" i]','[class*="consent" i]','[role="dialog"]'];
                    kill.forEach(s => document.querySelectorAll(s).forEach(el => el.remove()));
                    document.documentElement.style.overflow = 'auto';
                }""")
            except Exception:
                pass
            page.wait_for_timeout(300)
            return page.screenshot(type="jpeg", quality=72)
        finally:
            page.close()
    except Exception as e:
        print(f"    screenshot skipped: {str(e)[:90]}")
        return None


# ─── EMAIL COPY ────────────────────────────────────────────────────────────────

# Tracks only the most recently produced subject line so two consecutive
# sends never carry the exact same one — a real (if small) spam signal, since
# inbox providers do look at whether a sender's recent messages all look
# alike. Lock + list-of-one because the batch script runs MAX_WORKERS leads
# concurrently; a plain module variable would race.
_subject_state_lock = threading.Lock()
_last_subject = [""]


def _pick_subject(pool: list) -> str:
    """random.choice, but excludes whatever the previous call returned first
    (falls back to the full pool only if that empties it out, e.g. a 1-item
    pool)."""
    with _subject_state_lock:
        choices = [p for p in pool if p != _last_subject[0]] or pool
        pick = random.choice(choices)
        _last_subject[0] = pick
        return pick


def _vary(options: list) -> str:
    """Pick one of a few close paraphrases of the same sentence — same claim,
    same facts, slightly different wording — so two sent emails aren't
    byte-identical in their boilerplate paragraphs. Deliberately subtle: this
    swaps connector words and phrasing, not the substance, so the tested copy
    stays intact. Not meant to defeat spam filtering by itself, just to avoid
    handing providers dozens of textually-identical outgoing emails in a row,
    which is its own small signal on top of everything else already varying
    per lead (name, opener, subject)."""
    return random.choice(options)


def _hu_subject(first: str, niche: str, hero_img: bool = False, hook: str = "") -> str:
    """Short, all-lowercase Hungarian subject line.

    Refreshed against 2025-2026 cold-email subject-line research (Belkins'
    5.5M-email study, Gong's 85M-email study, Woodpecker's 20M-email study —
    all three agree on the same shape):
      - all-lowercase beats Title Case by ~21% (reads like a colleague, not
        a marketing blast)
      - 2-4 words is the best-performing length band in every dataset (46%
        open rate); Belkins/Gong/Lavender all independently converge here
      - a first name alone is worth ~+31% opens vs. none
      - hype/selling words ('AI', 'chatbot', 'ajánlat', 'ingyenes') cut opens
        by up to ~18% — none of these variants use any
      - the SINGLE strongest lever across all three studies is trigger-event
        personalization: referencing something concretely, verifiably true
        about the recipient (not just their name) lifts opens ~42% over a
        plain personalized line (54.7% vs. ~38%)
      - 33 characters is roughly where mobile clients start truncating mid-
        word, so these stay under or close to that

    hero_img=True means there's an actual screenshot of the lead's own site
    or gallery to point at — a real trigger event, not filler — so those
    variants replace the generic ones rather than just adding to the pool.
    Two follow-up fixes on top of the original research pass:
      - 'a képeitek alapján' was dropped for being a near-duplicate of 'jó
        kis galéria' (both just say "I saw your photos") — replaced with 'a
        saját oldalad alapján', which echoes the body's own line ("Fent a
        kép a saját oldaladról") word-for-word. Subject/body wording match
        is its own small deliverability signal on top of everything else —
        a subject that doesn't quote the same fact the body opens with
        reads as less coherent even though both are individually true.
      - 'neked, {f}' was dropped outright: it carries zero information scent
        (no hint of a question, an offer, or what it's about), the one
        clearly weakest entry in either pool.

    Third fix — 'gyors kérdés' removed from every pool: current data flags
    it specifically (not "questions" generally, that exact phrase) as one of
    the two most fatigued cold-email openers left — one 2026 measurement put
    it at 28.7% opens against 32.9% for non-generic alternatives, because a
    trained eye (and increasingly spam filters) now reads it as templated
    rather than actually personal. Every remaining source agrees the fix is
    the same lever as the body copy already leans on: reference something
    real about THIS business, not a mail-merge-shaped phrase that could
    point at anyone.

    Fourth fix — switched to magázás (formal 'Ön' address) throughout, per
    explicit request: this list skews toward an older trade-business owner
    generation where tegezés from a stranger reads as presumptuous rather
    than friendly. 'neked' became 'Önnek'; possessive forms ('munkáitokat',
    'oldalad') dropped their 2nd-person suffix in favour of the
    3rd-person-singular form that doubles as formal Ön ('munkáit', 'oldala')
    — grammatically identical to plain 3rd person, which is exactly how
    Hungarian formal address is built. 'kérdés hozzád' → 'kérdés Önhöz' was
    the first pass, but a bare noun-phrase + 'Önhöz' reads noticeably
    stiffer in formal register than the same clipped shape did in tegező
    (informal register tolerates fragments far better) — replaced with
    'kérdésem volna', a standard, fully natural magázás-register phrase for
    exactly this, that also sidesteps needing 'Ön' as a bare noun at all.

    That's also why niche now actually gets used here — it was accepted as
    a parameter this whole time but never touched a single subject variant,
    which was the one cheap, always-available specificity signal (every
    lead has a niche; not every lead has a hero image) sitting unused. Kept
    to a length guard (~34 chars) since niche labels vary wildly in length
    ('burkoló' vs 'homlokzatszigetelő' vs 'garázskapu és árnyékolás') and a
    subject that wraps past ~40 chars just gets cut on mobile instead of
    read.

    Fifth fix — 'hook': niche is a specificity signal, but it's the WEAKEST
    one available ("tetőfedő kérdés" is true of every roofer in Hungary, not
    just this one). 'hook' is the same concrete, page-sourced fact the
    deterministic opener tier already extracts — a real brand/material name
    or a founding-year claim actually printed on the lead's own site (see
    _find_niche_brand_mention / _find_years_in_business, called once in
    process_lead and threaded through to both the opener AND here) — which
    is exactly the trigger-event personalization the research above flags
    as the single strongest lever (+42% opens). When one was found, it
    overrides niche/generic entirely rather than being diluted into the
    same random pool as weaker filler: a verifiably-true specific fact
    should win every time it's available, not 1-in-5. No verb conjugation
    in either hook phrasing (mirrors the niche_variant's own noun-phrase
    shape) so there's no tegezés/magázás risk to worry about here at all."""
    f = first.lower() if first else ""

    hook_pool = []
    if hook:
        if hook.endswith(("éve", "óta")):
            cands = [
                f"{f}, {hook} a szakmában" if first else f"{hook} a szakmában",
                f"{f}, {hook} a piacon" if first else f"{hook} a piacon",
            ]
        else:
            cands = [
                f"{f}, {hook} kérdés" if first else f"{hook} kérdés",
                f"{f}, {hook} téma" if first else f"{hook} téma",
            ]
        hook_pool = [c for c in cands if len(c) <= 34]
    if hook_pool:
        return _pick_subject(hook_pool)

    niche_variant = None
    if niche:
        cand = f"{first.lower()}, {niche} kérdés" if first else f"{niche} kérdés"
        if len(cand) <= 34:
            niche_variant = cand

    if first:
        f = first.lower()
        if hero_img:
            pool = [
                f"{f}, megnéztem a munkáit",
                f"{f}, jó kis galéria",
                f"{f}, a saját oldala alapján",
                f"{f}, megéri egy percet?",
            ]
        else:
            pool = [
                f"összeraktam Önnek valamit, {f}",
                f"{f}, megéri egy percet?",
                f"kérdésem volna, {f}",
                f"{f}, az árkérésekről",
            ]
        if niche_variant:
            pool.append(niche_variant)
        return _pick_subject(pool)

    pool = [
        "összeraktam Önnek valamit",
        "megéri egy percet?",
        "az árkérésekről",
    ]
    if niche_variant:
        pool.append(niche_variant)
    return _pick_subject(pool)


# ── Named-client proof, EMAIL VERSION ────────────────────────────────────────
# The DM carries dm_app_hu._HU_NICHE_PROOF, which is the trade's problem PLUS
# the client name. That works in a DM because nothing else there states the
# problem. In an email it is redundant: the reframe paragraph ("Nem több
# érdeklődőt ígér, hanem jobbakat") has already run three paragraphs above,
# so restating "pedig felmérés nélkül nem lehet pontosat mondani" asks the
# reader to agree to something they agreed to already, and costs 19 words to
# do it. Email keeps only the half the rest of the message does NOT contain:
# who else has one.
#
# Names live here once, in the nominative-plus-suffix form the sentence needs,
# so there is a single place to correct them. They must match
# landscale.agency/hu/case-studies or the whole point of naming a client
# (that a prospect can go and check) is lost.
_CLIENT_BY_NICHE = {
    "fürdőszoba felújító": "Az NM Baunak",
    "felújító kivitelező": "Az NM Baunak",
    "fűtésszerelő":        "A kecskeméti Hírös Kft-nek",
    "klíma":               "A Klímaszerelő Kecskemétnek",
}


def _named_proof(niche: str) -> str:
    """'<Client>-nak is pont ezt építettem, ott él most is.', or '' if there is
    no real client in that trade yet (caller then uses the anonymous line)."""
    client = _CLIENT_BY_NICHE.get(niche)
    return f"{client} is pont ezt építettem, ott él most is." if client else ""


def _hu_email_body(opener: str, first: str, niche: str, hero_img: bool = False,
                   city: str = "") -> str:
    """Email-native version of the DM pitch — same substance (tested copy, not
    rewritten from scratch), but re-PACED to read like a Charlie Morgan style
    cold email: one short thought per paragraph, a blank line between every
    single one, no em dashes anywhere (the user's explicit note: reads more
    human, closer to how these high-volume cold-email senders actually write).

    The opener is the FIRST thing after the greeting and stays its own short
    paragraph — it is the personalization, and it needs to stand alone so it's
    the first thing read, not buried after a scene-setting intro paragraph.

    city: if the lead is from Miskolc, drops in a one-line "I'm from there
    too" — the user's own hometown, and the one personal-connection line that
    actually applies to a real subset of these leads. Kept to a single short
    paragraph right after the opener, same rhythm as everything else.

    hero_img: True once a screenshot of the lead's own site is embedded above
    this text (see process_lead / hu_email_app.py) — swaps the "összeraktam
    egyet" line to point at the image instead of an abstract claim, since the
    picture IS the proof at that point.

    Two deliberate differences from the DM, unchanged from before:
    - 'Kint van róla egy poszt a profilomon' is DROPPED — no profile to check
      in an email.
    - No link anywhere, by explicit decision. Keeps 'Elküldhetem?' as the ask.
    """
    niche_label = f"{niche} " if niche and niche != "felújító kivitelező" else "építőipari "
    # Shared with the DM channel. Without it this email said only "rákérdez a
    # részletekre", which is the vaguest line in the message and the one that
    # most needs to be concrete — a mover reading a materials question, or a
    # dentist reading "mekkora felület", knows instantly it was not written
    # for them.
    questions = dm_app_hu._HU_NICHE_QUESTIONS.get(
        niche, dm_app_hu._HU_DEFAULT_QUESTIONS)

    blocks = [f"Kedves {first}!" if first else "Jó napot kívánok!"]

    op = (opener or "").strip()
    if op:
        op = op[:1].upper() + op[1:]
        if not op.endswith((".", "!", "?")):
            op += "."
        blocks.append(op)

    is_miskolc = "miskolc" in (city or "").lower()

    # Moved to the END of the pitch (was mid-email, between the proof and the
    # mechanism). It's the setup for the CTA, not a feature — "I made you one,
    # can I send it" reads as a single thought, and splitting them across four
    # paragraphs made the ask arrive out of nowhere.
    if hero_img:
        sample_lines = [
            "Fent a kép az Ön oldaláról.",
            "Ez alapján raktam össze Önnek egyet, hogy lássa élesben.",
        ]
    else:
        sample_lines = ["Összeraktam Önnek egyet, hogy lássa élesben."]

    # LENGTH BUDGET: the pitch (greeting through sign-off, excluding the legal
    # footer) targets UNDER 80 words. 2025-2026 reply-rate data is consistent
    # that 50-125 words wins overall, but first touch specifically does best
    # at the BOTTOM of that band: Sales.co's 3.1M-email analysis puts 75-100
    # words at the highest reply rate and 50-125 at 2.4x the reply rate of
    # 200+, Instantly recommends under 80 for a first email, and Hunter's
    # 34M-email study found 20-39 words topping every band at 4.5%.
    #
    # STRUCTURE: every mechanism line is immediately followed by what it GETS
    # them. A tradesperson does not care that a bot asks qualifying questions;
    # they care that they stop driving across the county for a job that was
    # never real. Feature and outcome were previously separated by the proof
    # line and the sample line, which meant the mechanism landed with no
    # payoff attached and the payoff arrived with no mechanism to explain it.
    # Now they're adjacent pairs, and the proof/sample come after, once the
    # value is already understood.
    #
    # POSITIONING — the single most important line in the email:
    # "nem a felmérést váltja ki" runs FIRST, before any mechanism. The
    # instinctive objection from someone who quotes by walking the job is
    # "a bot can't price my work" — and they're right, which is why the
    # product doesn't try to. It removes the WASTED visits, not the visits.
    # This is deliberately NOT the same category as the two defensive lines
    # cut earlier ("a végleges ajánlat mindig Önnél marad", "semmi nem
    # változik"), which were reassurance about nothing in particular. This
    # one defines what the product IS, and getting it wrong loses the reader
    # before any feature line can land. It also aligns the whole email with
    # the one thing that actually reads as valuable to a booked-out trade:
    # their evenings back, not more leads.
    blocks += [
        # What-I-build in one line. "Diák vagyok" removed on request (same
        # change as HU_MENTION_STUDENT in dm_app_hu.py — the two channels hit
        # the same market, so a lead who gets both would otherwise see the age
        # framing in one and not the other). "Mesterséges intelligencia" stays
        # dropped here: "árkalkulátor chatbot" already says it, in the lead's
        # own vocabulary, and this is the email's first line where the DM's
        # isn't.
        _vary([
            f"Árkalkulátorokat készítek {niche_label}cégeknek.",
            f"Árkalkulátorokat építek {niche_label}cégeknek.",
        ]),

        # Two short sentences rather than one long one: the reframe lands
        # harder as a flat denial followed by the correction, and it matches
        # the one-thought-per-paragraph rhythm the rest of the email uses.
        # This paragraph used to promise fewer wasted callouts. A working
        # klímaszerelő (Keresztesi Károly, k-klima) killed that framing with
        # real numbers: of 10 cold quote requests 6 reply, 3 go on-site, and
        # the 3 who showed up in person are the ones who win the contract.
        # The survey is where the sale and the margin happen — placement
        # renegotiation, equipment advice — so selling its removal is selling
        # against the owner's own best move. The promise is better customers
        # reaching that survey, never fewer surveys.
        _vary([
            "Nem több érdeklődőt ígér, hanem jobbakat. Olyanokat, akik "
            "tényleg meg akarják csináltatni.",
            "A felmérést nem váltja ki. Abban segít, hogy azokhoz jusson el, "
            "akik tényleg meg akarják csináltatni.",
            "A felmérés marad a lényeg. Csak azok jutnak el odáig, akiknél "
            "van is miért kimenni.",
        ]),

        # FEATURE. "nem fix árat" is load-bearing and stays word-for-word —
        # it's the difference between offering an estimate and appearing to
        # quote a binding price, which is what a tradesperson objects to
        # hardest.
        _vary([
            f"Amikor valaki árat kér, rákérdez a részletekre: {questions}, "
            f"stb., majd megmondja az alapárat, és azt is, mi az, ami azon "
            f"felül jön.",
            f"Ha valaki árat kér, rákérdez a részletekre: {questions}, stb., "
            f"majd megmondja az alapárat, és hogy mi az, ami külön árazás.",
        ]),

        # The example questions are illustrations, not the script. Without
        # this a company whose work spans more than the three listed items
        # discounts the whole thing on a question list that misses them.
        #
        # The PRICES are named here too, not just the questions. The paragraph
        # above now promises an alapár, and the unasked question that leaves
        # behind is "whose base price?" — a tradesman assumes a stranger's
        # calculator carries a stranger's numbers, which is the same reflex
        # behind "arat kalkulalni sem igazan lehet". Saying both are set from
        # his own pricing answers it in one clause.
        _vary([
            "A kérdéseket és az árakat is az Ön munkájához igazítanám.",
            "A kérdések és az árak is az Ön saját árazása alapján mennének.",
        ]),

        # OUTCOME of that feature, in the same breath: the filtering already
        # happened, so the notification only ever describes someone worth the
        # drive. Name + phone number stay concrete (a vague "értesítést kap"
        # sounds like a dashboard to log into; a name and a number sounds
        # like a job).
        _vary([
            "Így csak azokról kap értesítést névvel és telefonszámmal, akik "
            "komolyan gondolják.",
            "Így arról kap értesítést névvel és telefonszámmal, aki komolyan "
            "gondolja.",
        ]),

        # The "6 milliós munka" figure itself is a factual claim and stays
        # fixed across every variant below — only the surrounding wording
        # changes, never the number.
        # Named client + that trade's own problem where one exists, falling
        # back to the anonymous 6M line for trades with no client yet. Shared
        # with the DM channel via dm_app_hu._HU_NICHE_PROOF so the two cannot
        # drift apart again — see the note there on why the problem is stated
        # as the trade's, never as a private fact about the named company.
        _named_proof(niche) or _vary([
            "Az egyik ügyfelemnél az első héten bejött rajta egy 6 milliós munka.",
            "Az egyik ügyfelemnél már az első héten befutott rajta egy 6 milliós munka.",
        ]),

        *sample_lines,
    ]

    # Moved here from right after the opener on purpose: up top it read as
    # just another personalization fact stacked with the rest. Down here,
    # right before the ask, it lands as a genuine closing gesture — "we're
    # from the same place, happy to meet up" — and doubles as a soft, local,
    # low-pressure way into the CTA rather than a cold "can I send this".
    if is_miskolc:
        blocks += [
            "Egyébként én is miskolci vagyok.",
            "Ha szeretné, akár személyesen is összefuthatunk.",
        ]

    blocks += [
        # The ask, made deliberately small — same change as the DM channel.
        # Conventional cold-email advice closes on the value; this closes on
        # how little is being requested, because persuasion is not where this
        # funnel loses people (they DO say yes) — the size of the next step
        # is. And it is literally true: nothing is being sold in this message,
        # the only request is permission to send a link. Saying so costs
        # nothing and removes the reason to stall.
        _vary([
            "Csak megmutatnám, mást nem kérek. Elküldhetem?",
            "Csak megmutatnám, semmi mást nem kérek. Küldhetem?",
            "Egyelőre csak megmutatnám, mást nem kérek. Átküldjem?",
        ]),

        # OPT-OUT — one line, and it sits ABOVE the closing on purpose so the
        # message ends "Üdvözlettel," → Gmail's signature, with nothing
        # wedged between them.
        #
        # The identification half of the old footer is gone: Gmail's own
        # signature already states Landscale, Milán and the site on every
        # message, so repeating it here printed both twice. What the
        # signature does NOT carry is an opt-out, and that's the half with
        # legal weight behind it — Hungary is a strict opt-in jurisdiction
        # for e-mail advertising (Grt. 2008. évi XLVIII. tv. §6(1) requires
        # prior express consent to advertise to a NATURAL PERSON, which in
        # Hungarian practice covers an egyéni vállalkozó and any
        # named-person address, however business-like it looks). So the
        # sentence stays even though the rest of the footer went.
        #
        # A reply-based opt-out rather than an unsubscribe link is the
        # honest form for a solo sender: there is no list to remove anyone
        # from, only a person who reads the replies. It also adds no link of
        # its own.
        #
        # TARGETING is still the real control, not this line — see
        # _address_risk() in hu_email_app.py, which flags a non-functional
        # address in the review UI before it can be sent to. This reduces
        # exposure; it does not make the campaign definitively compliant,
        # and none of it is legal advice.
        "Ha nem kéri, válaszoljon egy szóval, és többé nem keresem.",

        # No name and no company here: Gmail appends the account signature
        # ("-- / Landscale / Milán / URL") directly below this line.
        "Üdvözlettel,",
    ]
    return "\n\n".join(blocks)


# ─── PER-LEAD PIPELINE ─────────────────────────────────────────────────────────

# Canonical HU niche labels, as used by _NICHE_VOCAB / NICHE_OPENER_INSTRUCTIONS.
# dm_app_hu._resolve_niche() only direct-maps ENGLISH spreadsheet labels
# ("roofing", "bathroom", ...) because the DM lead sheets were built in English.
# A Hungarian list will naturally say "tetőfedő" or "klíma", which would fall
# through to text scoring and could be overruled by noisy website copy. So an
# exact match on the canonical label is trusted outright — if you typed it in
# the sheet, you know better than the scorer.
_HU_CANON_NICHES = {
    "klíma", "napelemes", "tetőfedő", "villanyszerelő", "vízvezetékszerelő",
    "kőműves", "festő", "gipszkartonos", "burkoló", "padlóburkoló",
    "fürdőszoba felújító", "asztalos", "medenceépítő", "garázskapu és árnyékoló",
    "térkövező", "kertépítő", "terasz és deszka", "nyílászáró",
    "homlokzatszigetelő", "felújító kivitelező", "fűtésszerelő",
}
# Everyday spellings someone would actually type into a spreadsheet column.
# dm_app_hu._resolve_niche() trusts an Excel "Niche" hint only on an EXACT
# lowercase match against _EXCEL_NICHE_DIRECT_MAP's four keys ("roofing",
# "bathroom", "construction", "renovation") or the literal "boiler/hvac" — it
# does not substring-match. This lead sheet's real Niche column (checked
# against all 95 rows) contains exactly four English values, none of them
# exact: "Bathroom remodeling", "Gas boiler installation/service", "Roofing",
# "Construction". Left unmapped, every lead's real ground-truth niche would be
# silently discarded in favour of a guess from scraped site text. Substring
# match here, in insertion order, down to the tokens dm_app_hu actually checks.
_EN_NICHE_MAP = {
    "boiler": "boiler/hvac",
    "hvac": "boiler/hvac",
    "heat pump": "boiler/hvac",
    "bathroom": "bathroom",
    "roofing": "roofing",
    "roof": "roofing",
    "renovation": "renovation",
    "general contractor": "construction",
    "construction": "construction",
}

_HU_NICHE_ALIASES = {
    "klímaszerelő": "klíma", "klimaszerelo": "klíma", "hőszivattyú": "klíma",
    "napelem": "napelemes", "szolár": "napelemes",
    "tetőfedés": "tetőfedő", "tetofedo": "tetőfedő", "tető": "tetőfedő",
    "villanyszerelés": "villanyszerelő", "vízszerelő": "vízvezetékszerelő",
    "gázszerelő": "fűtésszerelő", "kazán": "fűtésszerelő", "fűtés": "fűtésszerelő",
    "fűtésszerelés": "fűtésszerelő", "épületgépész": "fűtésszerelő",
    "festés": "festő", "szobafestő": "festő", "mázoló": "festő",
    "burkolás": "burkoló", "csempéző": "burkoló",
    "fürdőszoba": "fürdőszoba felújító", "fürdőszoba felújítás": "fürdőszoba felújító",
    "ablak": "nyílászáró", "nyílászárók": "nyílászáró", "ablakos": "nyílászáró",
    "térkő": "térkövező", "térkövezés": "térkövező",
    "kertépítés": "kertépítő", "kertész": "kertépítő",
    "homlokzat": "homlokzatszigetelő", "szigetelés": "homlokzatszigetelő",
    "generálkivitelező": "felújító kivitelező", "felújítás": "felújító kivitelező",
    "kivitelező": "felújító kivitelező", "építőipar": "felújító kivitelező",
}


def canon_niche(raw: str) -> str:
    """Normalise the sheet's Niche column into whatever dm_app_hu will actually
    recognise: the four exact English tokens it exact-matches for the Excel
    hint, or a Hungarian canonical label for plain-text scoring. Falls through
    to the raw value so the text scorer still has something to look at rather
    than nothing."""
    r = (raw or "").strip().lower()
    if r in _HU_CANON_NICHES:
        return r
    if r in _HU_NICHE_ALIASES:
        return _HU_NICHE_ALIASES[r]
    for needle, token in _EN_NICHE_MAP.items():
        if needle in r:
            return token
    return r


# ─── DETERMINISTIC (NO-API) OPENER FALLBACK ────────────────────────────────
# The old fallback tier, when the AI opener failed for ANY reason (Groq 429,
# timeout, every guard rejecting the draft), was one line: "Rátaláltam az
# oldalatokra <city>." — true, but says nothing about THEM, which is exactly
# the "there's no personalized line" complaint this was meant to fix in the
# first place. That AI call needs Groq to be up and under quota; this tier
# doesn't — it's a regex/keyword match against site["text"], which is fetched
# with plain requests before any API call happens at all, so it's available
# even when Groq is completely down. It can never be factually wrong either:
# every line it can produce is built from a literal substring already sitting
# on the lead's own page, not a guess.
_YEARS_IN_BUSINESS_RE = re.compile(
    r"(\d{4})\s*óta|(\d{1,2})\s*év(?:e|es tapasztalat)", re.IGNORECASE)

# A bare substring match can't tell "we work WITH Daikin" from "we work
# WITHOUT Daikin" — both contain the word "daikin". Checked in a small window
# right before a match; cheap (one more `in` check per candidate, no regex,
# no I/O), but it's the difference between a real fact and a claim the lead's
# own page actually contradicts.
_NEGATION_WORDS = ("nem ", "nélkül", "kivéve", "sem ")
# Requiring one of these near a bare "X éve" / "X óta" keeps it anchored to an
# actual business-history claim, not an unrelated number on the page (a staff
# member's tenure, a building's age, a testimonial's timeframe) that happens
# to match the same pattern.
_YEARS_CONTEXT_WORDS = ("cég", "vállalkoz", "tapasztalat", "piacon",
                        "szakmában", "alapít", "működ", "csapat")


def _find_years_in_business(text: str) -> str:
    """'2005 óta' or '15 éve' style claims are common on Hungarian trade
    sites' About/home sections. Sanity-bounded so a stray unrelated 4-digit
    number (a postcode, a phone number fragment) can't be mistaken for a
    founding year, AND requires a business-context word nearby so an
    unrelated "X éve" elsewhere on the page (staff tenure, a testimonial)
    isn't repeated back as if it were the company's own history."""
    m = _YEARS_IN_BUSINESS_RE.search(text or "")
    if not m:
        return ""
    t_l = (text or "").lower()
    start, end = m.span()
    window = t_l[max(0, start - 60):end + 60]
    if not any(w in window for w in _YEARS_CONTEXT_WORDS):
        return ""
    if m.group(1):
        year = int(m.group(1))
        if 1950 <= year <= datetime.now().year:
            return f"{year} óta"
    elif m.group(2):
        n = int(m.group(2))
        if 1 <= n <= 60:
            return f"{n} éve"
    return ""


def _find_niche_brand_mention(niche: str, text: str) -> str:
    """Reuses dm_app_hu's own niche keyword vocabulary (hu_niche.py's
    _TEXT_KEYWORDS — real brand names and specific trade terms, the same list
    the niche classifier itself scores against) to find something concrete
    actually printed on the lead's page. Only weight>=4 entries count: those
    are specific brand/material names ('bramac', 'daikin', 'velux'), not
    generic trade words ('kazán') that would be true of almost any competitor.

    That list also contains entries that are just the niche's own name spelled
    a few ways ('tetőfedő', 'tetőfedés' for the 'tetőfedő' niche) — matching
    those tells the reader nothing ("I see you do roofing" on a roofer's own
    site isn't personalization). Skipped via a shared-root check on the
    accent-stripped forms, so a real product/brand name further down the same
    list gets picked instead.

    Also skips a match sitting right after a negation word ("nem dolgozunk
    Daikin klímákkal") — a plain substring hit can't otherwise distinguish
    that from a real endorsement, and repeating a denial back as a compliment
    would be worse than saying nothing."""
    text_l = (text or "").lower()
    niche_n = dm_app_hu._norm(niche)[:5]
    for term, weight in dm_app_hu._TEXT_KEYWORDS.get(niche, []):
        if weight < 4:
            continue
        idx = text_l.find(term)
        if idx == -1:
            continue
        if dm_app_hu._norm(term)[:5] == niche_n:
            continue
        window = text_l[max(0, idx - 25):idx]
        if any(neg in window for neg in _NEGATION_WORDS):
            continue
        return term
    return ""


_VOWELS = "aáeéiíoóöőuúüű"


def _a_az(word: str) -> str:
    """Hungarian's definite article is 'a' before a consonant sound and 'az'
    before a vowel sound. brand comes from a fixed keyword list (see
    _find_niche_brand_mention), so this only ever needs the plain
    first-letter check, not the harder acronym/number edge cases ('az USB',
    'a 5-ös') that a general-purpose version would."""
    return "az" if word[:1].lower() in _VOWELS else "a"


# Proper nouns that _TEXT_KEYWORDS happens to contain. Everything NOT in here
# is treated as an ordinary Hungarian word and left lowercase mid-sentence.
# Kept as an explicit list rather than a heuristic (capitalisation in the
# source list, word length, etc.) because the failure is silent and only shows
# up in a sent email.
_BRAND_TOKENS = {
    "daikin", "mitsubishi", "kaisai", "midea", "gree", "panasonic", "toshiba",
    "vaillant", "viessmann", "bosch", "buderus", "junkers", "baxi", "ferroli",
    "ariston", "nibe", "stiebel", "immergas", "worcester",
    "longi", "jinko", "fronius", "sma", "risen", "trina", "huawei",
    "bramac", "tondach", "eternit", "creaton", "icopal", "velux", "fakro",
    "lindab", "tegola", "zenit",
    "schüco", "schuco", "rehau", "veka", "internorm", "gealan", "kbe",
    "semmelrock", "bradstone", "marshalls", "tobermore",
    "knauf", "rigips", "somfy", "hörmann", "hormann", "raffstore",
    "cerec", "invisalign", "carglass",
}


def _is_brand_token(term: str) -> bool:
    """True only for real proper nouns, so ordinary trade words stay lowercase."""
    return term.strip().lower() in _BRAND_TOKENS


def _deterministic_opener(niche: str, site_text: str) -> str:
    """Best zero-API opener available from what's already been scraped.
    Grammar is kept deliberately simple (term stays in nominative case as a
    subject complement) so it's correct no matter what noun/brand matched —
    Hungarian's case suffixes would otherwise need per-word handling this
    isn't worth building. Returns '' if nothing concrete was found, so the
    caller falls through to the plain city line as the true last resort."""
    brand = _find_niche_brand_mention(niche, site_text)
    if brand:
        # _find_niche_brand_mention returns the term exactly as it's stored
        # in dm_app_hu._TEXT_KEYWORDS — all-lowercase, since that list is
        # matched against already-lowercased page text. Fine for the SUBJECT
        # line (all-lowercase is the deliberate house style there, see
        # _hu_subject's hook), but this is a full sentence with normal
        # capitalization rules, and 'internorm'/'daikin' sitting uncapitalized
        # mid-sentence reads as a typo, not a real brand name — .capitalize()
        # fixes the common case (proper first-letter-cap brands); it won't
        # get an all-caps name like 'KAISAI' exactly right, but "Kaisai" is
        # still correct-looking Hungarian, which uncapitalized never is.
        # Only capitalise ACTUAL BRANDS. The matched term comes from
        # dm_app_hu._TEXT_KEYWORDS, which mixes proper nouns ("daikin",
        # "vaillant") with ordinary trade words ("öntözőrendszer",
        # "padlófűtés", "kijelzőcsere"). Capitalising the second kind
        # mid-sentence produced "Látom, hogy Önnél is szerepel az
        # Öntözőrendszer." — a capitalised common noun is the single loudest
        # mail-merge tell there is, and this line exists precisely to look
        # hand-written.
        brand_disp = brand.capitalize() if _is_brand_token(brand) else brand
        return f"Látom, hogy Önnél is szerepel {_a_az(brand)} {brand_disp}."
    years = _find_years_in_business(site_text)
    if years:
        return f"Látom, hogy már {years} ezt csinálja."
    return ""


# NOT a second niche classifier. An earlier version of this file resolved the
# niche itself and passed the result to the email body, while the opener was
# written and validated against dm_app_hu._resolve_niche(text_data) — which
# resolves independently. Confirmed live on a real test lead: the body said
# "fűtésszerelő" while the prompt builder logged niche='tüzeléstechnikai' and
# injected a different niche's opener instructions. Two disagreeing answers for
# one lead is worse than either answer alone, because the guards then validate
# the opener against a niche the email never mentions. One resolver, always —
# the sheet's label reaches it via the normalised TRADE TYPE line in
# _build_text_data() below, which is exactly the channel dm_app_hu already reads.
def resolve_niche(text_data: str) -> str:
    return dm_app_hu._resolve_niche(text_data)


def _build_text_data(lead: dict, site: dict) -> str:
    """Assemble the string the HU guards expect.

    The '--- HEADER ---' markers are not decoration: dm_app_hu._td_section()
    parses them, and three separate guards read them by name. COMPANY NAME feeds
    the name-first niche classifier (the most reliable one), TRADE TYPE feeds the
    spreadsheet hint, and LISTED CITY is what _hu_city_guard compares the
    opener's city against. Omit any of them and that guard silently stops
    working."""
    parts = []
    if lead.get("name"):
        parts.append(f"--- COMPANY NAME ---\n{lead['name']}")
    if lead.get("niche"):
        parts.append(f"--- TRADE TYPE ---\n{canon_niche(lead['niche'])}")
    if site.get("title"):
        parts.append(f"--- PAGE TITLE ---\n{site['title']}")
    parts.append(site.get("text", ""))
    if lead.get("city"):
        parts.append(f"--- LISTED CITY ---\n{lead['city']}")
    return "\n\n".join(p for p in parts if p)


def _clean_city(raw: str) -> str:
    """Strip parenthetical/appositive junk off a spreadsheet location value.

    Real data, not hypothetical: this sheet has rows like 'Bükkábrány
    (Miskolc/Borsod region)', and dm_app_hu._city_to_ragos() only knows ~40
    major-city names — anything else, it returns unchanged. Left as-is, that
    English annotation ends up verbatim inside a Hungarian sentence
    ('Rátaláltam az oldalatokra Bükkábrány (Miskolc/Borsod region).'), which is
    a far more obvious tell than a small town simply not being grammatically
    inflected. Can't fix the missing inflection without a real dict entry, but
    can stop the visibly broken part."""
    c = (raw or "").strip()
    c = re.split(r"[(,/]", c)[0].strip()
    return c


def _first_name(raw: str) -> str:
    """Given name for the greeting, in Hungarian name order (family name first,
    so the given name is the LAST token). Returns '' unless the token is a name
    we actually recognise — greeting a company by its trading name ('Szia Kft.,')
    is worse than a plain 'Szia,'."""
    for token in reversed(re.split(r"[\s,.()&/-]+", (raw or "").strip())):
        t = token.strip(".,()")
        if t in dm_app_hu._HU_GIVEN_NAMES:
            return t
    return ""


def _first_name_for(lead: dict) -> str:
    """The owner column if the sheet has one, otherwise whatever real given
    name is sitting inside the COMPANY name.

    Hungarian trade lists very often have no separate owner column at all,
    but the company name carries the owner outright: 'Kun Imre E.V.',
    'Bodnár Zsolt Építőipari Munkák', 'CSÓ-Szerelés (Pulai Tamás)',
    'Komplettfelújítás (Siegel Csaba EV)'. On the current 150-lead list
    that's 31 leads whose greeting would otherwise fall back to the
    impersonal 'Jó napot kívánok!' and whose subject line would lose every
    first-name variant — the single cheapest personalisation signal there
    is (worth roughly +31% opens on its own, per the research quoted in
    _hu_subject).

    Safe because _first_name() only ever returns a token from
    dm_app_hu._HU_GIVEN_NAMES: an unrecognised trading name yields '' and
    the greeting stays impersonal, exactly as before. The reverse scan also
    does the right thing on these — 'Kun Imre E.V.' reads right-to-left as
    V, E, Imre, so the given name wins over the surname, which is the same
    Hungarian name-order logic the owner column already relies on."""
    return _first_name(lead.get("owner", "") or "") or _first_name(lead.get("name", "") or "")


def process_lead(lead: dict) -> dict:
    """Scrape -> screenshot -> opener (guarded) -> email. Never raises; a failed
    lead comes back with an error and a usable generic email rather than
    stopping the batch."""
    tag = lead.get("name") or lead.get("email") or f"row {lead['row']}"
    print(f"  [{lead['row']}] {tag}")

    site = fetch_site_text(lead.get("website", ""))
    if not site["ok"]:
        print(f"    site: {site['error']}")

    text_data = _build_text_data(lead, site)
    niche = resolve_niche(text_data)
    first = _first_name_for(lead)

    opener = ""
    hero_bytes = None
    err = site.get("error", "")
    if site["ok"]:
        # Same screenshot doubles as the vision model's input AND the hero
        # image embedded in the sent email — one screenshot, two uses, no
        # extra page load.
        hero_bytes = screenshot_site(site["url"], site.get("gallery_url", ""))
        try:
            images = [(hero_bytes, "image/jpeg")] if hero_bytes else None
            # Every HU guard runs inside this call via the patched
            # CUSTOM_OPENER_POSTPROCESS — a rejected opener comes back as the
            # safe "Rátaláltam az oldalatokra <városban>." fallback, not as junk.
            opener = _call_gemini_throttled(text_data, images)
        except Exception as e:
            err = f"opener: {str(e)[:120]}"
            print(f"    {err}")

    # Two cases land here: the API call raised (429/timeout -> opener is ""),
    # or a guard rejected the AI's draft and CUSTOM_OPENER_POSTPROCESS already
    # substituted dm_app_hu's own generic "Rátaláltam..." string in its place
    # (opener is non-empty but still generic). Either way, try the zero-API
    # deterministic tier first — it costs nothing and is strictly more
    # personalized than a bare city name whenever it finds something.
    # Matches on "Rátaláltam" alone, NOT the full string: dm_app_hu's fallback
    # was "Rátaláltam az oldalatokra ..." (tegező) and is now "Rátaláltam az
    # oldalukra ..." (magázás). Pinning the old wording here would have made
    # this check silently stop firing, passing the DM channel's generic opener
    # straight into a sent email instead of trying the deterministic tier below.
    if not opener or opener.strip().startswith("Rátaláltam"):
        det = _deterministic_opener(niche, site.get("text", ""))
        if det:
            opener = det
        else:
            # Covers BOTH cases that land here: opener was empty (API down),
            # or opener was dm_app_hu's own generic fallback string — which
            # is tegező ("...oldalatokra...") because that shared postprocess
            # also serves the tegező FB-DM channel. Always overwrite with our
            # own magázás version rather than only overwriting when opener
            # was empty; leaving the tegező string in place whenever det was
            # also empty was a real leak of the wrong register into a sent
            # email.
            city = dm_app_hu._city_to_ragos(lead.get("city", ""))
            opener = (f"Rátaláltam az Ön weboldalára {city}." if city
                      else "Rátaláltam az Ön weboldalára.")

    lead["opener"]  = opener
    lead["niche_resolved"] = niche

    # ADDRESS CORROBORATION — free, because the page text is already here.
    #
    # A 42% bounce day traced back to addresses that are plausible but were
    # never verified against anything: every domain in the list has valid MX,
    # so nothing DNS-level can catch them, and the failure only shows up as
    # "recipient not found" after the send has already cost reputation.
    # The page the opener was written from is the one place that can settle
    # it for free — if the company prints this address, it exists.
    #
    # Asymmetric on purpose: finding it is proof, NOT finding it is only a
    # hint. Plenty of real addresses live on a /kapcsolat page this scraper
    # never fetched, or are rendered as an image, or written as
    # "info [kukac] ceg.hu" to dodge harvesters. So this reports which of the
    # two it is and lets the human decide, exactly like _address_risk does.
    site_text_l = (site.get("text", "") or "").lower()
    addr = (lead.get("email", "") or "").strip().lower()
    if not addr or not site["ok"]:
        lead["email_on_site"] = "unknown"
    elif addr in site_text_l:
        lead["email_on_site"] = "confirmed"
    else:
        local, _, dom = addr.partition("@")
        # Catch the common anti-harvester spellings before calling it missing.
        obfuscated = any(v in site_text_l for v in (
            f"{local} kukac {dom}", f"{local}[kukac]{dom}", f"{local} (kukac) {dom}",
            f"{local}@ {dom}", f"{local} @{dom}", f"{local}(at){dom}",
            f"{local} at {dom}"))
        lead["email_on_site"] = "confirmed" if obfuscated else "not_found"

    # Same zero-API, regex/keyword extraction the deterministic opener tier
    # uses (a real brand/material name or a founding-year claim actually
    # printed on the lead's own page) — computed independently of which
    # opener tier actually fired, so the subject line gets a concrete,
    # page-sourced fact even on leads where the AI wrote a good photo-based
    # opener that this text-only regex wouldn't have found on its own. See
    # the 'hook' section of _hu_subject's docstring for why this beats a
    # niche label as a subject-line specificity signal.
    hook = (_find_niche_brand_mention(niche, site.get("text", ""))
            or _find_years_in_business(site.get("text", "")))

    # hero_bytes still exists above (it's what the vision model just looked
    # at to write the opener) but is deliberately NOT passed to hero_img=
    # here — the screenshot is no longer attached to the outgoing email.
    # Two reasons: an image-heavy first-touch cold email reads as more
    # "marketing" to spam filters than plain text, and the personalization
    # value it bought (a concrete opener) already landed via the vision call
    # above — attaching the photo itself was mostly decorative on top of
    # that. hero_img=False here always selects each pool's plain-text
    # subject/body variants.
    lead["subject"] = _hu_subject(first, niche, hero_img=False, hook=hook)
    lead["body"]    = _hu_email_body(opener, first, niche, hero_img=False,
                                     city=lead.get("city", ""))
    # No longer forwarded to hu_email_app.py's /send route — see above.
    # Kept as None (not deleted) so any code still reading this key doesn't
    # need a None-check added; it just always gets nothing to attach.
    lead["hero_image"] = None
    lead["error"]   = err
    print(f"    -> [{niche}] {opener[:70]}")
    return lead


# ─── EXCEL I/O ─────────────────────────────────────────────────────────────────

# Fuzzy header matching, same approach as dm_app._load_excel. Each entry is
# (key, [substrings]) checked in order, so the most specific match wins — 'email'
# is tested before the looser 'contact', and the company-name match must not
# swallow a separate owner-name column.
_COL_RULES = [
    ("email",   ["email", "e-mail", "mail"]),
    ("website", ["website", "weboldal", "honlap", "url", "web", "site"]),
    ("owner",   ["owner", "contact name", "tulajdonos", "kapcsolattartó", "személy"]),
    ("name",    ["company", "business", "cég", "cegnev", "cégnév", "name", "név"]),
    ("city",    ["city", "város", "varos", "location", "telephely", "megye"]),
    ("niche",   ["niche", "trade", "szakma", "kategória", "kategoria", "tevékenység"]),
]


def read_leads(path: str) -> tuple:
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    headers = [str(c.value or "").strip().lower() for c in ws[1]]

    cols = {}
    for key, needles in _COL_RULES:
        for i, h in enumerate(headers):
            if i in cols.values():
                continue
            if any(n in h for n in needles):
                cols[key] = i
                break

    missing = [k for k in ("email",) if k not in cols]
    if missing:
        raise ValueError(
            f"Couldn't find a column for: {', '.join(missing)}.\n"
            f"Headers found: {headers}")

    print(f"Column map: " + ", ".join(f"{k}={headers[i]!r}" for k, i in cols.items()))

    leads = []
    for row in ws.iter_rows(min_row=2):
        def v(key):
            i = cols.get(key)
            return str(row[i].value or "").strip() if i is not None else ""
        if not v("email"):
            continue
        leads.append({
            "row":     row[0].row,
            "email":   v("email"),
            "website": v("website"),
            "owner":   v("owner"),
            "name":    v("name"),
            "city":    _clean_city(v("city")),
            "niche":   v("niche"),
        })
    return leads, headers


def write_output(leads: list, headers: list, out_path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Emailek"

    columns = ["Sor", "Email", "Cég", "Város", "Niche (Excel)", "Niche (felismert)",
               "Weboldal", "Nyitósor", "Tárgy", "Levél", "Hiba", "Elküldve?"]
    head_fill = PatternFill("solid", fgColor="1F4E79")
    for i, c in enumerate(columns, 1):
        cell = ws.cell(row=1, column=i, value=c)
        cell.fill = head_fill
        cell.font = Font(color="FFFFFF", bold=True, size=11)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for r, lead in enumerate(leads, 2):
        vals = [lead["row"], lead["email"], lead["name"], lead["city"], lead["niche"],
                lead.get("niche_resolved", ""), lead["website"], lead.get("opener", ""),
                lead.get("subject", ""), lead.get("body", ""), lead.get("error", ""), ""]
        for i, val in enumerate(vals, 1):
            cell = ws.cell(row=r, column=i, value=val)
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    for i, w in enumerate([6, 28, 26, 14, 16, 16, 30, 46, 26, 70, 24, 12], 1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    ws.freeze_panes = "A2"
    for r in range(2, len(leads) + 2):
        ws.row_dimensions[r].height = 150
    wb.save(out_path)


# ─── MAIN ──────────────────────────────────────────────────────────────────────

TEMPLATE_HEADERS = ["Email", "Company Name", "Website", "City", "Niche", "Owner Name"]
TEMPLATE_EXAMPLE = ["info@peldakft.hu", "Példa Tetőfedő Kft.", "peldakft.hu",
                    "Miskolc", "tetőfedő", "Kovács Gábor"]


def make_template(path: str = "leads_template.xlsx"):
    """Write an empty lead sheet with the exact headers this script expects."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Leads"
    for i, h in enumerate(TEMPLATE_HEADERS, 1):
        c = ws.cell(row=1, column=i, value=h)
        c.fill = PatternFill("solid", fgColor="1F4E79")
        c.font = Font(color="FFFFFF", bold=True)
    for i, v in enumerate(TEMPLATE_EXAMPLE, 1):
        ws.cell(row=2, column=i, value=v).font = Font(italic=True, color="888888")
    for i, w in enumerate([30, 30, 30, 16, 22, 22], 1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    ws.freeze_panes = "A2"
    wb.save(path)
    print(f"Template written: {path}\nRow 2 is an example — delete it before use.")


def main():
    if "--template" in sys.argv:
        make_template()
        return

    if not dm_app.API_KEY:
        print("GROQ_API_KEY is not set — check the .env file in this folder.")
        sys.exit(1)

    # --limit N processes only the first N leads, for a cheap smoke test on a
    # real list before committing the whole batch to Groq's token budget.
    limit = None
    if "--limit" in sys.argv:
        limit = int(sys.argv[sys.argv.index("--limit") + 1])

    # The shared engine's own log lines (facts extraction, prompt builder, every
    # guard rejection) print from inside dm_app with no lead identifier, so at 5
    # workers they interleave into something unreadable — a rejection for one
    # lead appears directly under another lead's niche line. "--workers 1" makes
    # the log linear for debugging a specific lead, at the cost of throughput.
    global MAX_WORKERS
    if "--workers" in sys.argv:
        MAX_WORKERS = int(sys.argv[sys.argv.index("--workers") + 1])

    args = [a for a in sys.argv[1:] if not a.startswith("--") and not a.isdigit()]
    path = args[0] if args else input("Leads Excel path: ").strip().strip('"')
    if not os.path.isfile(path):
        print(f"File not found: {path}")
        sys.exit(1)

    leads, headers = read_leads(path)
    if limit:
        leads = leads[:limit]
    print(f"\n{len(leads)} lead betöltve. Generálás {MAX_WORKERS} szálon...\n")
    started = time.time()

    done = []
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
        futures = {pool.submit(process_lead, l): l for l in leads}
        for fut in as_completed(futures):
            lead = futures[fut]
            try:
                done.append(fut.result())
            except Exception as e:
                traceback.print_exc()
                lead["error"] = f"fatal: {str(e)[:150]}"
                done.append(lead)

    # Browsers are per-thread, so they have to be closed from their own thread.
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
        for _ in range(MAX_WORKERS):
            pool.submit(_close_browser)

    done.sort(key=lambda l: l["row"])
    base, _ = os.path.splitext(path)
    out = f"{base}_EMAILEK_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    write_output(done, headers, out)

    generic = sum(1 for l in done if l.get("opener", "").startswith("Rátaláltam"))
    mins = (time.time() - started) / 60
    print(f"\nKész: {len(done)} levél {mins:.1f} perc alatt")
    print(f"Általános nyitósor (fallback): {generic}/{len(done)} "
          f"({generic / max(len(done), 1) * 100:.0f}%)")
    print(f"Fájl: {out}")


if __name__ == "__main__":
    main()

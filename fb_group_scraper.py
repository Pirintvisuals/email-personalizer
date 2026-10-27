"""
fb_group_scraper.py — Facebook group lead scraper (company owners only).

Connects to the ALREADY-LOGGED-IN Chrome (remote debugging port 9222), opens a
Hungarian Facebook group, scrolls the feed slowly, and writes the people who
read as company owners / tradespeople advertising their own work into an Excel
sheet that dm_app_hu.py can load directly.

THE ONE RULE: never generate, guess, or recall a lead. Every name, URL and phone
number written to the output must have been read off a page that was actually
loaded in the browser. If something can't be read, the cell stays empty. Every
kept row records WHICH signal on the page caused it to be kept ("Miert lead"),
so the sheet can always be audited back to real page content.

Usage:
    python fb_group_scraper.py --group "https://www.facebook.com/groups/XXXX" --out leads_fb.xlsx
    python fb_group_scraper.py --group URL1 --group URL2 --out leads_fb.xlsx --max 300 --minutes 15
    python fb_group_scraper.py --group URL --out leads_fb.xlsx --all --debug
"""

import argparse
import random
import re
import shutil
import socket
import subprocess
import sys
import tempfile
import time
from pathlib import Path

import os

# Same convention as dm_app.py, so an alternate port can be used for testing.
CDP_PORT = int(os.environ.get("CDP_PORT", "9222"))

# ── Safety limits ────────────────────────────────────────────────────────
DEFAULT_MAX_PEOPLE = 300
DEFAULT_MAX_MINUTES = 25   # raised from 15: hitting a 30-lead target needs it
SCROLL_PAUSE_MIN = 2.0
SCROLL_PAUSE_MAX = 5.0
SCROLL_PX_MIN = 700
SCROLL_PX_MAX = 1500
DEFAULT_TARGET_LEADS = 30           # keep scrolling until this many leads
# Stop only when the feed stops producing NEW POSTS this many rounds in a row.
# Deliberately generous: Facebook is often slow to lazy-load the next batch.
FEED_STALL_LIMIT = 12
GROUP_PAUSE_MIN = 25.0
GROUP_PAUSE_MAX = 55.0
ARTICLE_WAIT_SECONDS = 25           # how long to wait for the feed to render

# ── Non-lead / directory pages to always drop ───────────────────────────
EXCLUDE_URL_SUBSTRINGS = [
    "facebook.com/joszaki",
    "facebook.com/aranyoldalak.hu",
    "facebook.com/szakmaikalauz",
    "facebook.com/people",
    "/pages/create",
    "/policy",
    "/help",
    "sharer",
    "/plugins",
    "dialog",
    "login",
]

RESERVED_HANDLES = {
    "groups", "watch", "marketplace", "gaming", "events", "photo",
    "photo.php", "story.php", "permalink.php", "sharer", "sharer.php",
    "pages", "plugins", "login", "help", "policy", "notifications",
    "messages", "ads", "business", "settings", "people", "profile.php",
    "reel", "reels", "video.php", "hashtag", "bookmarks", "friends",
    "privacy", "legal", "terms", "home.php", "search",
}

# ── "Is this a company owner?" signals ──────────────────────────────────
# Each list is a separate, auditable reason to keep a row.

OFFER_PHRASES = [
    "vállalok", "vállalunk", "vállaljuk", "kivitelezés", "kivitelezést",
    "ajánlatot adok", "ajánlatot adunk", "megrendelhető", "elérhető vagyok",
    "munkát keresek", "szabad kapacitás", "bemutatom a munkám", "referencia",
    "referenciáim", "referenciák", "árajánlat", "árajánlatot",
    "keressen bizalommal", "keressetek bizalommal", "hívjon bizalommal",
    "hívjatok bizalommal", "megbízható", "precíz munka", "garanciával",
    "több éves tapasztalat", "éves tapasztalattal", "munkáimból",
    "munkáinkból", "szolgáltatásaink", "cégünk", "vállalkozásom",
    "vállalkozásunk", "csapatunk", "bérelhető", "kiadó gép",
]

TRADE_WORDS = [
    "burkoló", "burkolás", "gázszerelő", "gázszerelés", "tetőfedő",
    "tetőfedés", "festő", "festés", "mázoló", "villanyszerelő",
    "villanyszerelés", "kőműves", "épületgépész", "fürdőszoba felújítás",
    "vízvezeték", "vízszerelő", "vízszerelés", "gipszkarton", "asztalos",
    "ács", "bádogos", "hegesztő", "klímaszerelés", "klímatelepítés",
    "szigetelés", "homlokzat", "térkövezés", "térkő", "parketta",
    "laminált padló", "álmennyezet", "felújítás", "generálkivitelezés",
    "lakásfelújítás", "házépítés", "redőnyös", "nyílászáró", "kertépítés",
    "földmunka", "betonozás", "vakolás", "glettelés",
]

BUSINESS_NAME_MARKERS = [
    " kft", " bt", " zrt", " e.v", " ev.", "épít", "szerel", "burkol",
    "festés", "tető", "víz-gáz", "gipszkarton", "kőműves", "asztalos",
    "villany", "klíma", "fürdőszoba", "felújít", "kivitelez", "studio",
    "stúdió", "design", "szerviz", "és fia", "trans", "team", "bau",
]

MEGRENDELO_PHRASES = [
    "keresek valakit", "tudna valaki ajánlani", "kit ajánlotok",
    "tudnátok ajánlani", "mennyibe kerül", "segítséget kérek",
    "mennyiért", "mibe kerül", "ajánlanátok", "van valakinek ötlete",
    "kérdésem lenne", "sürgősen keresek", "keresnék egy",
]

# Hungarian-ish phone number: optional +36/06, then 8-9 digits with separators.
PHONE_RE = re.compile(
    r"(?:(?:\+?36|06)[\s\-/.]?)?(?:\(?\d{1,2}\)?[\s\-/.]?)\d{3}[\s\-/.]?\d{3,4}"
)
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
WEBSITE_RE = re.compile(r"(?:https?://|www\.)[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", re.I)

CHECKPOINT_INDICATORS = [
    "checkpoint", "gyanús bejelentkezés", "erősítsd meg, hogy te vagy",
    "confirm your identity", "nem vagy robot", "captcha",
    "jelentkezz be újra", "you must log in", "gyanús tevékenység",
    "fiókod ideiglenesen", "biztonsági ellenőrzés",
]

LOGIN_WALL_INDICATORS = [
    "jelentkezz be a facebookra", "log in to facebook",
    "csatlakozz a csoporthoz", "join group to see",
    "a tartalom megtekintéséhez", "ehhez a csoporthoz nem",
]

OUTPUT_HEADERS = [
    "Name", "Facebook URL", "Group", "Group URL",
    "Phone", "Niche", "Miert lead", "Post date", "Post snippet", "Looks like",
]

# Extract post/comment authors. Anchored on role="article" + profile-link URL
# shapes, never on Facebook's obfuscated class names (they rotate).
JS_EXTRACT = r"""
() => {
    const RESERVED = new Set(['groups','watch','marketplace','gaming','events',
        'photo','photo.php','story.php','permalink.php','sharer','sharer.php',
        'pages','plugins','login','help','policy','notifications','messages',
        'ads','business','settings','people','profile.php','reel','reels',
        'video.php','hashtag','bookmarks','friends','privacy','legal','terms',
        'home.php','search']);

    function personHref(href) {
        if (!href) return false;
        if (/facebook\.com\/groups\/[^/]+\/user\/\d+/.test(href)) return true;
        if (/profile\.php\?id=\d+/.test(href)) return true;
        const m = href.match(/^https?:\/\/(?:www\.)?facebook\.com\/([A-Za-z0-9.\-]{2,})(?:[/?#]|$)/);
        if (m && !RESERVED.has(m[1].toLowerCase())) return true;
        return false;
    }

    function localText(a) {
        let el = a.parentElement;
        let text = '';
        for (let i = 0; i < 7 && el; i++) {
            text = el.innerText || '';
            if (text.length > 40) break;
            el = el.parentElement;
        }
        return text.slice(0, 600);
    }

    const results = [];
    const articles = document.querySelectorAll('[role="article"]');
    articles.forEach(article => {
        const full = (article.innerText || '');
        // Post date: FB renders it as a link/abbr in the post header.
        let dateText = '';
        const tl = article.querySelector('a[href*="/posts/"], a[href*="permalink"], abbr');
        if (tl) dateText = (tl.getAttribute('aria-label') || tl.getAttribute('title') || tl.innerText || '').trim();

        const links = Array.from(article.querySelectorAll('a[href]'));
        const seenHref = new Set();
        let idx = 0;
        for (const a of links) {
            const href = a.href || '';
            if (!personHref(href)) continue;
            const name = (a.innerText || '').trim();
            if (!name || name.length > 70 || name.indexOf('\n') !== -1) continue;
            if (seenHref.has(href)) continue;
            seenHref.add(href);
            results.push({
                href: href,
                name: name,
                snippet: idx === 0 ? full.slice(0, 1500) : localText(a),
                date: idx === 0 ? dateText : '',
                is_author: idx === 0,
            });
            idx++;
        }
    });
    return { articles: articles.length, items: results };
}
"""


class CheckpointDetected(Exception):
    pass


class FeedNotFound(Exception):
    pass


# ── Chrome / CDP connection ─────────────────────────────────────────────

def check_port_open(timeout=1.5):
    try:
        s = socket.create_connection(("127.0.0.1", CDP_PORT), timeout=timeout)
        s.close()
        return True
    except Exception:
        return False


CHROME_CANDIDATES = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]


def find_chrome():
    import os
    for c in CHROME_CANDIDATES:
        if Path(c).exists():
            return c
    local = os.environ.get("LOCALAPPDATA", "")
    if local:
        c = Path(local) / r"Google\Chrome\Application\chrome.exe"
        if c.exists():
            return str(c)
    found = shutil.which("chrome")
    return found or None


def scraper_profile_dir():
    """A Chrome profile that belongs to this tool only.

    Chrome 136+ REFUSES to open a remote-debugging port when it runs on the
    default profile directory (measured on Chrome 151: default profile never
    opens the port, dedicated --user-data-dir opens it in ~2s). So the scraper
    gets its own profile. Bonus: a separate profile is a separate Chrome
    process, so the user's normal browser is never touched or closed.

    Kept in LOCALAPPDATA, never in OneDrive — a Chrome profile is large and
    changes constantly, which would thrash the sync client.
    """
    import os
    base = os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()
    p = Path(base) / "fb_group_scraper_chrome"
    p.mkdir(parents=True, exist_ok=True)
    return str(p)


def ensure_chrome_debug():
    """Make sure a debug-enabled Chrome is reachable, printing every step.

    Never closes the user's normal Chrome — it launches a second Chrome on a
    dedicated profile, which coexists with it (verified).
    """
    if check_port_open():
        print(f"A gyujto Chrome mar fut ({CDP_PORT} port nyitva).")
        return True

    chrome = find_chrome()
    if not chrome:
        print("Nem talalom a chrome.exe-t a szokasos helyeken.")
        return False

    profile = scraper_profile_dir()
    first_run = not (Path(profile) / "Default").exists()

    print("A gyujto sajat Chrome ablakat inditom (a rendes Chrome-od nem zarom be).")
    print(f"  profil: {profile}")
    try:
        subprocess.Popen(
            [chrome,
             f"--user-data-dir={profile}",
             f"--remote-debugging-port={CDP_PORT}",
             "--no-first-run",
             "--no-default-browser-check",
             "about:blank"],
            close_fds=True)
    except Exception as e:
        print(f"Nem tudtam elinditani a Chrome-ot: {e}")
        return False

    print("Varom, hogy a debug port kinyiljon", end="")
    for i in range(40):
        if check_port_open():
            print(f"\nKesz — a {CDP_PORT} port nyitva ({i + 1} masodperc alatt).")
            time.sleep(1.5)
            if first_run:
                print("\n  >> Ez a gyujto Chrome profil most eloszor indult, tehat")
                print("  >> MEG NEM vagy bejelentkezve a Facebookra benne.")
                print("  >> Mindjart megnyitom a Facebookot - jelentkezz be ott EGYSZER,")
                print("  >> utana a program magatol folytatja, es tobbe nem kell.")
            return True
        print(".", end="", flush=True)
        time.sleep(1)

    print("\nA Chrome 40 masodperc alatt sem nyitotta meg a debug portot.")
    print("Probald kezzel: zarj be minden Chrome ablakot es futtasd ujra.")
    return False


FB_LOGIN_WAIT_MINUTES = 10


def has_fb_session(context):
    """True when the profile holds a logged-in Facebook session.

    Checks the `c_user` cookie, which Facebook only sets for an authenticated
    user. Deliberately cookie-based and NOT DOM-based: the login page must
    never be reloaded or scripted while the user is typing into it.
    """
    try:
        for c in context.cookies("https://www.facebook.com/"):
            if c.get("name") == "c_user" and str(c.get("value") or "").strip():
                return True
    except Exception:
        pass
    return False


def ensure_facebook_login(context):
    """Make sure the scraper's Chrome profile has a live Facebook session.

    On a brand-new profile it won't, so open the login page once and then wait
    passively, polling only the cookie jar. The page is never reloaded or
    touched again — an earlier version reloaded it every few seconds and wiped
    what the user was typing.
    """
    if has_fb_session(context):
        print("Facebook: be vagy jelentkezve a gyujto Chrome-ban.")
        return True

    page = None
    try:
        page = context.new_page()
        page.goto("https://www.facebook.com/", wait_until="domcontentloaded",
                  timeout=45000)
    except Exception as e:
        print(f"Nem tudtam megnyitni a facebook.com-ot: {e}")

    print("\n" + "=" * 60)
    print("  JELENTKEZZ BE A FACEBOOKRA a most megnyilt Chrome ablakban.")
    print("  Ezt csak EGYSZER kell megtenni, utana emlekszik ra.")
    print("  Nyugodtan irj - NEM tolti ujra az oldalt, nem torlom a beirtakat.")
    print("  (Jelszot en soha nem irok be, azt te teszed.)")
    print("=" * 60)

    deadline = time.time() + FB_LOGIN_WAIT_MINUTES * 60
    while time.time() < deadline:
        # Poll ONLY the cookie jar — never reload, never script the page.
        if has_fb_session(context):
            print("\n\nKesz, be vagy jelentkezve. Folytatom a gyujtest.\n")
            time.sleep(2)
            return True
        left = int(deadline - time.time())
        print(f"  varok a bejelentkezesre... ({left // 60}:{left % 60:02d})   ",
              end="\r", flush=True)
        time.sleep(3)

    print(f"\n\n{FB_LOGIN_WAIT_MINUTES} perc alatt nem tortent meg a bejelentkezes.")
    print("Jelentkezz be a gyujto Chrome ablakban, majd futtasd ujra a bat fajlt.")
    print("(A bejelentkezes megmarad, szoval legkozelebb mar nem fog kerdezni.)")
    return False


# ── URL normalisation / filtering ───────────────────────────────────────

def normalize_profile_url(href):
    """Strip tracking junk (__cft__, __tn__, eid, ...) and canonicalise."""
    if not href:
        return None
    href = href.split("#")[0]

    m = re.search(r"facebook\.com/groups/[^/]+/user/(\d+)", href)
    if m:
        return f"https://www.facebook.com/profile.php?id={m.group(1)}"

    m = re.search(r"profile\.php\?id=(\d+)", href)
    if m:
        return f"https://www.facebook.com/profile.php?id={m.group(1)}"

    m = re.match(r"https?://(?:www\.)?facebook\.com/([A-Za-z0-9.\-]{2,})(?:[/?]|$)", href)
    if m:
        handle = m.group(1)
        if handle.lower() in RESERVED_HANDLES:
            return None
        return f"https://www.facebook.com/{handle}"

    return None


def is_excluded_url(url):
    lu = url.lower()
    return any(s in lu for s in EXCLUDE_URL_SUBSTRINGS)


def clean_group_url(url):
    """Point a group link at the FEED, which is the only place post text lives.
    /members, /about, /media etc. have no post text to judge people by, so a
    link ending in one of those is silently corrected to the group root."""
    u = (url or "").strip().strip('"').strip("'")
    if not u:
        return ""
    if not u.startswith("http"):
        u = "https://" + u.lstrip("/")
    u = u.split("?")[0].split("#")[0].rstrip("/")
    m = re.match(r"(https?://(?:www\.|web\.|m\.)?facebook\.com/groups/[^/]+)(/.*)?$", u, re.I)
    if m:
        base = m.group(1)
        # always normalise the host to www so the desktop DOM renders
        base = re.sub(r"^https?://(?:www\.|web\.|m\.)?facebook\.com",
                      "https://www.facebook.com", base, flags=re.I)
        return base
    return u


# ── Company-owner detection (all signals read off the page) ─────────────

def find_phone(text):
    """Return the first plausible phone number found in the text, or ''."""
    for m in PHONE_RE.finditer(text or ""):
        raw = m.group(0)
        digits = re.sub(r"\D", "", raw)
        # Hungarian mobile/landline lengths, with or without country code.
        if 8 <= len(digits) <= 12:
            return raw.strip()
    return ""


def owner_signals(text, name):
    """Which auditable signals say 'this is a company owner / tradesperson'.
    Returns a list of short Hungarian reason strings (empty = no signal)."""
    t = (text or "").lower()
    n = (name or "").lower()
    reasons = []

    hit = next((p for p in OFFER_PHRASES if p in t), None)
    if hit:
        reasons.append(f'munkát hirdet ("{hit}")')

    trade = next((w for w in TRADE_WORDS if w in t), None)
    if trade:
        reasons.append(f"szakma a posztban ({trade})")

    if find_phone(text):
        reasons.append("telefonszám a posztban")

    if EMAIL_RE.search(text or ""):
        reasons.append("e-mail a posztban")

    if WEBSITE_RE.search(text or ""):
        reasons.append("weboldal a posztban")

    marker = next((b for b in BUSINESS_NAME_MARKERS if b in f" {n}"), None)
    if marker:
        reasons.append(f"céges név ({marker.strip()})")

    return reasons


def looks_like_customer(text):
    t = (text or "").lower()
    return any(p in t for p in MEGRENDELO_PHRASES)


def classify(text, name):
    """szakember = company owner/tradesperson, megrendelő = homeowner asking,
    bizonytalan = not enough evidence either way."""
    reasons = owner_signals(text, name)
    if reasons:
        # A customer asking for a recommendation can still mention a trade word.
        # Require more than a lone trade-word hit before overriding that.
        if looks_like_customer(text) and len(reasons) == 1 and reasons[0].startswith("szakma"):
            return "megrendelő", []
        return "szakember", reasons
    if looks_like_customer(text):
        return "megrendelő", []
    return "bizonytalan", []


def detect_niche(text, niche_flag):
    t = (text or "").lower()
    for w in TRADE_WORDS:
        if w in t:
            return w
    return niche_flag or ""


# ── Snippet cleanup ─────────────────────────────────────────────────────

_WS_RE = re.compile(r"\s+")


def clean_snippet(snippet, name):
    t = snippet or ""
    if name:
        t = t.replace(name, " ")
    t = re.sub(
        r"(?i)\b(tetszik|hozzászólás|megosztás|válasz|like|comment|share|reply|"
        r"összes megjelenítése|see more|see all|további|több)\b", " ", t)
    return _WS_RE.sub(" ", t).strip()


# ── Page-state detection ────────────────────────────────────────────────

def _body_text(page, n=4000):
    try:
        return (page.evaluate(
            "() => document.body ? document.body.innerText : ''") or "")[:n].lower()
    except Exception:
        return ""


def is_checkpoint(page):
    try:
        u = page.url.lower()
    except Exception:
        u = ""
    if "checkpoint" in u or "facebook.com/login" in u:
        return True
    txt = _body_text(page)
    return any(k in txt for k in CHECKPOINT_INDICATORS)


def diagnose_no_feed(page):
    """Explain WHY no posts were found, instead of silently returning 0 rows."""
    txt = _body_text(page)
    for k in LOGIN_WALL_INDICATORS:
        if k in txt:
            return ("Ez a csoport nem olvasható ezzel a fiókkal (belépés/csatlakozás kell), "
                    "vagy nem vagy a tagja.")
    try:
        u = page.url.lower()
    except Exception:
        u = ""
    if "/members" in u:
        return ("A megadott link a TAGOK oldal, nem a csoport hírfolyama. "
                "Használd a csoport fő linkjét (a /members rész nélkül).")
    if not txt.strip():
        return "A csoport oldala üresen jött be (nem töltött be tartalom)."
    return ("Nem találtam posztokat a hírfolyamban. Vagy nem töltött be, vagy a Facebook "
            "megváltoztatta az oldal szerkezetét — futtasd --debug módban és nézd meg a dump fájlt.")


# ── Excel I/O ─────────────────────────────────────────────────────────────

def load_existing_urls(path):
    """Profile URLs already in a workbook, matched by fuzzy header like
    dm_app.py does. Used for cross-run dedupe and --exclude."""
    from openpyxl import load_workbook

    p = Path(path)
    if not p.exists():
        return set()
    try:
        wb = load_workbook(p, data_only=True, read_only=True)
        ws = wb.active
    except Exception as e:
        print(f"  Figyelem: {path} nem olvasható ({e})")
        return set()

    rows = ws.iter_rows(values_only=True)
    try:
        header_row = next(rows)
    except StopIteration:
        return set()
    headers = [str(h).strip().lower() if h is not None else "" for h in header_row]
    url_idx = next((i for i, h in enumerate(headers) if "url" in h or "facebook" in h), None)
    if url_idx is None:
        return set()

    urls = set()
    for raw in rows:
        if url_idx < len(raw) and raw[url_idx]:
            urls.add(str(raw[url_idx]).strip())
    try:
        wb.close()
    except Exception:
        pass
    return urls


def load_existing_rows(path):
    """Previously-scraped rows in OUR canonical format, for append."""
    from openpyxl import load_workbook

    p = Path(path)
    if not p.exists():
        return []
    try:
        wb = load_workbook(p, data_only=True)
        ws = wb.active
    except Exception as e:
        print(f"  Figyelem: a meglévő {path} nem olvasható ({e})")
        return []

    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    rows = []
    for raw in ws.iter_rows(min_row=2, values_only=True):
        row = {h: (raw[i] if i < len(raw) and raw[i] is not None else "")
               for i, h in enumerate(headers)}
        if any(str(v).strip() for v in row.values()):
            rows.append(row)
    return rows


def save_output(path, rows):
    """Write the sheet, retrying while it's locked by Excel/OneDrive."""
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Leads"
    ws.append(OUTPUT_HEADERS)
    for r in rows:
        ws.append([r.get(h, "") for h in OUTPUT_HEADERS])

    widths = {"A": 26, "B": 44, "C": 30, "D": 40, "E": 16,
              "F": 18, "G": 34, "H": 14, "I": 60, "J": 12}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"

    for _ in range(8):
        try:
            wb.save(path)
            return
        except PermissionError:
            time.sleep(0.5)
    raise PermissionError(
        f"Nem tudok írni ide: {path} — nyitva van Excelben (vagy a OneDrive még szinkronizál). "
        "Zárd be a táblázatot, várj pár másodpercet, és futtasd újra."
    )


# ── Core scrape loop ─────────────────────────────────────────────────────

def wait_for_feed(page, debug_path=None):
    """Wait until posts render. Raises FeedNotFound with a real reason."""
    deadline = time.time() + ARTICLE_WAIT_SECONDS
    while time.time() < deadline:
        try:
            n = page.evaluate("() => document.querySelectorAll('[role=\"article\"]').length")
        except Exception:
            n = 0
        if n:
            return n
        time.sleep(1.5)

    if debug_path:
        try:
            body = page.evaluate("() => document.body ? document.body.innerText : ''")
            Path(debug_path).write_text(
                f"URL: {page.url}\n\n--- BODY INNERTEXT ---\n{body[:20000]}",
                encoding="utf-8")
            print(f"  Debug dump: {debug_path}")
        except Exception:
            pass
    raise FeedNotFound(diagnose_no_feed(page))


def scroll_feed(page, strategy):
    """Scroll the feed so Facebook lazy-loads more posts.

    Rotates strategies because a single one is not reliable: the group feed
    sometimes scrolls the window, sometimes an inner container, and a stalled
    feed often unsticks with a keypress instead.
    """
    try:
        if strategy == 0:
            page.evaluate(
                "() => { const e = document.scrollingElement || document.documentElement;"
                " e.scrollTop = e.scrollHeight; }")
        elif strategy == 1:
            page.evaluate(f"() => window.scrollBy(0, {random.randint(SCROLL_PX_MIN, SCROLL_PX_MAX)})")
        elif strategy == 2:
            page.keyboard.press("PageDown")
        elif strategy == 3:
            page.mouse.wheel(0, random.randint(SCROLL_PX_MIN, SCROLL_PX_MAX))
        else:
            # last resort: scroll the tallest scrollable container on the page
            page.evaluate("""() => {
                let best = null, bestH = 0;
                document.querySelectorAll('div').forEach(d => {
                    if (d.scrollHeight > d.clientHeight + 200 && d.scrollHeight > bestH) {
                        bestH = d.scrollHeight; best = d;
                    }
                });
                if (best) best.scrollTop = best.scrollHeight;
            }""")
    except Exception:
        pass


def write_debug_dump(path, page, payload, note=""):
    """Dump exactly what was read off the page, so a zero-lead run is
    diagnosable without asking the user to run again."""
    try:
        items = (payload or {}).get("items") or []
        lines = [
            f"URL: {page.url}",
            f"note: {note}",
            f"articles: {(payload or {}).get('articles')}",
            f"items: {len(items)}",
            "",
            "Az alabbi az, amit a program tenylegesen kiolvasott az oldalrol.",
            "A 'snippet' az a szoveg, amibol eldonti, hogy szakember-e.",
            "=" * 60,
            "",
        ]
        for i, it in enumerate(items[:40]):
            sn = it.get("snippet") or ""
            lines.append(
                f"[{i}] author={it.get('is_author')} name={it.get('name')!r}\n"
                f"     href={it.get('href')!r}\n"
                f"     date={it.get('date')!r}\n"
                f"     snippet_len={len(sn)}\n"
                f"     snippet={sn[:900]!r}\n")
        try:
            body = page.evaluate("() => document.body ? document.body.innerText : ''")
        except Exception:
            body = ""
        lines += ["", "=" * 60, "--- TELJES OLDAL SZOVEG (elso 15000 karakter) ---", "",
                  body[:15000]]
        Path(path).write_text("\n".join(lines), encoding="utf-8")
        print(f"  Debug dump kiirva: {path}")
    except Exception as e:
        print(f"  (debug dump sikertelen: {e})")


def scrape_group(page, group_url, group_name, seen_urls, rows_out, args, deadline, stats,
                 debug_path=None):
    n_articles = wait_for_feed(page, debug_path)
    print(f"  Hírfolyam betöltve ({n_articles} poszt látszik). Gyűjtés indul...")
    print(f"  Cél: {args.target} lead. Addig görgetek, amíg új poszt töltődik be.")

    started_with = len(rows_out)
    last_payload = None
    max_articles = 0
    stall_rounds = 0
    strategy = 0
    scrolls = 0

    while True:
        if time.time() >= deadline:
            print("\n  Időkorlát elérve, csoport lezárva.")
            break
        if len(rows_out) >= args.max:
            print("\n  Létszámkorlát (--max) elérve, csoport lezárva.")
            break
        # Target counts across the whole run, not per group.
        if len(rows_out) - stats["run_start"] >= args.target:
            print(f"\n  Megvan a cél ({args.target} lead), leállok.")
            break
        if is_checkpoint(page):
            raise CheckpointDetected(
                "Facebook biztonsági ellenőrzés / bejelentkezési oldal jelent meg. Azonnal leállok — "
                "jelentkezz be kézzel a Chrome-ban, aztán indítsd újra."
            )

        try:
            payload = page.evaluate(JS_EXTRACT)
            last_payload = payload
            raw_items = payload.get("items") or []
        except Exception as e:
            print(f"  Kinyerés hiba ezen a görgetésen ({e}); folytatom.")
            raw_items = []

        new_count = 0
        for item in raw_items:
            url = normalize_profile_url(item.get("href", ""))
            if not url:
                continue
            if is_excluded_url(url):
                stats["skipped_filtered"] += 1
                continue
            if url in seen_urls:
                continue
            seen_urls.add(url)
            stats["people_seen"] += 1

            name = (item.get("name") or "").strip()
            if not name:
                continue

            raw_snippet = item.get("snippet") or ""
            label, reasons = classify(raw_snippet, name)
            stats[label] = stats.get(label, 0) + 1

            if not args.all and label != "szakember":
                continue

            cleaned = clean_snippet(raw_snippet, name)
            rows_out.append({
                "Name": name,
                "Facebook URL": url,
                "Group": group_name,
                "Group URL": group_url,
                "Phone": find_phone(raw_snippet),
                "Niche": detect_niche(raw_snippet, args.niche),
                "Miert lead": "; ".join(reasons),
                "Post date": (item.get("date") or "").strip()[:40],
                "Post snippet": cleaned[:300],
                "Looks like": label,
            })
            new_count += 1
            print(f"  [{len(rows_out)}] {name}  <- {'; '.join(reasons) or label}")

            if len(rows_out) >= args.max:
                break

        # Whether to keep going depends on whether NEW POSTS are still loading,
        # not on whether this batch produced leads. Most posts in a homeowner-
        # heavy group yield nobody; that is not the end of the feed. (An earlier
        # version stopped on "no new people" and quit after ~28 seconds.)
        n_articles = payload.get("articles") or 0
        if n_articles > max_articles:
            max_articles = n_articles
            stall_rounds = 0
        else:
            stall_rounds += 1

        scrolls += 1
        found = len(rows_out) - started_with
        if new_count == 0:
            print(f"  ...{scrolls}. görgetés | {n_articles} poszt | "
                  f"{stats['people_seen']} ember | {found} lead", end="\r", flush=True)

        if stall_rounds >= FEED_STALL_LIMIT:
            print(f"\n  A hírfolyam nem tölt be több posztot ({max_articles} poszt után) "
                  "— a végére értem.")
            break

        # Rotate scroll strategies once the feed looks stuck.
        strategy = 0 if stall_rounds == 0 else (stall_rounds % 5)
        scroll_feed(page, strategy)
        time.sleep(random.uniform(SCROLL_PAUSE_MIN, SCROLL_PAUSE_MAX))

    found = len(rows_out) - started_with
    print(f"\n  Csoport kész: {found} lead, {scrolls} görgetés, {max_articles} poszt átnézve.")

    # Auto-dump when a group yielded nothing, so the next step needs no re-run.
    if found == 0 and last_payload is not None:
        auto = debug_path or (str(Path(args.out).with_suffix("")) + "_debug.txt")
        print("  Egy leadet sem talaltam ebben a csoportban — kiirom, mit olvastam:")
        write_debug_dump(auto, page, last_payload,
                         note=f"group={group_name}; 0 lead; {max_articles} poszt")


def get_group_name(page):
    try:
        title = page.evaluate("() => document.title || ''") or ""
        title = re.sub(r"\s*[|·]\s*Facebook\s*$", "", title).strip()
        title = re.sub(r"^\(\d+\+?\)\s*", "", title).strip()
        if title and title.lower() != "facebook":
            return title
    except Exception:
        pass
    return "Ismeretlen csoport"


# ── Main ─────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(
        description="Facebook csoport lead gyűjtő — csak cégtulajdonosok (bejelentkezett Chrome-ból).")
    ap.add_argument("--group", action="append", required=True, help="Csoport link. Többször is megadható.")
    ap.add_argument("--out", required=True, help="Kimeneti .xlsx (dm_app_hu.py kompatibilis).")
    ap.add_argument("--target", type=int, default=DEFAULT_TARGET_LEADS,
                    help=f"Ennyi lead osszegyujtese utan leall (alapertelmezett {DEFAULT_TARGET_LEADS}).")
    ap.add_argument("--max", type=int, default=DEFAULT_MAX_PEOPLE)
    ap.add_argument("--minutes", type=float, default=DEFAULT_MAX_MINUTES)
    ap.add_argument("--exclude", default=None, help="Meglévő lead lista, aminek a linkjeit kihagyom.")
    ap.add_argument("--niche", default="")
    ap.add_argument("--all", action="store_true",
                    help="Mindenkit megtart (megrendelőket és bizonytalanokat is).")
    ap.add_argument("--debug", action="store_true",
                    help="Kiírja az első poszt nyers tartalmát egy dump fájlba.")
    ap.add_argument("--no-chrome-launch", action="store_true",
                    help="Ne inditsa/zarja be a Chrome-ot, csak hasznalja a mar nyitott debug portot.")
    args = ap.parse_args()

    if args.no_chrome_launch:
        if not check_port_open():
            print(f"A {CDP_PORT} port ZÁRVA — a Chrome nem debug módban fut.")
            sys.exit(2)
        print(f"A {CDP_PORT} port nyitva.")
    elif not ensure_chrome_debug():
        sys.exit(2)

    from playwright.sync_api import sync_playwright

    seen_urls = set(load_existing_urls(args.out))
    print(f"{len(seen_urls)} ember már a kimeneti fájlban van — őket kihagyom.")
    if args.exclude:
        excl = load_existing_urls(args.exclude)
        print(f"{len(excl)} embert kihagyok innen: {args.exclude}")
        seen_urls |= excl

    rows_out = load_existing_rows(args.out)
    new_rows_start = len(rows_out)

    stats = {"people_seen": 0, "skipped_filtered": 0,
             "szakember": 0, "megrendelő": 0, "bizonytalan": 0,
             "run_start": new_rows_start}
    problems = []

    deadline = time.time() + args.minutes * 60

    # /members, /about etc. carry no post text — silently point them at the feed.
    groups = []
    for g in args.group:
        cleaned_g = clean_group_url(g)
        if not cleaned_g:
            continue
        if cleaned_g != g.strip().rstrip("/"):
            print(f"Link javítva: {g}  ->  {cleaned_g}")
        if cleaned_g not in groups:
            groups.append(cleaned_g)

    if not groups:
        print("Nem adtál meg használható csoport linket.")
        sys.exit(1)

    pw = sync_playwright().start()
    try:
        browser = pw.chromium.connect_over_cdp(f"http://localhost:{CDP_PORT}")
        context = browser.contexts[0]
    except Exception as e:
        print(f"Nem tudok csatlakozni a Chrome-hoz: {e}")
        try:
            pw.stop()
        except Exception:
            pass
        sys.exit(2)

    if not args.no_chrome_launch and not ensure_facebook_login(context):
        try:
            pw.stop()
        except Exception:
            pass
        sys.exit(2)

    try:
        for gi, group_url in enumerate(groups):
            if time.time() >= deadline:
                print("Az össz-időkorlát elfogyott a következő csoport előtt.")
                break
            if len(rows_out) - new_rows_start >= args.max:
                print("Az össz-létszámkorlát elfogyott a következő csoport előtt.")
                break
            if len(rows_out) - new_rows_start >= args.target:
                print(f"Megvan a cél ({args.target} lead) — a többi csoportot már nem nyitom meg.")
                break

            print(f"\n=== Csoport {gi + 1}/{len(groups)}: {group_url} ===")
            page = context.new_page()
            try:
                try:
                    page.goto(group_url, wait_until="domcontentloaded", timeout=45000)
                except Exception as e:
                    msg = f"Nem tudtam megnyitni a csoportot ({e})"
                    print(f"  {msg}")
                    problems.append(f"{group_url}: {msg}")
                    continue

                time.sleep(random.uniform(2.5, 4.5))

                if is_checkpoint(page):
                    msg = ("Facebook biztonsági ellenőrzés / bejelentkezési oldal. "
                           "Jelentkezz be kézzel a Chrome-ban, aztán futtasd újra.")
                    print(f"  !! {msg}")
                    problems.append(msg)
                    break

                group_name = get_group_name(page)
                print(f"  Csoport neve: {group_name}")

                debug_path = None
                if args.debug:
                    debug_path = str(Path(args.out).with_suffix("")) + f"_debug_{gi + 1}.txt"

                try:
                    scrape_group(page, group_url, group_name, seen_urls, rows_out,
                                 args, deadline, stats, debug_path)
                except FeedNotFound as e:
                    print(f"  !! {e}")
                    problems.append(f"{group_name or group_url}: {e}")
                except CheckpointDetected as e:
                    print(f"  !! {e}")
                    problems.append(str(e))
                    break
            finally:
                try:
                    page.close()
                except Exception:
                    pass

            if gi < len(groups) - 1 and time.time() < deadline:
                pause = random.uniform(GROUP_PAUSE_MIN, GROUP_PAUSE_MAX)
                print(f"  Szünet {pause:.0f} másodperc a következő csoport előtt...")
                time.sleep(pause)
    finally:
        try:
            pw.stop()
        except Exception:
            pass

    new_rows = rows_out[new_rows_start:]
    order = {"szakember": 0, "bizonytalan": 1, "megrendelő": 2}
    new_rows.sort(key=lambda r: order.get(r.get("Looks like"), 3))
    final_rows = rows_out[:new_rows_start] + new_rows

    # Always write the sheet — even 0 new rows, so the file visibly exists.
    save_output(args.out, final_rows)

    print("\n=== Összegzés ===")
    print(f"Emberek megnézve a posztokban:   {stats['people_seen']}")
    print(f"  cégtulaj/szakember jelnek megfelelt: {stats['szakember']}")
    print(f"  megrendelőnek tűnt (kihagyva):       {stats['megrendelő']}")
    print(f"  nem volt elég jel (kihagyva):        {stats['bizonytalan']}")
    print(f"Katalógus/kacat oldal kihagyva:  {stats['skipped_filtered']}")
    print(f"ÚJ lead a táblázatban:           {len(new_rows)}")
    print(f"Fájl: {args.out}  (összesen {len(final_rows)} sor)")

    if problems:
        print("\n!! Problémák:")
        for p in problems:
            print(f"  - {p}")

    if not new_rows and not problems:
        print("\nNem találtam új cégtulajdonost. Ha ez váratlan, futtasd --debug módban,")
        print("és nézd meg a *_debug_1.txt fájlt — abból látszik, mit olvasott az oldalról.")


if __name__ == "__main__":
    main()

# Facebook Group lead scraper — build spec

Hand this whole file to a fresh chat. It is self-contained.

Working directory: `C:\Users\pirin\OneDrive\Mellékletek\personalized dm\`

---

## Goal

Build `fb_group_scraper.py`: connect to Milán's **already-logged-in Chrome**, open a
Hungarian Facebook group he is a member of, scroll it, and extract the real people
posting in it. Output an Excel that his existing DM sender (`dm_app_hu.py`) can load
directly.

He is a solo founder selling an AI quoting-assistant widget to Hungarian trade
businesses (bathroom fitters, gas/heating engineers, roofers, tilers, renovators).
Facebook DMs are his only channel that has produced paying clients. He ran out of
leads. This tool refills them.

---

## THE ONE RULE

**Never generate, guess, or recall a lead.** Every name and profile URL must be copied
off a page that was actually loaded in the browser.

This is not a style preference. His previous lead list was produced by asking an AI to
"find Hungarian tradespeople", and it was measured at roughly **50% fabricated** — 13% of
the websites did not resolve at all, and a real send day produced a **42% hard-bounce
rate**. Weeks of work were spent debugging email copy, subject lines and deliverability
for a problem that was entirely fake data.

If you cannot read something from the live page, leave the field empty. An empty cell is
useful. An invented one is worse than nothing.

---

## Existing infrastructure to reuse (do not rebuild)

### The logged-in Chrome
`Open Chrome for DM App.bat` launches Chrome with remote debugging on port **9222**.
`dm_app.py` connects to it like this (copy this pattern — it is known to work):

```python
import socket
from playwright.sync_api import sync_playwright

CDP_PORT = 9222
socket.create_connection(("127.0.0.1", CDP_PORT), timeout=3)   # fail fast if not open
pw = sync_playwright().start()
browser = pw.chromium.connect_over_cdp(f"http://localhost:{CDP_PORT}")
ctx = browser.contexts[0]
page = ctx.new_page()
```

See `dm_app.py` around lines 566–581 and 3229–3235 for the exact working usage,
including how it handles the port being closed.

**Do not launch a fresh browser.** A new Playwright browser is not logged into Facebook.
The whole point is to reuse his real session.

### Output format — what `dm_app_hu.py` reads
Column matching is fuzzy, by header substring (see `dm_app.py` ~line 2937). Produce a
sheet with these headers:

| Header | Required | Notes |
|---|---|---|
| `Facebook URL` | **yes** | matched by `"url"` or `"facebook"`. The profile/page URL. |
| `Name` | yes | matched by `"name"` or `"company"`. Person or business name. |
| `City` | no | matched by `"city"` or `"location"`. |
| `Niche` | no | matched by `"niche"` or `"trade"`. |

The DM app auto-creates `Contacted`, `Sent Date`, `Replied?` and `Follow-Ups` columns on
load — do not add them yourself.

Also write these extra columns for his own review (harmless, the DM app ignores them):
`Post snippet`, `Post date`, `Group`, `Looks like`.

---

## What to extract

For every **post author** and **comment author** in the group:

- **Name** — as displayed
- **Profile URL** — normalise to `https://www.facebook.com/<handle>` or
  `https://www.facebook.com/profile.php?id=<id>`. Strip all tracking query params
  (`__cft__`, `__tn__`, `eid`, etc.) — they are session-specific junk.
- **Post text snippet** — first ~200 chars. This is how he tells a tradesperson from a
  homeowner.
- **Post date** — relative ("3 napja") is fine. He prioritises people active this week.
- **Group name/URL** — provenance.

### Classify each person: `Looks like`
Cheap keyword pass over their post text, no AI needed:

- `szakember` (tradesperson — the target): text contains offers of work —
  `vállalok`, `vállalunk`, `kivitelezés`, `ajánlatot adok`, `megrendelhető`,
  `elérhető vagyok`, `munkát keresek`, `szabad kapacitás`, `bemutatom a munkám`,
  `referencia`, or a trade word (`burkoló`, `gázszerelő`, `tetőfedő`, `festő`,
  `villanyszerelő`, `kőműves`, `épületgépész`, `fürdőszoba felújítás`)
- `megrendelő` (homeowner — not the target): `keresek`, `tudna valaki ajánlani`,
  `kit ajánlotok`, `mennyibe kerül`, `segítséget kérek`
- `bizonytalan` — anything else

Both `szakember` and `bizonytalan` are worth keeping; sort `szakember` first. Never drop
rows silently — he decides.

---

## Behaviour requirements

### 1. Account safety — this is the highest priority
His Facebook account **is his business**. If it gets restricted, he loses the only channel
that has ever produced clients. Scraping speed is worth nothing compared to that.

- Scroll in small steps with randomised pauses (**2–5 seconds**), not a tight loop
- Hard cap per run: **~300 people** or **~15 minutes**, whichever first, then stop cleanly
- One group per run by default
- Never click anything (no reactions, no "see more" spam clicking, no friend requests)
- If a checkpoint/login/CAPTCHA page appears, **stop immediately** and tell him — do not
  attempt to solve or work around it

### 2. Resilience to Facebook's DOM
Facebook's class names are obfuscated and rotate. Do **not** depend on them.
Anchor on stable things instead:
- links matching `facebook.com/groups/<id>/user/<uid>/` or `/profile.php?id=` or a bare
  `/<handle>` inside a post header
- `role="article"` for post containers
- `aria-label` text where present

Assume the layout will break eventually. Fail loudly with a clear message rather than
silently returning 0 rows.

### 3. Deduplication
- Within a run: dedupe by profile URL
- Across runs: if the output file exists, load it and skip profile URLs already present
- Also skip anyone already contacted in his existing lead sheets if a path is given
  (optional flag `--exclude <xlsx>`)

### 4. Filter out non-leads
Drop these — they are aggregator/directory pages, not businesses. Learned the hard way
when a website-based scrape returned them as "leads":
`facebook.com/joszaki`, `facebook.com/aranyoldalak.hu`, `facebook.com/szakmaikalauz`,
`facebook.com/people`, and anything matching `/pages/create`, `/policy`, `/help`,
`sharer`, `/plugins`, `dialog`, `login`.

### 5. CLI

```bash
python fb_group_scraper.py --group "https://www.facebook.com/groups/XXXX" --out leads_fb.xlsx
```

Flags: `--max 300`, `--minutes 15`, `--exclude <xlsx>`, `--niche "burkoló"`.
Multiple `--group` allowed; process sequentially with a longer pause between groups.

Print a live count as it goes, and a summary at the end (found / new / szakember /
megrendelő / skipped).

---

## Verify before declaring done

1. `python -m py_compile fb_group_scraper.py`
2. Run it against one real group with `--max 20` and confirm:
   - profile URLs are real and open to a real profile when pasted in a browser
   - names are not empty
   - the `Looks like` classification is roughly sane
   - re-running skips everything already captured (dedupe works)
3. Load the output in `dm_app_hu.py` (port 5002) and confirm it reads the leads.
4. Report the real numbers. If it found 12 people, say 12 — do not round up or estimate.

---

## Context that will save you time

- Python is `python`, **not** `py` — `py.exe` on this machine is a Windows Store stub that
  exits silently doing nothing.
- Playwright is already installed and used by `dm_app.py` and `hu_email_gen.py`.
- `openpyxl` and `beautifulsoup4` are already installed.
- **Never edit `dm_app.py`, `dm_app_hu.py` or `dm_app_de.py`.** They are the working DM
  sender. Import from them if useful, but never write to them.
- Hungarian UI text for anything user-facing; code comments in English.
- Excel files may be locked by Excel/OneDrive — retry saves a few times before failing,
  with a clear message (see `_save_wb` in `dm_app.py` for the existing pattern).

---

## Known-good fallback source (build later if groups underdeliver)

Proven on his real data: **company website → their Facebook page link.**
Fetch a business's website, parse `<a href>` for `facebook.com/...`, take the first that
isn't in the filter list above. Measured **42% yield** on a 55-lead sample (~32% after
removing directory false-positives).

That needs a source of businesses with websites — Google Maps (city × trade) is the
obvious one, and has the side benefit that "has a website" is a product requirement,
since his widget needs a site to live on.

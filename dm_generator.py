"""
Cold DM Generator for Landscaping Outreach (Facebook / Instagram)
-----------------------------------------------------------------
Input:  Excel or .txt file with profile URLs (+ optional Notes column)
Output: Excel file with a ready-to-send DM per row

Usage:
  python dm_generator.py urls.txt
  python dm_generator.py leads.xlsx
  python dm_generator.py                  # prompts for file path

Cost: ~$0.00002 per DM (gemini-2.0-flash)
"""

import os, sys, re, time, random
from pathlib import Path
import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from google import genai

# ── Load .env if present ──────────────────────────────────────────────────────
_env = Path(__file__).parent / ".env"
if _env.exists():
    for line in _env.read_text().splitlines():
        if "=" in line and not line.startswith("#"):
            k, v = line.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip())

# ── Config ────────────────────────────────────────────────────────────────────
API_KEY = os.environ.get("GEMINI_API_KEY", "")
MODEL   = "gemini-2.5-flash"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-GB,en;q=0.9",
}

# ── Prompt (kept tight — every token costs) ───────────────────────────────────
DM_BODY = (
    " I know you're about a billion DMs deep from marketing lads promising you the world, "
    "so I'll skip the fluff. I made a short video showing how landscapers automatically filter out "
    "the shitty jobs so they only quote for jobs worth their time. No ads, no website, no bullshit. "
    "Easier to show than explain, want me to send it over?"
)

PROMPT_TEMPLATE = """\
You write the opening line of a cold Facebook DM to a UK landscaping business owner. Output ONLY the opening line — nothing else.

The opening line follows this format:
"Cheers [Name], [personalized observation]."

RULES FOR [Name]:
- Use the first name if visible in the profile data.
- If no first name is visible, write "Cheers, [personalized observation]." (no name).

RULES FOR [personalized observation]:
- ONE short specific thing from their page — a job, a project type, a location — something that proves you looked.
- 1 sentence only. Casual and direct. No compliments. No filler.
- Examples: "saw that resin driveway you posted, looks tidy" / "noticed you do a lot of block paving up in Leeds" / "that garden transformation you shared looks spot on"
- If nothing specific is visible, write: "came across your page"

STRICT RULES:
- Output ONLY the opening line. No explanation. No labels. No extra sentences.
- No em dashes. No exclamation marks. No emojis.

Profile data:
{data}"""

# ── Page scraper ──────────────────────────────────────────────────────────────
def scrape(url: str) -> str:
    """Try to get useful text from a public page. FB/IG often return login walls."""
    try:
        r = requests.get(url, headers=HEADERS, timeout=10, allow_redirects=True)
        soup = BeautifulSoup(r.text, "lxml")
        for tag in soup(["script", "style", "nav", "footer", "noscript", "svg"]):
            tag.decompose()

        # Check for login wall
        text = soup.get_text(separator=" ")
        text = re.sub(r"\s+", " ", text).strip()

        login_signals = ["log in to facebook", "sign in to instagram",
                         "you must log in", "create new account", "log into facebook"]
        if any(s in text.lower() for s in login_signals):
            return ""   # login wall — caller will handle

        return text[:1200]
    except Exception:
        return ""


# ── DM generator ─────────────────────────────────────────────────────────────
def generate_dm(client: genai.Client, page_data: str) -> str:
    prompt = PROMPT_TEMPLATE.format(
        data=page_data or "No page data available — write a generic but warm opener for a landscaping business."
    )
    response = client.models.generate_content(
        model=MODEL,
        contents=prompt,
        config={"max_output_tokens": 400, "temperature": 0.9},
    )
    return response.text.strip() + DM_BODY


# ── Input readers ─────────────────────────────────────────────────────────────
def read_txt(path: str) -> list[dict]:
    with open(path, encoding="utf-8") as f:
        lines = [l.strip() for l in f if l.strip() and not l.startswith("#")]
    return [{"URL": l, "Notes": ""} for l in lines]


def read_xlsx(path: str) -> list[dict]:
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    header = [str(c or "").strip().lower() for c in rows[0]]

    url_col   = next((i for i, h in enumerate(header) if "url" in h or "link" in h or "facebook" in h or "instagram" in h or "profile" in h), 0)
    notes_col = next((i for i, h in enumerate(header) if "note" in h or "detail" in h or "info" in h or "about" in h), None)
    name_col  = next((i for i, h in enumerate(header) if "name" in h), None)

    records = []
    for row in rows[1:]:
        if not row or all(c is None for c in row):
            continue
        url   = str(row[url_col] or "").strip()
        notes = str(row[notes_col] or "").strip() if notes_col is not None else ""
        name  = str(row[name_col]  or "").strip() if name_col  is not None else ""
        if url:
            records.append({"URL": url, "Notes": notes, "Name": name})
    return records


# ── Output writer ─────────────────────────────────────────────────────────────
def write_xlsx(records: list[dict], out_path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DMs"

    hdr_fill = PatternFill("solid", fgColor="1F4E79")
    hdr_font = Font(color="FFFFFF", bold=True)
    cols = ["URL", "Notes", "DM", "Status"]

    for ci, c in enumerate(cols, 1):
        cell = ws.cell(row=1, column=ci, value=c)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center")

    fill_a = PatternFill("solid", fgColor="D6E4F0")
    fill_b = PatternFill("solid", fgColor="FFFFFF")

    for ri, rec in enumerate(records, 2):
        fill = fill_a if ri % 2 == 0 else fill_b
        for ci, key in enumerate(cols, 1):
            cell = ws.cell(row=ri, column=ci, value=rec.get(key, ""))
            cell.fill = fill
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    ws.column_dimensions["A"].width = 45
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 70
    ws.column_dimensions["D"].width = 15
    ws.freeze_panes = "A2"

    wb.save(out_path)
    print(f"\n  Saved → {out_path}")


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    if not API_KEY:
        print("ERROR: GEMINI_API_KEY not set.")
        print("  Windows: set GEMINI_API_KEY=AIza...")
        sys.exit(1)

    client = genai.Client(api_key=API_KEY)

    if len(sys.argv) > 1:
        path = sys.argv[1]
    else:
        path = input("Path to URL list (.txt or .xlsx): ").strip().strip('"')

    if not os.path.isfile(path):
        print(f"File not found: {path}")
        sys.exit(1)

    ext = os.path.splitext(path)[1].lower()
    records = read_txt(path) if ext == ".txt" else read_xlsx(path)

    if not records:
        print("No URLs found in file.")
        sys.exit(1)

    print(f"Loaded {len(records)} profiles\n")

    scraped_ok = 0
    login_wall = 0

    for i, rec in enumerate(records, 1):
        url = rec["URL"]
        print(f"[{i}/{len(records)}] {url[:60]}")

        # Build context: prefer manual Notes, else try scraping
        if rec.get("Notes"):
            page_data = f"URL: {url}\nNotes: {rec['Notes']}"
            if rec.get("Name"):
                page_data = f"Name: {rec['Name']}\n" + page_data
            print(f"  using notes column")
        else:
            raw = scrape(url)
            if raw:
                page_data = f"URL: {url}\n\n{raw}"
                scraped_ok += 1
                print(f"  scraped {len(raw)} chars")
            else:
                page_data = f"URL: {url}\n(Page could not be scraped — login wall or error)"
                login_wall += 1
                print(f"  login wall / failed — DM will be generic")

        dm = generate_dm(client, page_data)
        rec["DM"] = dm
        rec["Status"] = ""
        print(f"  {dm}\n")

        if i < len(records):
            time.sleep(0.3)   # small pause to avoid rate limits

    # Summary
    print(f"Done. Scraped: {scraped_ok}  Login walls: {login_wall}  Notes-based: {len(records)-scraped_ok-login_wall}")

    # Save output
    base = os.path.splitext(path)[0]
    out_path = f"{base}_DMs.xlsx"
    write_xlsx(records, out_path)


if __name__ == "__main__":
    main()

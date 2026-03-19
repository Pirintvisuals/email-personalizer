"""
Landscaper Email Campaign Generator — Flask Web App
Uses Gemini 2.5 Flash for personalized cold email generation.
"""

import os
import re
import json
import time
import uuid
import threading
from datetime import datetime
from urllib.parse import urlparse

# Load .env file if present (no external dependency needed)
_env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
if os.path.isfile(_env_path):
    with open(_env_path) as _f:
        for _line in _f:
            _line = _line.strip()
            if _line and not _line.startswith("#") and "=" in _line:
                _k, _v = _line.split("=", 1)
                os.environ.setdefault(_k.strip(), _v.strip())

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from flask import (Flask, render_template, request, jsonify,
                   send_from_directory, redirect, url_for)
from google import genai

# ─── CONFIG ────────────────────────────────────────────────────────────────────

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
GEMINI_MODEL   = "gemini-2.5-flash"

BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR  = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR  = os.path.join(BASE_DIR, "outputs")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

REQUEST_TIMEOUT = 12
REQUEST_DELAY   = 1.0
API_DELAY       = 1.5

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}

# ─── FLASK APP ─────────────────────────────────────────────────────────────────

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB upload limit

gemini_client = genai.Client(api_key=GEMINI_API_KEY)

# In-memory job store  { job_id: { status, progress, total, log, output_file } }
jobs: dict[str, dict] = {}
jobs_lock = threading.Lock()


# ─── WEB SCRAPER ───────────────────────────────────────────────────────────────

def clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def fetch_website(url: str) -> dict:
    result = {
        "success": False, "url": url, "title": "", "body_text": "",
        "has_contact_form": False, "has_booking": False,
        "has_cta": False, "page_meta": "", "error": "",
    }
    if not url or url.strip() in ("", "N/A", "n/a", "-"):
        result["error"] = "No URL provided"
        return result
    if not url.startswith(("http://", "https://")):
        url = "https://" + url
    result["url"] = url

    try:
        resp = requests.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT, allow_redirects=True)
        resp.raise_for_status()
    except requests.exceptions.SSLError:
        try:
            resp = requests.get(url.replace("https://", "http://"),
                                headers=HEADERS, timeout=REQUEST_TIMEOUT, allow_redirects=True)
            resp.raise_for_status()
        except Exception as e:
            result["error"] = f"SSL + HTTP fallback failed: {str(e)[:120]}"
            return result
    except Exception as e:
        result["error"] = str(e)[:150]
        return result

    try:
        soup = BeautifulSoup(resp.text, "lxml")
        for tag in soup(["script", "style", "nav", "footer", "head", "noscript", "iframe", "svg"]):
            tag.decompose()
        title_tag = soup.find("title")
        result["title"] = clean_text(title_tag.get_text()) if title_tag else ""
        meta = soup.find("meta", attrs={"name": re.compile(r"description", re.I)})
        if meta and meta.get("content"):
            result["page_meta"] = clean_text(meta["content"])
        body = soup.find("body")
        raw = body.get_text(separator=" ") if body else soup.get_text(separator=" ")
        result["body_text"] = clean_text(raw)[:3000]
        forms = soup.find_all("form")
        result["has_contact_form"] = bool(forms)
        page_lower = resp.text.lower()
        result["has_booking"] = any(k in page_lower for k in
            ["book", "schedule", "appointment", "calendar", "reserve",
             "get a quote", "free quote", "request a quote", "estimate"])
        result["has_cta"] = any(k in page_lower for k in
            ["call us", "contact us", "get started", "free estimate",
             "call now", "get a quote", "request service"])
        result["success"] = True
    except Exception as e:
        result["error"] = f"Parse error: {str(e)[:120]}"
    return result


# ─── GEMINI EMAIL GENERATOR ────────────────────────────────────────────────────

def build_prompt(company: str, contact: str, site_data: dict) -> str:
    if site_data["success"]:
        site_summary = f"""
WEBSITE DATA:
- Title: {site_data['title']}
- Meta description: {site_data['page_meta']}
- Has contact/lead form: {site_data['has_contact_form']}
- Has booking / quote CTA: {site_data['has_booking']}
- Has call-to-action: {site_data['has_cta']}
- Page text excerpt:
{site_data['body_text'][:2500]}
"""
    else:
        site_summary = f"WEBSITE FETCH FAILED: {site_data['error']}"

    return f"""You are an expert B2B cold email copywriter for a lead-filtering service that helps landscaping companies stop wasting time on bad leads.

Company: {company}
Contact: {contact}
Website: {site_data['url']}

{site_summary}

YOUR TASK — produce a JSON object with exactly these keys:
1. "research_notes" — 2-4 sentences: what you observed on their site (services, lead capture quality, obvious gaps). If site failed to load, note it and describe what a typical landscaper site looks like.
2. "email_1" — initial cold outreach (100-150 words). Rules:
   - First line is "Subject: ..."
   - Reference something SPECIFIC from their website or business
   - Identify a lead capture / lead quality problem they likely have
   - Explain how we pre-qualify leads so they only talk to serious prospects
   - End with ONE CTA (e.g., "Worth a 15-min call this week?")
   - Tone: direct, peer-to-peer, no fluff
   - NO "I hope this email finds you well"
   - Sign off: [Your Name]
3. "email_2" — follow-up 1, send day 3 (~80-100 words):
   - Subject line first
   - Add a specific market insight: local competition, seasonal lead spikes, or stat about bad leads wasting crew time
   - One pain point, one nudge — no hard sell
4. "email_3" — follow-up 2, send day 8 (~80-100 words):
   - Subject line first
   - Open with a brief case study / success metric (e.g. "One landscaper cut quote-to-close time by 40% after filtering...")
   - Soft close referencing {company} specifically
   - This is the final touch

Return ONLY valid JSON. No markdown fences. No extra text. Make email_1 specific to THIS company."""


def generate_emails(company: str, contact: str, site_data: dict) -> dict:
    prompt = build_prompt(company, contact, site_data)
    try:
        response = gemini_client.models.generate_content(
            model=GEMINI_MODEL,
            contents=prompt,
        )
        raw = response.text.strip()
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw.strip())
        return json.loads(raw)
    except json.JSONDecodeError:
        # Gemini sometimes wraps in backticks — try harder
        try:
            match = re.search(r'\{.*\}', raw, re.DOTALL)
            if match:
                return json.loads(match.group())
        except Exception:
            pass
        return {
            "research_notes": f"JSON parse error. Raw: {raw[:300]}",
            "email_1": "ERROR", "email_2": "ERROR", "email_3": "ERROR",
        }
    except Exception as e:
        return {
            "research_notes": f"API error: {str(e)[:200]}",
            "email_1": "ERROR", "email_2": "ERROR", "email_3": "ERROR",
        }


# ─── EXCEL I/O ─────────────────────────────────────────────────────────────────

def get_excel_headers(path: str) -> dict:
    """Return headers and a sample row for column mapping UI."""
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError("File is empty")
    headers = [str(c).strip() if c is not None else "" for c in rows[0]]
    sample  = [str(c).strip() if c is not None else "" for c in (rows[1] if len(rows) > 1 else [])]
    # Auto-detect best guess for each field
    guesses = {}
    for i, h in enumerate(headers):
        hl = h.lower()
        if any(k in hl for k in ["company", "business", "name"]) and "contact" not in hl:
            guesses.setdefault("company", i)
        elif any(k in hl for k in ["contact", "first name", "person"]):
            guesses.setdefault("contact", i)
        elif any(k in hl for k in ["website", "url", "web", "site"]):
            guesses.setdefault("website", i)
        elif any(k in hl for k in ["phone", "tel", "mobile", "cell"]):
            guesses.setdefault("phone", i)
    return {"headers": headers, "sample": sample, "guesses": guesses, "total": len(rows) - 1}


def read_excel(path: str, col_map: dict) -> list[dict]:
    """Read excel with an explicit col_map {company, contact, website, phone} -> int index."""
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    records = []
    for row in rows[1:]:
        if all(c is None or str(c).strip() == "" for c in row):
            continue
        def safe(idx):
            if idx is None or idx < 0: return ""
            return str(row[idx]).strip() if idx < len(row) and row[idx] is not None else ""
        records.append({
            "Company Name": safe(col_map.get("company")),
            "Contact Name": safe(col_map.get("contact")),
            "Website URL":  safe(col_map.get("website")),
            "Phone Number": safe(col_map.get("phone")),
        })
    return records


def write_excel(records: list[dict], output_path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Email Campaign"
    columns = [
        "Company Name", "Contact Name", "Website URL", "Phone Number",
        "Research Notes", "Email 1 (Initial)", "Email 2 (Follow-up Day 3)",
        "Email 3 (Follow-up Day 8)", "Status"
    ]
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    for ci, col in enumerate(columns, 1):
        cell = ws.cell(row=1, column=ci, value=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    fill_e = PatternFill("solid", fgColor="D6E4F0")
    fill_o = PatternFill("solid", fgColor="FFFFFF")
    for ri, rec in enumerate(records, 2):
        fill = fill_e if ri % 2 == 0 else fill_o
        for ci, val in enumerate([
            rec.get("Company Name",""), rec.get("Contact Name",""),
            rec.get("Website URL",""),  rec.get("Phone Number",""),
            rec.get("Research Notes",""), rec.get("Email 1",""),
            rec.get("Email 2",""),        rec.get("Email 3",""),
            rec.get("Status",""),
        ], 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.fill = fill
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    for ci, w in enumerate([28,20,35,18,50,70,70,70,15], 1):
        ws.column_dimensions[ws.cell(row=1,column=ci).column_letter].width = w
    ws.row_dimensions[1].height = 30
    for ri in range(2, len(records)+2):
        ws.row_dimensions[ri].height = 120
    ws.freeze_panes = "A2"
    wb.save(output_path)


# ─── BACKGROUND WORKER ─────────────────────────────────────────────────────────

def process_job(job_id: str, input_path: str, col_map: dict):
    def log(msg: str):
        with jobs_lock:
            jobs[job_id]["log"].append(msg)

    def set_progress(n: int):
        with jobs_lock:
            jobs[job_id]["progress"] = n

    try:
        records = read_excel(input_path, col_map)
        total = len(records)
        with jobs_lock:
            jobs[job_id]["total"] = total
        log(f"✅ Loaded {total} landscapers from file")

        for i, rec in enumerate(records):
            company = rec["Company Name"] or f"Landscaper #{i+1}"
            contact = rec["Contact Name"] or "there"
            url     = rec["Website URL"]

            log(f"[{i+1}/{total}] {company}")
            log(f"  🌐 Fetching: {url or '(no URL)'}")

            site_data = fetch_website(url)
            if site_data["success"]:
                log(f"  ✅ Site loaded — form:{site_data['has_contact_form']} booking:{site_data['has_booking']}")
            else:
                log(f"  ⚠️  Site failed: {site_data['error']}")

            time.sleep(REQUEST_DELAY)

            log(f"  🤖 Generating emails with Gemini...")
            output = generate_emails(company, contact, site_data)

            rec["Research Notes"] = output.get("research_notes", "")
            rec["Email 1"]        = output.get("email_1", "")
            rec["Email 2"]        = output.get("email_2", "")
            rec["Email 3"]        = output.get("email_3", "")
            rec["Status"]         = ""

            log(f"  ✅ Done\n")
            set_progress(i + 1)
            time.sleep(API_DELAY)

        # Save output
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_filename = f"campaign_{ts}_{job_id[:8]}.xlsx"
        out_path = os.path.join(OUTPUT_DIR, out_filename)
        write_excel(records, out_path)

        with jobs_lock:
            jobs[job_id]["status"]      = "done"
            jobs[job_id]["output_file"] = out_filename
        log(f"✅ All done! File ready to download.")

    except Exception as e:
        with jobs_lock:
            jobs[job_id]["status"] = "error"
        log(f"❌ Fatal error: {str(e)}")
    finally:
        try:
            os.remove(input_path)
        except Exception:
            pass


# ─── ROUTES ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    """Step 1: save file, return headers for column mapping."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    f = request.files["file"]
    if not f.filename.endswith((".xlsx", ".xls")):
        return jsonify({"error": "Please upload an Excel file (.xlsx or .xls)"}), 400

    file_id = str(uuid.uuid4())
    filename = f"input_{file_id}.xlsx"
    input_path = os.path.join(UPLOAD_DIR, filename)
    f.save(input_path)

    try:
        info = get_excel_headers(input_path)
    except Exception as e:
        os.remove(input_path)
        return jsonify({"error": str(e)}), 400

    return jsonify({"file_id": file_id, **info})


@app.route("/start", methods=["POST"])
def start():
    """Step 2: receive confirmed column mapping, kick off processing."""
    data = request.get_json()
    file_id  = data.get("file_id")
    col_map  = data.get("col_map")   # {company, contact, website, phone} -> int

    if not file_id or not col_map:
        return jsonify({"error": "Missing file_id or col_map"}), 400

    input_path = os.path.join(UPLOAD_DIR, f"input_{file_id}.xlsx")
    if not os.path.isfile(input_path):
        return jsonify({"error": "File not found — please re-upload"}), 404

    # Convert col_map values to int
    col_map = {k: int(v) for k, v in col_map.items() if v is not None and v != ""}

    job_id = str(uuid.uuid4())
    with jobs_lock:
        jobs[job_id] = {
            "status": "running",
            "progress": 0,
            "total": 0,
            "log": [],
            "output_file": None,
        }

    thread = threading.Thread(target=process_job, args=(job_id, input_path, col_map), daemon=True)
    thread.start()
    return jsonify({"job_id": job_id})


@app.route("/status/<job_id>")
def status(job_id: str):
    with jobs_lock:
        job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    return jsonify({
        "status":      job["status"],
        "progress":    job["progress"],
        "total":       job["total"],
        "log":         job["log"],
        "output_file": job["output_file"],
    })


@app.route("/download/<filename>")
def download(filename: str):
    # Basic path safety
    safe = os.path.basename(filename)
    return send_from_directory(OUTPUT_DIR, safe, as_attachment=True)


# ─── ENTRY POINT ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    host = "0.0.0.0" if os.environ.get("PORT") else "127.0.0.1"
    print("\n" + "="*55)
    print("  Landscaper Email Campaign Generator")
    print(f"  Running at http://{host}:{port}")
    print("="*55 + "\n")
    app.run(debug=False, host=host, port=port)

"""
Landscaper Email Campaign Generator — Stateless Flask App
Works on Vercel (serverless) and locally.
Frontend drives the loop: one /process-record call per landscaper.
"""

import io
import os
import re
import json
from datetime import datetime

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from flask import Flask, render_template, request, jsonify, send_file
from google import genai

# ─── CONFIG ────────────────────────────────────────────────────────────────────

# Load .env if present (local dev)
_env = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
if os.path.isfile(_env):
    with open(_env) as _f:
        for _line in _f:
            _line = _line.strip()
            if _line and not _line.startswith("#") and "=" in _line:
                _k, _v = _line.split("=", 1)
                os.environ.setdefault(_k.strip(), _v.strip())

GEMINI_MODEL    = "gemini-2.5-flash"
REQUEST_TIMEOUT = 10

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
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024

# Lazy Gemini client — initialised on first use so a missing env var
# doesn't crash the entire serverless function at import time.
_gemini_client = None

def get_gemini_client():
    global _gemini_client
    if _gemini_client is None:
        api_key = os.environ.get("GEMINI_API_KEY", "")
        if not api_key:
            raise RuntimeError("GEMINI_API_KEY environment variable is not set.")
        _gemini_client = genai.Client(api_key=api_key)
    return _gemini_client


# ─── HELPERS ───────────────────────────────────────────────────────────────────

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
        soup = BeautifulSoup(resp.text, "html.parser")
        for tag in soup(["script", "style", "nav", "footer", "head",
                         "noscript", "iframe", "svg"]):
            tag.decompose()
        title_tag = soup.find("title")
        result["title"] = clean_text(title_tag.get_text()) if title_tag else ""
        meta = soup.find("meta", attrs={"name": re.compile(r"description", re.I)})
        if meta and meta.get("content"):
            result["page_meta"] = clean_text(meta["content"])
        body = soup.find("body")
        raw = body.get_text(separator=" ") if body else soup.get_text(separator=" ")
        result["body_text"] = clean_text(raw)[:3000]
        result["has_contact_form"] = bool(soup.find_all("form"))
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


def build_prompt(company: str, contact: str, site_data: dict) -> str:
    if site_data["success"]:
        site_summary = f"""WEBSITE DATA:
- Title: {site_data['title']}
- Meta description: {site_data['page_meta']}
- Has contact/lead form: {site_data['has_contact_form']}
- Has booking / quote CTA: {site_data['has_booking']}
- Has call-to-action: {site_data['has_cta']}
- Page text excerpt:
{site_data['body_text'][:2500]}"""
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
        response = get_gemini_client().models.generate_content(
            model=GEMINI_MODEL,
            contents=prompt,
        )
        raw = response.text.strip()
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw.strip())
        return json.loads(raw)
    except json.JSONDecodeError:
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


def build_excel_bytes(records: list) -> bytes:
    """Generate the campaign Excel in memory and return raw bytes."""
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
            rec.get("Company Name", ""),   rec.get("Contact Name", ""),
            rec.get("Website URL", ""),    rec.get("Phone Number", ""),
            rec.get("research_notes", ""), rec.get("email_1", ""),
            rec.get("email_2", ""),        rec.get("email_3", ""),
            "",
        ], 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.fill = fill
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    for ci, w in enumerate([28, 20, 35, 18, 50, 70, 70, 70, 15], 1):
        ws.column_dimensions[ws.cell(row=1, column=ci).column_letter].width = w
    ws.row_dimensions[1].height = 30
    for ri in range(2, len(records) + 2):
        ws.row_dimensions[ri].height = 120
    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ─── ROUTES ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    """Read Excel in memory, return headers + all rows as JSON for frontend mapping."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    f = request.files["file"]
    if not f.filename.endswith((".xlsx", ".xls")):
        return jsonify({"error": "Please upload an Excel file (.xlsx or .xls)"}), 400

    try:
        raw = f.read()
        wb = openpyxl.load_workbook(io.BytesIO(raw))
        ws = wb.active
        all_rows = list(ws.iter_rows(values_only=True))
    except Exception as e:
        return jsonify({"error": f"Could not read file: {str(e)}"}), 400

    if not all_rows:
        return jsonify({"error": "File is empty"}), 400

    headers = [str(c).strip() if c is not None else "" for c in all_rows[0]]
    sample  = [str(c).strip() if c is not None else "" for c in (all_rows[1] if len(all_rows) > 1 else [])]

    # Auto-guess column indices
    guesses = {}
    for i, h in enumerate(headers):
        hl = h.lower()
        if any(k in hl for k in ["company", "business"]) and "contact" not in hl:
            guesses.setdefault("company", i)
        elif any(k in hl for k in ["contact", "first name", "person"]):
            guesses.setdefault("contact", i)
        elif any(k in hl for k in ["website", "url", "web", "site"]):
            guesses.setdefault("website", i)
        elif any(k in hl for k in ["phone", "tel", "mobile", "cell"]):
            guesses.setdefault("phone", i)

    # Return all data rows as plain arrays (frontend applies col mapping)
    data_rows = []
    for row in all_rows[1:]:
        if all(c is None or str(c).strip() == "" for c in row):
            continue
        data_rows.append([str(c).strip() if c is not None else "" for c in row])

    return jsonify({
        "headers":  headers,
        "sample":   sample,
        "guesses":  guesses,
        "rows":     data_rows,
        "total":    len(data_rows),
    })


@app.route("/process-record", methods=["POST"])
def process_record():
    """Process a single landscaper: fetch website + generate emails.
    Called once per record from the frontend loop — stays well under Vercel timeout."""
    data = request.get_json()
    company = data.get("company", "") or "Unknown Company"
    contact = data.get("contact", "") or "there"
    website = data.get("website", "")
    phone   = data.get("phone", "")

    site_data = fetch_website(website)
    result    = generate_emails(company, contact, site_data)

    return jsonify({
        "Company Name":   company,
        "Contact Name":   contact,
        "Website URL":    website,
        "Phone Number":   phone,
        "site_ok":        site_data["success"],
        "site_error":     site_data.get("error", ""),
        "research_notes": result.get("research_notes", ""),
        "email_1":        result.get("email_1", ""),
        "email_2":        result.get("email_2", ""),
        "email_3":        result.get("email_3", ""),
    })


@app.route("/export", methods=["POST"])
def export_excel():
    """Receive all completed records, return Excel file as download."""
    records = request.get_json()
    if not records:
        return jsonify({"error": "No records to export"}), 400

    try:
        xlsx_bytes = build_excel_bytes(records)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return send_file(
        io.BytesIO(xlsx_bytes),
        as_attachment=True,
        download_name=f"campaign_{ts}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ─── ENTRY POINT ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    host = "0.0.0.0" if os.environ.get("PORT") else "127.0.0.1"
    print("\n" + "=" * 55)
    print("  Landscaper Email Campaign Generator")
    print(f"  Running at http://{host}:{port}")
    print("=" * 55 + "\n")
    app.run(debug=False, host=host, port=port)

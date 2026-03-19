# Landscaper Email Campaign Generator

A local web app that turns a list of landscaping leads into a fully personalised 3-email cold outreach sequence — powered by **Gemini 2.5 Flash**.

## What it does

1. **Reads** your Excel file (Company Name, Contact Name, Website URL, Phone Number)
2. **Scrapes** each landscaper's website live — extracts services, lead forms, booking flows, CTAs
3. **Generates** with Gemini:
   - Research notes on their lead capture gaps
   - **Email 1** — personalised initial outreach (100–150 words)
   - **Email 2** — day-3 follow-up with a market insight
   - **Email 3** — day-8 follow-up with a case study + soft close
4. **Downloads** a formatted Excel with everything ready to send

## Setup

### 1. Clone & install

```bash
git clone https://github.com/Pirintvisuals/email-personalizer.git
cd email-personalizer
pip install flask google-genai requests beautifulsoup4 openpyxl lxml
```

### 2. Add your Gemini API key

Create a `.env` file in the project root:

```
GEMINI_API_KEY=your_key_here
```

Get a free key at [aistudio.google.com](https://aistudio.google.com)

### 3. Run

```bash
py app.py
```

Open **http://127.0.0.1:5000** in your browser.

## Input format

Your Excel file needs these columns (order and exact name don't matter):

| Company Name | Contact Name | Website URL | Phone Number |
|---|---|---|---|
| Green Thumb Co | Mike Johnson | greenthumb.com | 555-0101 |

A `sample_landscapers.xlsx` is included for testing.

## Tech stack

- **Backend:** Python / Flask
- **AI:** Google Gemini 2.5 Flash
- **Scraping:** requests + BeautifulSoup
- **Excel:** openpyxl

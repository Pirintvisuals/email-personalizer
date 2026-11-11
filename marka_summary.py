"""
Márkaszerviz campaign summary: sent, reply rate, positive reply rate,
overall, by make and by city.

    py marka_summary.py                      # every leads-markaszerviz-*.xlsx in Asztali gép/Leads/Hungary
    py marka_summary.py path\\to\\file.xlsx   # one or more specific files

Reads only what the sender writes into the sheet: "Elküldve?" (Sent ...),
"Követés", "Válaszolt?", "Kézbesítés", "Válasz típusa".
"""

import glob
import re
import sys
from collections import defaultdict
from pathlib import Path

import openpyxl

import hu_email_marka as marka

sys.stdout.reconfigure(encoding="utf-8")

_SENT_RE = re.compile(r"^sent\s+\d{4}-\d{2}-\d{2}", re.IGNORECASE)


def _col(headers, *needles):
    return next((i for i, h in enumerate(headers) if any(n in h for n in needles)), None)


def rows_from(path):
    ws = openpyxl.load_workbook(path, read_only=True).active
    it = ws.iter_rows(values_only=True)
    headers = [str(h or "").strip().lower() for h in next(it)]
    c = {
        "city": _col(headers, "city", "város"),
        "make": _col(headers, "make(s)", "make"),
        "sent": _col(headers, "elküldve", "küldve"),
        "fup": _col(headers, "követés"),
        "replied": _col(headers, "válaszolt"),
        "bounced": _col(headers, "kézbesít"),
        "type": _col(headers, "válasz típus"),
    }
    if c["sent"] is None:
        return
    get = lambda r, k: str(r[c[k]] or "").strip() if c[k] is not None and c[k] < len(r) else ""
    for r in it:
        if not _SENT_RE.match(get(r, "sent")):
            continue
        yield {
            "city": get(r, "city") or "?",
            "make": marka.make_label(get(r, "make")),
            "followed_up": bool(get(r, "fup")),
            "replied": bool(get(r, "replied")),
            "bounced": bool(get(r, "bounced")),
            "type": get(r, "type"),
        }


def line(label, rs, width=22):
    sent = len(rs)
    rep = sum(r["replied"] for r in rs)
    pos = sum(r["type"] == marka.REPLY_TYPES["interested"][0] for r in rs)
    pct = lambda n: f"{n / sent * 100:4.0f}%" if sent else "   -"
    return f"{label[:width]:<{width}} {sent:>5} {rep:>5} {pct(rep)} {pos:>5} {pct(pos)}"


def table(title, groups):
    print(f"\n{title}")
    print(f"{'':<22} {'küld.':>5} {'vál.':>5} {'arány':>5} {'pozit':>5} {'arány':>5}")
    for k in sorted(groups, key=lambda k: (-len(groups[k]), k)):
        print(line(k, groups[k]))


def main():
    files = sys.argv[1:] or sorted(glob.glob(str(Path.home() / "OneDrive" / "Asztali gép" / "Leads" / "Hungary" / "leads-markaszerviz-*.xlsx")))
    files = [f for f in files if "backup" not in f.lower()]
    if not files:
        print("Nem találok leads-markaszerviz-*.xlsx fájlt.")
        return
    rs = [r for f in files for r in rows_from(f)]
    print("Fájlok:", ", ".join(Path(f).name for f in files))
    if not rs:
        print("Még nincs elküldött levél.")
        return

    types = defaultdict(int)
    for r in rs:
        if r["replied"]:
            types[r["type"] or "(nincs besorolva)"] += 1

    print()
    print(f"{'':<22} {'küld.':>5} {'vál.':>5} {'arány':>5} {'pozit':>5} {'arány':>5}")
    print(line("ÖSSZESEN", rs))
    print(f"\nEmlékeztető elküldve: {sum(r['followed_up'] for r in rs)}   "
          f"Visszapattant: {sum(r['bounced'] for r in rs)}   "
          f"Tiltólistán: {len(marka.dnc())}")
    print("Válaszok típus szerint: " + (", ".join(f"{k}: {v}" for k, v in types.items()) or "-"))

    by_make, by_city = defaultdict(list), defaultdict(list)
    for r in rs:
        by_make[r["make"]].append(r)
        by_city[r["city"]].append(r)
    table("Márka szerint", by_make)
    table("Város szerint", by_city)


if __name__ == "__main__":
    main()

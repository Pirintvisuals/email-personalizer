"""
Autószerelő (Hungarian mechanics) section for the HU DM app.

A second, fully separate sender that lives next to the Auto Sender: its own
Excel sheet, its own progress state, its own routes (/mech_*), its own message.
Nothing here touches the Auto Sender's state, so a different sheet can be loaded
in each at the same time and either one can be changed without affecting the
other.

Wired in by dm_app_hu.py (`hu_mech.register(dm_app)`), so it only exists on the
HU app (port 5002). The wording lives in MECH_PARAGRAPHS at the top of this
file - that is the only thing to edit when the copy changes.
"""

import os
import random
import re
from pathlib import Path

from flask import request, jsonify

# ── THE MESSAGE ───────────────────────────────────────────────────────────
# Edit MECH_PARAGRAPHS to change the wording. Everything below the state dict
# is plumbing (sheet, Chrome, clipboard, progress tracking).
#
# Style rules this copy follows. All of them are load-bearing:
#   * No em dash anywhere. tidy_message() rewrites them to ", " anyway, and
#     they read as an AI tell (standing rule across the HU copy).
#   * Magázás, singular "Ön", and the pronoun only where it carries weight.
#     Hungarian conjugation already marks the formal register, so repeating it
#     in every clause is what made the first drafts read translated.
#   * No link, and no finished product claimed. Meta filters links in a first
#     cold DM, so the close asks permission and the link follows in the reply.
#     What exists for them is a DEMÓ, which is also what makes the close work:
#     "nem kell elhinnie" only lands while the ask stays small.
#   * NO quoting-time figure, in either direction. A mechanic knows how long
#     his own quoting takes, so putting a number on it is how you earn "ez nem
#     rólam szól" in the first three lines. The gain is named instead: the
#     quotes written for people who never become customers, which nobody pays
#     for. He supplies the numbers himself, and his are more persuasive.
#   * "Milyen autó és mi a hiba" is stated as the START of the questions, never
#     the whole of them ("onnan pedig addig kérdez tovább..."). Without that a
#     mechanic reads two questions, concludes it cannot price anything, and
#     that is the objection most likely to kill this message.
#   * The accuracy objection is pre-handled twice over: the questions are his
#     ("A kérdéseket Ön adja meg") and nothing leaves without him reading it.
#     Never by out-experting a mechanic on his own trade, which is an argument
#     we lose and deserve to.
#   * Mechanics-only is said twice, scoped to what is BUILT ("autószerelőknek
#     csinálok", "most csak autószerelőknek csinálok ilyet"), not to who the
#     clients are. The UK trades work means the wider claim would not be true.

# 2026-10-07 rewrite (user's copy). The old version had the CUSTOMER typing
# in the car and the fault. The new one answers the two objections a
# mechanic actually raises: the customer cannot diagnose his own car (so the
# shop's intake person fills it in after the inspection), and a generic tool
# cannot know his prices (so it uses his brands, prices and labour rate).
MECH_PARAGRAPHS = [
    "Autószerelőknek csinálok árajánlat-készítőt, Önnek pedig összeraktam "
    "egy demót.",

    "Nem az ügyfél tölti ki, és nem ő mondja a hibát. Átvizsgálás után a "
    "munkafelvevő megadja az autót és a munkát, a program pedig összerakja a "
    "tételes ajánlatot az Önök márkáival, áraival és munkadíjával. Mindig Ön "
    "nézi át, mielőtt kimegy.",

    "Árajánlatot annak is írni kell, akiből végül nem lesz munka. Ez az az "
    "idő, amit senki nem fizet ki. Ezt a részt veszi le Önről.",

    "Nem kell elhinnie, hogy működik. Megnézi, és eldönti. Elküldhetem?",
]

# Follow-up nudges, same cadence as Auto Sender (3 / 7 / 11 days, max 3).
#
# Written fresh rather than reusing dm_app_hu._hu_followup_text, which is the
# WRONG script for these leads twice over: it names "a chatbot" and "az
# árkalkulátor" (the mechanic message deliberately says árajánlatkészítő and
# demó), and its touch 2 pitches lead qualification ("azok jutnak el Önig,
# akik tényleg meg akarják csináltatni") — the promise we dropped for
# mechanics in favour of the time-saver. A nudge that contradicts the first
# message reads as a different person selling a different product.
#
# Same rules as MECH_PARAGRAPHS: magázás, no em dash, no link, and the ask
# stays small ("csak szóljon").
MECH_FOLLOWUPS = {
    1: [
        "Jó napot! Csak visszajelzek, nem sürgős. A demó, amit Önnek "
        "összeraktam, továbbra is áll, csak szóljon, és küldöm.",
        "Jó napot! Nem akarok tolakodni, csak egy gyors emlékeztető. Az "
        "árajánlatkészítő demója megvan Önnek, egy szavába kerül, és átküldöm.",
        "Jó napot! Csak egy rövid emlékeztető, nem sürgős. A demó kész van, "
        "bármikor meg tudom mutatni, csak jelezzen.",
    ],
    2: [
        "Jó napot, még egyszer visszatérek erre. A lényeg röviden: árajánlatot "
        "annak is írni kell, akiből végül nem lesz munka, és ezt a részt veszi "
        "le Önről. Megmutathatom a demót?",
        "Jó napot! Csak még egyszer megkérdezem. Nem kell elhinnie, hogy "
        "működik, megnézi és eldönti. Küldjem a demót?",
        "Jó napot, még egyszer jelentkezem. Az árajánlatírás az az idő, amit "
        "soha nem fizet ki senki. Pont ezt veszi le Önről. Átküldhetem a demót?",
    ],
    3: [
        "Jó napot! Utoljára jelentkezem ezzel, nem szeretnék tolakodni. Ha "
        "később mégis aktuális lenne, szóljon nyugodtan, a demó megmarad Önnek.",
        "Jó napot! Ez az utolsó üzenetem ebben. Ha most nem aktuális, az "
        "teljesen rendben, csak jelezzen, ha egyszer mégis megnézné.",
    ],
}


def mech_followup_text(touch: int, name: str = "") -> str:
    """Nudge text for touch 1, 2 or 3. `name` is accepted for signature
    parity with dm_app._followup_text and deliberately unused: the greeting is
    already in the line, and pasting a company name into a nudge ("Kedves
    Generálkivitelező Kft. pécs") is the exact mangling _hu_pick_body avoids."""
    t = max(1, min(int(touch or 1), dm_app_touches()))
    return random.choice(MECH_FOLLOWUPS[t])


def dm_app_touches() -> int:
    """FOLLOWUP_TOUCHES without importing dm_app at module level."""
    import sys
    dm = sys.modules.get("dm_app")
    return int(getattr(dm, "FOLLOWUP_TOUCHES", 3)) if dm else 3


# Opener hooks, distinctive service first. A shop that does DPF cleaning is
# worth naming for it; "olajcsere" is true of every shop in the country and
# says nothing, so it sits at the bottom as a last resort before the city
# fallback.
#
# Deliberately NOT in hu_niche.py: that file is shared with Auto Sender and the
# email app, and it already carries "autóüveg" and "autókozmetika" niches that
# a general szerviz page would collide with. Keeping the mechanic vocabulary
# here means this section cannot regress a working niche.
_MECH_OPENER_HOOKS = [
    ("részecskeszűrő",      "részecskeszűrő-tisztítást is vállalnak"),
    ("dpf",                 "DPF-tisztítást is vállalnak"),
    ("automata váltó",      "automataváltó-javítással is foglalkoznak"),
    ("automataváltó",       "automataváltó-javítással is foglalkoznak"),
    ("turbó",               "turbójavítással is foglalkoznak"),
    ("hibrid",              "hibrid autókkal is foglalkoznak"),
    ("eredetiségvizsgálat", "eredetiségvizsgálatot is végeznek"),
    ("futómű",              "futómű-beállítást is vállalnak"),
    ("vezérműszíj",         "vezérműszíj-cserét is vállalnak"),
    ("autóvillamosság",     "autóvillamossági munkákat is vállalnak"),
    ("klímatöltés",         "klímatöltést is vállalnak"),
    ("klíma",               "klímaszervizt is vállalnak"),
    ("karosszéria",         "karosszériamunkát is vállalnak"),
    ("fényezés",            "fényezést is vállalnak"),
    ("centírozás",          "centírozást is vállalnak"),
    ("gumiszerel",          "gumiszerelést is vállalnak"),
    ("eredetiség",          "eredetiségvizsgálatot is végeznek"),
    ("műszaki vizsga",      "műszaki vizsgáztatással is foglalkoznak"),
    ("vizsgáztatás",        "műszaki vizsgáztatással is foglalkoznak"),
    ("diagnosztika",        "számítógépes diagnosztikát is végeznek"),
    ("kuplung",             "kuplungcserét is vállalnak"),
    ("kipufogó",            "kipufogójavítást is vállalnak"),
    ("fékbetét",            "fékjavítást is vállalnak"),
    ("olajcsere",           "olajcserét is vállalnak"),
]


def _dm_app_hu():
    """The dm_app_hu module object, WITHOUT ever re-importing it.

    START_HU.bat runs `python dm_app_hu.py`, so that file is loaded under the
    name "__main__", not "dm_app_hu". A plain `import dm_app_hu` therefore does
    not find it in sys.modules and executes the entire file a SECOND time as a
    separate module object. That re-runs every top-level side effect in it:
    hu_mech.register() fires again against the already-serving Flask app and
    raises "The setup method 'route' can no longer be called on the
    application", and dm_app._run_groq gets wrapped around its own wrapper.

    So: prefer the live module under either name, and only fall back to a real
    import when it genuinely has not been loaded (a standalone test).
    """
    import sys
    mod = sys.modules.get("dm_app_hu")
    if mod is not None:
        return mod
    main = sys.modules.get("__main__")
    if str(getattr(main, "__file__", "") or "").replace("\\", "/").endswith("dm_app_hu.py"):
        return main
    import dm_app_hu as mod
    return mod


def _mech_opener(lead: dict, text_data: str) -> str:
    """One personalised line from a service actually named on their page.

    Text-sourced rather than vision-sourced on purpose: a car on a lift looks
    like every other car on a lift, so a photo yields far less here than a
    finished patio does for a landscaper. This costs no model call and no
    screenshot, which is also the entire per-lead wait.
    """
    # Resolved late, never imported at module level: dm_app_hu imports THIS
    # module, so a top-level import would be circular. See _dm_app_hu() for
    # why it is not a plain `import` either.
    dm_app_hu = _dm_app_hu()
    from hu_niche import _norm

    t = _norm(text_data)
    city = dm_app_hu._td_section(text_data, "LISTED CITY") or (lead.get("city") or "")
    where = f" {dm_app_hu._city_to_ragos(city)}" if city else ""

    for needle, phrase in _MECH_OPENER_HOOKS:
        if _norm(needle) in t:
            print(f"  Opener HOOK '{needle}' -> {phrase}")
            return f"Láttam, hogy {phrase}{where}."

    # Same evidence-gathering habit as dm_app_hu's non-photo niches: when
    # nothing matches, print what the page actually said, so the hook list can
    # be grown from real pages instead of guesses. After a batch, grep
    # server_hu.log for "NO HOOK" and the real vocabulary is sitting there.
    page = text_data
    for hdr in ("--- COMPANY NAME ---", "--- TRADE TYPE ---", "--- LISTED CITY ---"):
        page = page.replace(hdr, " ")
    page = " ".join(page.split())[:240]
    print(f"  NO HOOK [mechanic] page said: {page!r}")
    return dm_app_hu._hu_city_fallback(text_data)


def build_mechanic_dm(greeting: str, lead: dict, text_data: str) -> str:
    """greeting is already built ("Kedves Péter! " / "Jó napot kívánok! ")."""
    opener = _mech_opener(lead, text_data)
    if opener:
        return "\n\n".join([(greeting + opener).strip()] + MECH_PARAGRAPHS)
    # No opener at all: fold the greeting into the first real paragraph rather
    # than leaving it stranded on a line of its own.
    return "\n\n".join([(greeting + MECH_PARAGRAPHS[0]).strip()] + MECH_PARAGRAPHS[1:])


# ── State (separate from Auto Sender / FB Group / Owner Emails) ───────────────
_mech_state = {
    "excel_path":  None,
    "current_row": None,
    "sent_today":  0,
    "wb":          None,
    "ws":          None,
    "cols":        {},
}

_DONE = ("yes", "y", "sent", "1", "true", "skip")
_SENT = ("yes", "y", "sent", "1", "true")

# Prefilled in the UI, and used when the path box is left blank. Overridable
# with the MECH_EXCEL env var so a different list needs no code edit.
# Deliberately NOT dm_app.DEFAULT_EXCEL: that one belongs to Auto Sender, and
# the point of this section is that the two sheets stay independent.
# Forward slashes on purpose, they work fine on Windows and keep the string
# free of backslash escapes.
MECH_DEFAULT_EXCEL = os.environ.get(
    "MECH_EXCEL",
    # 2026-10-06: new Facebook mechanics batch. The workbook opens on its
    # "New FB mechanics" sheet (the loader reads the active sheet). The old
    # 200-lead list still holds the follow-ups for its own sends.
    "C:/Users/pirin/Downloads/hun_facebook_dm_leads 10. 06.xlsx")

_NAME_HINTS = ("szerel", "mechanic", "auto", "autó")


def register(dm_app):
    """Attach the /mech_* routes to the shared Flask app and turn the tab on.

    Idempotent on purpose. Flask refuses @app.route() once the app has served
    a request, so a second call would crash the running process rather than be
    ignored. Belt and braces alongside _dm_app_hu(): any future path that ends
    up importing dm_app_hu twice degrades to a no-op instead of a 500.
    """
    if any(r.rule == "/mech_load" for r in dm_app.app.url_map.iter_rules()):
        print("  hu_mech: routes already registered, skipping re-registration")
        dm_app.MECHANICS_SECTION = True
        dm_app.MECH_DEFAULT_EXCEL = MECH_DEFAULT_EXCEL
        return
    dm_app.MECHANICS_SECTION = True
    dm_app.MECH_DEFAULT_EXCEL = MECH_DEFAULT_EXCEL
    app = dm_app.app

    def _save():
        import time as _t
        wb, path = _mech_state.get("wb"), _mech_state.get("excel_path")
        if wb is None or not path:
            return
        for _ in range(6):
            try:
                wb.save(path)
                return
            except PermissionError:
                _t.sleep(0.4)
        raise PermissionError(
            "Can't write to the Excel file - it's open in Excel (or still syncing in "
            "OneDrive). Close the spreadsheet in Excel, wait for OneDrive to finish, "
            "then click again.")

    def _cell(row, cols, key):
        i = cols.get(key)
        return str(row[i].value or "").strip() if i is not None else ""

    def _next_lead(ws, cols, skip_row=None):
        for row in ws.iter_rows(min_row=2):
            excel_row = row[0].row
            if skip_row is not None and excel_row == skip_row:
                continue
            url = _cell(row, cols, "url")
            if not url.startswith("http"):
                continue
            low = url.lower()
            if "facebook.com" not in low and "instagram.com" not in low:
                continue   # website-only row: no page to message
            if _cell(row, cols, "contacted").lower() in _DONE:
                continue
            return {
                "row":   excel_row,
                "url":   url,
                "name":  _cell(row, cols, "name"),
                "city":  _cell(row, cols, "city"),
                "niche": _cell(row, cols, "niche"),
            }
        return None

    def _counts(ws, cols):
        total = sum(1 for r in ws.iter_rows(min_row=2)
                    if str(r[cols["url"]].value or "").strip())
        sent = sum(1 for r in ws.iter_rows(min_row=2)
                   if cols["contacted"] is not None
                   and str(r[cols["contacted"]].value or "").strip().lower() in _SENT)
        return total, sent

    def _find_default_sheet():
        if MECH_DEFAULT_EXCEL and Path(MECH_DEFAULT_EXCEL).exists():
            return MECH_DEFAULT_EXCEL
        for f in Path(__file__).parent.glob("*.xlsx"):
            if any(k in f.name.lower() for k in _NAME_HINTS):
                return str(f)
        return ""

    @app.route("/mech_load", methods=["POST"])
    def mech_load():
        try:
            path = (request.get_json() or {}).get("path", "").strip().strip('"')
            if not path:
                path = _find_default_sheet()
            if not path or not Path(path).exists():
                return jsonify({"error": f"Excel file not found: {path or '(no path given)'}"}), 400

            wb, ws, cols = dm_app._load_excel(path)
            _mech_state.update({"excel_path": path, "wb": wb, "ws": ws, "cols": cols,
                                "current_row": None})

            lead = _next_lead(ws, cols)
            total, sent = _counts(ws, cols)
            if not lead:
                # Not an error: a fully sent sheet is exactly when the
                # follow-ups matter, and the UI only shows them after a
                # successful load.
                return jsonify({"done": True, "lead": None, "total": total, "sent": sent})
            return jsonify({"lead": lead, "total": total, "sent": sent})
        except Exception as e:
            import traceback; traceback.print_exc()
            return jsonify({"error": str(e)}), 500

    @app.route("/mech_next", methods=["POST"])
    def mech_next():
        """Open the next lead in Chrome, build the mechanic message, copy it."""
        try:
            ws, cols = _mech_state["ws"], _mech_state["cols"]
            if ws is None:
                return jsonify({"error": "Load your Excel file first."}), 400

            # A lead that was opened but never marked sent is skipped, so it
            # can never resurface as a duplicate.
            prev_row = _mech_state["current_row"]
            if prev_row is not None and cols["contacted"] is not None:
                ws.cell(row=prev_row, column=cols["contacted"] + 1, value="Skip")
                _save()

            lead = _next_lead(ws, cols, skip_row=prev_row)
            if not lead:
                return jsonify({"error": "No more leads - all done!"}), 200
            _mech_state["current_row"] = lead["row"]

            import socket
            import time as _t
            from playwright.sync_api import sync_playwright

            connected = False
            for _ in range(5):
                try:
                    socket.create_connection(("127.0.0.1", dm_app.CDP_PORT), timeout=2).close()
                    connected = True
                    break
                except OSError:
                    _t.sleep(1)
            if not connected:
                return jsonify({"error": "Chrome not found. Close all Chrome windows, "
                                         "double-click START_HU.bat, wait for Chrome to fully "
                                         "open, then try again."}), 500

            with sync_playwright() as p:
                browser = p.chromium.connect_over_cdp(f"http://localhost:{dm_app.CDP_PORT}")
                ctx = browser.contexts[0]

                # Same tidy-up as the other senders: old FB/Messenger tabs pile
                # up fast and each one is a heavy app.
                for pg in list(ctx.pages):
                    u = pg.url.lower()
                    if "localhost" in u:
                        continue
                    if "facebook.com" in u or "messenger.com" in u or "m.me/" in u:
                        try:
                            pg.close()
                        except Exception:
                            pass

                page = ctx.new_page()
                page.bring_to_front()
                try:
                    page.goto(lead["url"], wait_until="domcontentloaded", timeout=45000)
                except Exception:
                    try:
                        page.goto(lead["url"], wait_until="commit", timeout=30000)
                    except Exception as e:
                        return jsonify({"error": f"Could not load page: {e}"}), 500
                page.wait_for_timeout(500)
                try:
                    page.keyboard.press("Escape")
                except Exception:
                    pass
                dm_app._strip_overlays(page)

                # Page text: only used to spot the owner's first name.
                try:
                    text = page.evaluate("() => document.body ? document.body.innerText : ''")
                    text = re.sub(r"\n{3,}", "\n\n", text).strip()[:12000]
                except Exception:
                    text = ""

                text_data = text
                if lead["niche"]:
                    text_data = f"--- TRADE TYPE ---\n{lead['niche']}\n\n" + text_data
                if lead["name"]:
                    text_data = f"--- COMPANY NAME ---\n{lead['name']}\n\n" + text_data
                if lead["city"]:
                    text_data += f"\n\n--- LISTED CITY ---\n{lead['city']}"

                name_hint = dm_app.extract_first_name(text or "", lead.get("name", ""))
                greeting = (dm_app.GREETING_NAME_FMT.format(name=name_hint)
                            if name_hint else dm_app.GREETING_SOLO_FMT)

                dm = dm_app.tidy_message(build_mechanic_dm(greeting, lead, text_data).strip())

                msg_opened = dm_app._click_message_btn(page)
                if not msg_opened:
                    print("  Could not click Message button - user can open chat manually.")
                clipboard = dm_app.copy_to_clipboard(dm)

            return jsonify({
                "dm":         dm,
                "lead":       lead,
                "clipboard":  clipboard,
                "msg_opened": msg_opened,
                "sent_today": _mech_state["sent_today"],
            })
        except Exception as e:
            import traceback; traceback.print_exc()
            return jsonify({"error": str(e)}), 500

    @app.route("/mech_skip", methods=["POST"])
    def mech_skip():
        try:
            ws, cols = _mech_state["ws"], _mech_state["cols"]
            row = _mech_state["current_row"]
            if ws is None:
                return jsonify({"error": "No Excel loaded."}), 400
            if row is not None and cols["contacted"] is not None:
                ws.cell(row=row, column=cols["contacted"] + 1, value="Skip")
                _save()
            _mech_state["current_row"] = None

            lead = _next_lead(ws, cols)
            if not lead:
                return jsonify({"error": "No more leads - all done!"}), 200
            _mech_state["current_row"] = lead["row"]
            total, sent = _counts(ws, cols)
            return jsonify({"ok": True, "lead": lead, "sent": sent, "total": total})
        except Exception as e:
            return jsonify({"error": str(e)}), 500

    _REPLIED_YES = ("yes", "y", "1", "true", "igen")

    @app.route("/mech_followups_due", methods=["POST"])
    def mech_followups_due():
        """Mechanic leads due for a nudge.

        Reads _mech_state, NOT _auto_state. The Follow-Up and Replies tabs in
        dm_app.py are both hardwired to the Auto Sender workbook, so before
        this existed a mechanic lead could never appear in either: with no
        sheet loaded there you got "Load your leads on the Auto Sender tab
        first", and with a DIFFERENT sheet loaded you got that sheet's leads
        and wrote the nudge counter into that workbook.
        """
        try:
            ws, cols = _mech_state["ws"], _mech_state["cols"]
            if ws is None:
                return jsonify({"error": "Load your mechanics Excel on this tab first."}), 400
            from datetime import date
            today = date.today()
            gap = getattr(dm_app, "_FOLLOWUP_GAP", {0: 3, 1: 7, 2: 11})
            due = []
            for r in ws.iter_rows(min_row=2):
                row = r[0].row
                url = _cell(r, cols, "url")
                if not url.startswith("http"):
                    continue
                if _cell(r, cols, "contacted").lower() not in _SENT:
                    continue          # never messaged (or skipped)
                if _cell(r, cols, "replied").lower() in _REPLIED_YES:
                    continue          # they answered, never nudge
                fups = dm_app._to_int(r[cols["followups"]].value) if cols.get("followups") is not None else 0
                if fups >= dm_app.FOLLOWUP_TOUCHES:
                    continue          # sequence finished
                entry = {"row": row, "url": url, "touch": fups + 1,
                         "name": _cell(r, cols, "name"),
                         "city": _cell(r, cols, "city"),
                         "niche": _cell(r, cols, "niche")}
                last = dm_app._parse_date(r[cols["sent_date"]].value) if cols.get("sent_date") is not None else None
                if last is None:
                    # Contacted but undated: cannot tell HOW overdue, but it is
                    # still a no-reply lead that would otherwise vanish.
                    entry.update({"days": None, "no_date": True})
                    due.append(entry)
                    continue
                days = (today - last).days
                if days < gap.get(fups, 99):
                    continue          # not due yet
                entry["days"] = days
                due.append(entry)
            due.sort(key=lambda d: (0 if d.get("no_date") else 1, -(d["days"] or 0)))
            return jsonify({"ok": True, "due": due, "count": len(due)})
        except Exception as e:
            import traceback; traceback.print_exc()
            return jsonify({"error": str(e)}), 500

    @app.route("/mech_followup_send", methods=["POST"])
    def mech_followup_send():
        """Copy one nudge to the clipboard and log the touch."""
        try:
            ws, cols = _mech_state["ws"], _mech_state["cols"]
            if ws is None:
                return jsonify({"error": "Load your mechanics Excel on this tab first."}), 400
            row = int((request.get_json(force=True) or {}).get("row", 0))
            if row < 2:
                return jsonify({"error": "Bad row."}), 400

            # Re-checked at send time, not just when the list was built: a
            # reply can land in between, and nudging someone who already
            # answered is the one mistake this queue exists to prevent.
            if cols.get("replied") is not None and str(
                    ws.cell(row=row, column=cols["replied"] + 1).value or ""
            ).strip().lower() in _REPLIED_YES:
                return jsonify({"error": "This lead has already replied - no follow-up sent."}), 400

            fups = dm_app._to_int(ws.cell(row=row, column=cols["followups"] + 1).value) \
                if cols.get("followups") is not None else 0
            if fups >= dm_app.FOLLOWUP_TOUCHES:
                return jsonify({"error": "The sequence is already finished for this lead."}), 400

            name = str(ws.cell(row=row, column=cols["name"] + 1).value or "").strip()
            url = str(ws.cell(row=row, column=cols["url"] + 1).value or "").strip()
            text = mech_followup_text(fups + 1, name)

            if cols.get("followups") is not None:
                ws.cell(row=row, column=cols["followups"] + 1, value=fups + 1)
            if cols.get("sent_date") is not None:
                # The cadence counts from the LAST touch, so this has to move.
                ws.cell(row=row, column=cols["sent_date"] + 1, value=dm_app._today_str())
            _save()

            clipboard = dm_app.copy_to_clipboard(text)
            return jsonify({"ok": True, "text": text, "url": url,
                            "clipboard": clipboard, "touch": fups + 1})
        except Exception as e:
            import traceback; traceback.print_exc()
            return jsonify({"error": str(e)}), 500

    @app.route("/mech_replies_pending", methods=["POST"])
    def mech_replies_pending():
        """Every messaged mechanic with no answer logged yet, plus reply rate.

        No due-date filter: a reply that lands on day 0 has to have somewhere
        to go, otherwise the rate has no denominator and no copy decision here
        is ever falsifiable.
        """
        try:
            ws, cols = _mech_state["ws"], _mech_state["cols"]
            if ws is None:
                return jsonify({"error": "Load your mechanics Excel on this tab first."}), 400
            if cols.get("replied") is None:
                return jsonify({"error": "This sheet has no Replied? column."}), 400
            from datetime import date
            today = date.today()
            pending, sent_total, replied_total = [], 0, 0
            for r in ws.iter_rows(min_row=2):
                row = r[0].row
                url = _cell(r, cols, "url")
                if not url.startswith("http"):
                    continue
                if _cell(r, cols, "contacted").lower() not in _SENT:
                    continue
                sent_total += 1
                replied = _cell(r, cols, "replied").lower()
                if replied:
                    # Any value counts as answered for the denominator, but only
                    # an affirmative counts as a reply, so an explicit "No" can
                    # be logged without inflating the rate.
                    if replied in _REPLIED_YES:
                        replied_total += 1
                    continue
                last = dm_app._parse_date(r[cols["sent_date"]].value) if cols.get("sent_date") is not None else None
                pending.append({"row": row, "url": url,
                                "name": _cell(r, cols, "name"),
                                "niche": _cell(r, cols, "niche"),
                                "days": (today - last).days if last else None,
                                "sent": last.isoformat() if last else ""})
            pending.sort(key=lambda x: (x["days"] is None, x["days"] if x["days"] is not None else 0))
            answered = sent_total - len(pending)
            return jsonify({"ok": True, "pending": pending, "count": len(pending),
                            "sent": sent_total, "replied": replied_total,
                            "logged": answered,
                            "rate": round(replied_total / answered * 100, 1) if answered else None})
        except Exception as e:
            import traceback; traceback.print_exc()
            return jsonify({"error": str(e)}), 500

    @app.route("/mech_mark_reply", methods=["POST"])
    def mech_mark_reply():
        """Log a yes/no answer. An explicit No is what empties the pending list
        and gives the reply rate a denominator."""
        try:
            ws, cols = _mech_state["ws"], _mech_state["cols"]
            if ws is None:
                return jsonify({"error": "Load your mechanics Excel on this tab first."}), 400
            data = request.get_json(force=True) or {}
            row = int(data.get("row", 0))
            if row < 2 or cols.get("replied") is None:
                return jsonify({"error": "Can't mark this row."}), 400
            ws.cell(row=row, column=cols["replied"] + 1,
                    value="Yes" if data.get("replied") else "No")
            _save()
            return jsonify({"ok": True})
        except Exception as e:
            return jsonify({"error": str(e)}), 500

    @app.route("/mech_mark_sent", methods=["POST"])
    def mech_mark_sent():
        try:
            ws, cols = _mech_state["ws"], _mech_state["cols"]
            row = _mech_state["current_row"]
            if ws is None or row is None:
                return jsonify({"ok": False, "error": "Nothing to mark."}), 400
            if cols["contacted"] is not None:
                ws.cell(row=row, column=cols["contacted"] + 1, value="Yes")
            if cols.get("sent_date") is not None:
                ws.cell(row=row, column=cols["sent_date"] + 1, value=dm_app._today_str())
            _save()

            _mech_state["sent_today"] += 1
            _mech_state["current_row"] = None
            total, sent = _counts(ws, cols)
            return jsonify({"ok": True, "sent_today": _mech_state["sent_today"],
                            "sent": sent, "total": total})
        except Exception as e:
            return jsonify({"ok": False, "error": str(e)}), 500

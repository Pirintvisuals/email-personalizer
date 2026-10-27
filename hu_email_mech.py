"""
Autószerelő (mechanics) mode for the HU email sender (hu_email_app.py).

The builders' pipeline in hu_email_gen.py stays exactly as it is. This file
only replaces the two things that differ for mechanics:
  * the OPENER: built from real services printed on the shop's own website
  * the BODY: the mechanic-side pitch (he diagnoses the car, types in the
    fault and the VIN, the parts come from the supplier, out comes the price)

WHY THE OPENER IS DETERMINISTIC (no AI, no screenshot)
The whole email rests on "ezek alapján összeraktam Önnek egy demót". That
line is only believable if the facts in front of it are true. A mechanic who
reads a service he does not offer stops believing the rest of the message.
So every fact comes from a fixed keyword list matched against his own site,
never from a model writing freely. No real fact found = a plain opener that
claims nothing ("Megnéztem a szerviz oldalát." / "Rátaláltam az Ön
autószervizére."), never an invented detail.

Copy rules (agreed 2026-09-25):
  * magázás, no em dash, no link, no time figure ("X perc helyett")
  * the benefit is what he GETS (paid hours back, free evenings), not
    "faster quoting"
  * close stays small: "Csak megmutatnám, mást nem kérek. Elküldhetem?"
"""

import random
import re
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup

import hu_email_gen as heg


def _free_domains():
    # Lazy: hu_email_app imports this module, so a top-level import would loop.
    import hu_email_app
    return hu_email_app._FREE_MAIL_DOMAINS

# ── Services a shop might list, most specific first ────────────────────────
# (keywords to find on the page, accusative phrase used in the sentence).
# Order is priority: the rarer, more telling services come first, so the
# opener names what makes THIS shop different, not "olajcsere", which every
# shop in the country does.
_SERVICES = [
    # Doing inspections and preparing cars for one are different services,
    # and a shop that only prepares would read "vizsgáztatást" as a mistake.
    (("műszaki vizsgáztat", "vizsgabázis", "vizsgaállomás"),
                                                         "műszaki vizsgáztatást"),
    (("műszaki vizsgára felkészít", "műszaki vizsgára való felkészít",
      "vizsgafelkészít", "vizsgára felkészít"),          "műszaki vizsgára felkészítést"),
    (("műszaki vizsga",),                                "műszaki vizsgát"),
    (("eredetiségvizsgálat", "eredetiség vizsgálat",
      "eredetiség-vizsgálat"),                            "eredetiségvizsgálatot"),
    (("dpf tisztítás", "dpf-tisztítás", "részecskeszűrő tisztítás",
      "részecskeszűrő-tisztítás"),                        "részecskeszűrő-tisztítást"),
    (("turbó javítás", "turbójavítás", "turbófelújítás", "turbó felújítás"),
                                                         "turbójavítást"),
    (("automata váltó", "automataváltó", "váltójavítás", "váltó javítás",
      "váltófelújítás"),                                  "váltójavítást"),
    (("klímatöltés", "klíma töltés", "klímaszerviz", "klíma szerviz",
      "klímajavítás", "klíma javítás", "autóklíma"),     "klímaszervizt"),
    (("futóműállítás", "futómű állítás", "futómű beállítás",
      "futóműbeállítás", "futómű-beállítás"),            "futóműállítást"),
    (("autóvillamosság", "autó villamosság", "autóvillamossági"),
                                                         "autóvillamossági javítást"),
    (("karosszériajavítás", "karosszéria javítás", "karosszérialakatos",
      "karosszéria-javítás"),                             "karosszériajavítást"),
    (("fényezés", "autófényezés"),                       "fényezést"),
    (("gumiszerviz", "gumiszerelés", "gumicsere", "abroncsszerviz"),
                                                         "gumiszervizt"),
    (("vezérlés csere", "vezérléscsere", "vezérműszíj", "vezérlés-csere"),
                                                         "vezérléscserét"),
    (("kuplungcsere", "kuplung csere", "kuplungjavítás", "tengelykapcsoló"),
                                                         "kuplungcserét"),
    (("hibakódolvasás", "hibakód olvasás", "computeres diagnosztika",
      "számítógépes diagnosztika", "diagnosztika"),       "számítógépes diagnosztikát"),
    (("autómentés", "autómentő", "gépjárműmentés", "trélerezés"),
                                                         "autómentést"),
    (("szélvédőcsere", "szélvédő csere", "szélvédőjavítás"),
                                                         "szélvédőcserét"),
]

_NEGATIONS = ("nem ", "nincs ", "nélkül", "sajnos", "jelenleg nem")

# Links worth one extra fetch: small shop sites list their services on a
# separate page, and the homepage alone often says little beyond a phone
# number.
_SERVICE_PAGE_WORDS = ("szolgáltatás", "szolgaltatas", "szolgáltatásaink",
                       "szolgaltatasaink", "munkáink", "rólunk", "rolunk")


def _services_page_text(url: str) -> str:
    """Text of the shop's services page, or '' if none is linked. One extra
    plain-HTTP request, no browser."""
    if not url:
        return ""
    if not url.startswith(("http://", "https://")):
        url = "https://" + url.strip()
    try:
        resp = requests.get(url, headers=heg.HEADERS, timeout=heg.REQUEST_TIMEOUT)
        soup = BeautifulSoup(resp.text, "lxml")
        link = None
        for a in soup.find_all("a", href=True):
            label = f"{a.get_text(' ')} {a['href']}".lower()
            if any(w in label for w in _SERVICE_PAGE_WORDS):
                link = urljoin(resp.url, a["href"])
                break
        if not link or link.rstrip("/") == resp.url.rstrip("/"):
            return ""
        sub = requests.get(link, headers=heg.HEADERS, timeout=heg.REQUEST_TIMEOUT)
        s2 = BeautifulSoup(sub.text, "lxml")
        for tag in s2(["script", "style", "head", "noscript", "iframe", "svg"]):
            tag.decompose()
        body = s2.find("body")
        return heg._clean(body.get_text(separator="\n") if body else "")[:8000]
    except Exception:
        return ""


def find_services(text: str, limit: int = 2) -> list:
    """Accusative service phrases actually printed on the page, in priority
    order. A keyword sitting right after a negation ("nem vállalunk
    fényezést") is ignored, since repeating a denial back as a fact is worse
    than saying nothing."""
    t = (text or "").lower()
    found = []
    for keywords, phrase in _SERVICES:
        # One inspection phrase at most ("vizsgáztatást és műszaki vizsgát").
        if phrase.startswith("műszaki vizsg") and any(
                f.startswith("műszaki vizsg") for f in found):
            continue
        for kw in keywords:
            idx = t.find(kw)
            if idx == -1:
                continue
            # Same sentence only: "Fényezést nem vállalunk. Futómű állítás"
            # must not see the previous sentence's "nem".
            window = re.split(r"[.!?\n]", t[max(0, idx - 25):idx])[-1]
            if any(n in window for n in _NEGATIONS):
                continue
            # Hungarian puts the negation after the object just as often:
            # "Fényezést nem vállalunk."
            after = t[idx + len(kw):idx + len(kw) + 30]
            if re.search(r"\bnem\s+(vállal|végez|végzünk|foglalkoz|csinál)", after):
                continue
            found.append(phrase)
            break
        if len(found) >= limit:
            break
    return found


def _city_locative(city: str) -> str:
    """'Tarcalon' style form, or '' when the city is not in the known list.
    dm_app_hu._city_to_ragos returns the bare name for unknown cities, and
    "vállalnak Tarcal" is broken Hungarian, so an unknown city is left out
    rather than guessed."""
    city = (city or "").strip()
    ragos = heg.dm_app_hu._CITY_RAGOS.get(city)
    return ragos or ""


def build_opener(services: list, years: str, city: str) -> str:
    """'Megnéztem a szerviz oldalát. Láttam, hogy ...' or '' when there is
    nothing concrete to say."""
    if not services:
        return ""
    loc = _city_locative(city)
    loc_part = f" {loc}" if loc else ""
    if years and len(services) == 1:
        seen = f"{years} működnek, és {services[0]} is vállalnak{loc_part}"
    elif len(services) >= 2:
        seen = (f"a javítás mellett {services[0]} és {services[1]} is "
                f"vállalnak{loc_part}")
    else:
        seen = f"a javítás mellett {services[0]} is vállalnak{loc_part}"
    return f"Megnéztem a szerviz oldalát. Láttam, hogy {seen}."


# Mailbox providers and ISPs: the domain is not the shop's own website.
_NOT_A_SHOP_DOMAIN = {"wsi.hu", "t-online.hu", "invitel.hu", "digi.hu",
                      "upcmail.hu", "chello.hu", "freemail.hu", "citromail.hu",
                      "externet.hu", "vnet.hu", "axelero.hu", "enternet.hu"}


def website_from_email(email: str, free_domains=()) -> str:
    """'https://lalisgarage.com' for szerviz@lalisgarage.com; '' for a gmail
    or ISP address. Mechanics lists often carry the address but no website,
    and a shop's own domain almost always serves its site."""
    dom = (email or "").strip().lower().rpartition("@")[2]
    if not dom or dom in free_domains or dom in _NOT_A_SHOP_DOMAIN:
        return ""
    return f"https://{dom}"


def mech_subject(first: str) -> str:
    f = first.lower() if first else ""
    if f:
        pool = [
            f"{f}, összeraktam Önnek valamit",
            f"{f}, az Ön szervizére szabva",
            f"{f}, az árajánlatokról",
        ]
    else:
        pool = [
            "összeraktam Önnek valamit",
            "az Ön szervizére szabva",
            "az árajánlatokról",
        ]
    return heg._pick_subject(pool)


def mech_body(opener: str, first: str, city: str = "", from_site: bool = True) -> str:
    """The mechanics email, user's copy of 2026-10-07. Fixed wording.

    The mechanism paragraph answers the two objections a mechanic raises:
    the customer cannot diagnose his own car (the shop's intake person fills
    it in after the inspection), and a generic tool cannot know his prices
    (it uses his brands, supplier prices and labour rate).

    from_site is kept for callers but no longer changes the wording: the new
    copy says "Összeraktam" for every lead."""
    blocks = [f"Kedves {first}!" if first else "Jó napot kívánok!"]
    blocks.append(opener)
    blocks += [
        "Összeraktam Önnek egy árajánlatkészítő demót, kifejezetten az Ön "
        "szervizére szabva.",

        "Nem az ügyfél tölti ki, és nem ő mondja a hibát. Átvizsgálás után a "
        "munkafelvevő megadja az autót és a munkát, a program pedig "
        "összerakja a tételes ajánlatot az Ön márkáival, beszállítói áraival "
        "és munkadíjával. Mindig Ön nézi át, mielőtt kimegy.",

        "Amíg árat számol, nem szerel, és ezt az időt senki nem fizeti ki. "
        "Ezt kapja vissza.",
    ]
    if "miskolc" in (city or "").lower():
        blocks += [
            "Egyébként én is miskolci vagyok.",
            "Ha szeretné, akár személyesen is összefuthatunk.",
        ]
    blocks += [
        "Csak megmutatnám, mást nem kérek. Elküldhetem?",
        "Ha nem kéri, válaszoljon egy szóval, és többé nem keresem.",
        # No name here: the Gmail signature below already carries it.
        "Üdvözlettel,",
    ]
    return "\n\n".join(blocks)


# No service found on the site, or no site at all: the email still goes out
# (user's call, 2026-10-08), with an opener that claims nothing it can't back.
PLAIN_OPENER_SITE = "Megnéztem a szerviz oldalát."
PLAIN_OPENER_NO_SITE = "Rátaláltam az Ön autószervizére."


def process_lead(lead: dict) -> dict:
    """Mechanics version of hu_email_gen.process_lead. Same output keys, plus
    'blocked', which this version always leaves empty."""
    tag = lead.get("name") or lead.get("email") or f"row {lead['row']}"
    print(f"  [mech {lead['row']}] {tag}")

    website = lead.get("website", "") or website_from_email(
        lead.get("email", ""), free_domains=_free_domains())
    site = heg.fetch_site_text(website)
    first = heg._first_name_for(lead)
    lead["niche_resolved"] = "autószerelő"
    lead["hero_image"] = None
    lead["error"] = site.get("error", "")
    lead["blocked"] = ""

    # Q-Service leads are sendable again (user's call, 2026-10-08). The copy
    # never names Q-Service any more, which is what caused their circular.

    text = site.get("text", "") if site["ok"] else ""
    # A domain that serves an ISP portal or a parked page is not his site.
    if text and "autó" not in text.lower() and "auto" not in text.lower():
        text = ""
    services = find_services(text) if text else []
    if text and len(services) < 2:
        extra = _services_page_text(site["url"])
        if extra:
            text = text + "\n" + extra
            services = find_services(text)
    years = heg._find_years_in_business(text)

    # Same free address check as the builders' pipeline.
    t_l = text.lower()
    addr = (lead.get("email", "") or "").strip().lower()
    lead["email_on_site"] = ("unknown" if not text
                             else "confirmed" if addr and addr in t_l else "not_found")

    opener = build_opener(services, years, lead.get("city", ""))
    from_site = bool(opener)
    if not opener:
        opener = PLAIN_OPENER_SITE if text else PLAIN_OPENER_NO_SITE

    lead["opener"] = opener
    lead["subject"] = mech_subject(first)
    lead["body"] = mech_body(opener, first, lead.get("city", ""), from_site=from_site)
    print(f"    -> {opener[:90]}")
    return lead


# ── Follow-ups (same +3 / +7 / +14 schedule as hu_followup.SCHEDULE) ───────
# The builders' follow-ups talk about wasted site visits and "a bot" — the
# wrong product for a mechanic. Same shape, mechanic wording.

def followup_body(step: int, first: str) -> str:
    g = f"Kedves {first}!" if first else "Jó napot kívánok!"
    if step == 1:
        blocks = [
            g,
            random.choice([
                "Csak felhozom a korábbi levelemet, hátha elsüllyedt.",
                "Csak visszajelzek a múltkori levelemre, hátha lecsúszott.",
            ]),
            random.choice([
                "A demó, amit az Ön szervizére raktam össze, megvan. Küldhetem?",
                "A demó megvan Önnek, csak szólnia kell. Elküldhetem?",
            ]),
            "Üdvözlettel,",
        ]
    elif step == 2:
        blocks = [
            g,
            random.choice([
                "Egy dolgot kérdeznék: hetente hány árajánlatot ír olyannak, "
                "akiből végül nem lesz munka?",
                "Egy kérdés csak: hány ajánlatot számol ki egy héten, amiből "
                "végül nem lesz javítás?",
            ]),
            "Ez az az idő, amit senki nem fizet ki. Ezt veszi le Önről.",
            "Megnézi a demót?",
            "Üdvözlettel,",
        ]
    else:
        blocks = [
            g,
            random.choice([
                "Nem szeretnék tolakodni, ez az utolsó levelem ebben az ügyben.",
                "Nem akarok tolakodni, ezért ez az utolsó levelem emiatt.",
            ]),
            random.choice([
                "Ha most nem aktuális, az teljesen rendben van.",
                "Ha most nem téma, azt is teljesen megértem.",
            ]),
            "Ha később mégis eszébe jut, elég egy szó, és küldöm a demót.",
            "Üdvözlettel,",
        ]
    return "\n\n".join(blocks)

"""
Hungarian follow-up sequence — copy + schedule, no I/O.

WHY THIS EXISTS
The first-touch email was the whole system until now, which meant throwing
away most of the campaign's potential: across the large 2026 datasets, ~42%
of all replies arrive on a follow-up rather than the opening email, a single
follow-up lifts total replies by ~66%, and sequences that follow up at all
run 8.3% reply against 4.1% for one-touch. Nothing about the opening email
was the constraint — its absence of a second touch was.

SHAPE OF THE SEQUENCE
Three steps at +3, +7 and +14 days. Instantly's 16.5M-email study puts the
per-step reply peak at follow-up 2 with a ~17% drop by follow-up 3, so this
stops at three: the third is a breakup, which earns its place by being the
one message that reliably reactivates a silent thread precisely because it
asks for nothing.

Each step is SHORTER than the one before (36 → 33 → 30 words). A follow-up
competes with the original, still sitting right above it in the same thread —
restating the pitch reads as not having noticed the silence. Each step
therefore does exactly one job:
    1  a bump: assume it got buried, re-offer, nothing new
    2  a new angle: the wasted-callout cost, asked as a question so there is
       something concrete to answer
    3  a breakup: close the loop, remove the pressure, leave the door open

RULES THAT APPLY TO EVERY STEP
- Magázás, same as the first touch.
- No link, no image, no attachment.
- Sent INTO THE ORIGINAL THREAD (see gmail_send.send_email's thread_id /
  in_reply_to) so it reads as a follow-up rather than a second cold email.
- Never sent to anyone who replied — hu_email_app checks the thread first.
- Subject is left to the caller, which reuses the original with a "Re: "
  prefix; a fresh subject would break threading in most clients.
"""

import random
from datetime import date

# (step number, days after the FIRST-touch send, label for the UI)
SCHEDULE = [
    (1,  3, "Emlékeztető"),
    (2,  7, "Új szempont"),
    (3, 14, "Lezárás"),
]

MAX_STEP = SCHEDULE[-1][0]


def _greet(first: str) -> str:
    return f"Kedves {first}!" if first else "Jó napot kívánok!"


def _vary(options: list) -> str:
    return random.choice(options)


def step_body(step: int, first: str, niche: str = "") -> str:
    """The follow-up text for a given step. Same block-join convention as
    hu_email_gen._hu_email_body: one thought per paragraph, blank line
    between, no em dashes."""
    g = _greet(first)

    if step == 1:
        # A bump adds no new argument on purpose. Its entire job is to
        # resurface the thread for someone who saw it on a scaffold and
        # meant to come back to it.
        blocks = [
            g,
            _vary([
                "Csak felhozom a korábbi levelemet, hátha elsüllyedt.",
                "Csak visszajelzek a múltkori levelemre, hátha lecsúszott róla.",
            ]),
            _vary([
                "Áll még, hogy összeraktam Önnek egy mintát. Küldhetem?",
                "A minta megvan Önnek, csak szólnia kell. Elküldhetem?",
            ]),
            "Üdvözlettel,",
        ]

    elif step == 2:
        # The one step that introduces something new. It leads with a
        # question the reader can answer from memory in two seconds, which
        # is a far lower bar than "do you want to see a demo".
        blocks = [
            g,
            _vary([
                "Egy dolgot kérdeznék: havonta hány olyan kiszállása van, "
                "amiből végül nem lett munka?",
                "Egy kérdés csak: havonta hányszor megy ki úgy felmérni, "
                "hogy abból végül nem lesz semmi?",
            ]),
            _vary([
                "A legtöbb helyen ez a legdrágább óra a hónapban, mert nem "
                "számlázható.",
                "A legtöbb cégnél ez a leginkább elpazarolt idő, mert nem "
                "lehet kiszámlázni.",
            ]),
            "Pont ezt szűri ki a bot, még a telefonhívás előtt.",
            "Érdekli, hogy nézne ki az Ön oldalán?",
            "Üdvözlettel,",
        ]

    else:
        # Breakup. No pitch, no ask, no "last chance" pressure — the reason
        # this step works is that it is genuinely the end, and reads as it.
        blocks = [
            g,
            _vary([
                "Nem szeretnék tolakodni, ez az utolsó levelem ebben az ügyben.",
                "Nem akarok tolakodni, ezért ez az utolsó levelem emiatt.",
            ]),
            _vary([
                "Ha most nem aktuális, az teljesen rendben van.",
                "Ha most nem téma, azt is teljesen megértem.",
            ]),
            "Ha később mégis eszébe jut, elég egy szó, és küldöm a mintát.",
            "Üdvözlettel,",
        ]

    return "\n\n".join(blocks)


def followup_subject(original_subject: str) -> str:
    """Reply-style subject on the ORIGINAL line. Threading in most clients
    keys partly off the subject, and a new one would start a new
    conversation — which is the one thing a follow-up must not do."""
    s = (original_subject or "").strip()
    if not s:
        return "Re:"
    return s if s.lower().startswith("re:") else f"Re: {s}"


def _delay_for(step: int) -> int:
    for s, delay, _ in SCHEDULE:
        if s == step:
            return delay
    return 10 ** 6


def min_gap(step: int) -> int:
    """Days that must pass since the PREVIOUS follow-up before this one is
    allowed — the sequence's own intended spacing (F2 comes 4 days after F1,
    F3 comes 7 after F2)."""
    return _delay_for(step) - _delay_for(step - 1) if step > 1 else _delay_for(1)


def due_step(first_sent: date, steps_done: int, today: date = None,
             last_followup: date = None) -> int:
    """Which follow-up step is due for a lead, or 0 if none is.

    Two clocks have to agree, not one. Measuring only from the first touch
    looks right until a lead goes stale: an 8-day-old lead with no follow-ups
    satisfies BOTH the +3 and the +7 rule at once, so it would take F1 and F2
    on the same afternoon, in the same thread, minutes apart. That is worse
    than never following up. So a step is due only when its day has arrived
    counting from the first touch AND enough time has passed since the
    previous follow-up actually went out (see min_gap). A backlog therefore
    drains one message per lead per few days, exactly as if it had been on
    schedule all along."""
    if first_sent is None or steps_done >= MAX_STEP:
        return 0
    today = today or date.today()
    nxt = steps_done + 1
    if (today - first_sent).days < _delay_for(nxt):
        return 0
    if steps_done and last_followup is not None:
        if (today - last_followup).days < min_gap(nxt):
            return 0
    return nxt


def step_label(step: int) -> str:
    for s, delay, label in SCHEDULE:
        if s == step:
            return f"{label} (+{delay} nap)"
    return f"{step}. utánkövetés"

"""
Real Gmail sending via the Gmail API (OAuth) — separate module, imported by
hu_email_app.py. Replaces the "open a pre-filled compose tab, click send
yourself" trick for any email that carries a hero image: that trick's body
comes through a URL query parameter, which can only ever be plain text —
there is no way to embed an inline image or any HTML through it. A real
image-in-email needs a real MIME message sent through the API.

FIRST RUN: calling get_service() opens a browser window once for you to grant
access. Log in as the SENDING account (milan@landscale.net), not your personal
Gmail — whichever account you authorize as is the account mail gets sent
FROM. If milan@landscale.net isn't the account that created the Cloud
project, add it as a Test User first: Google Cloud Console -> APIs & Services
-> OAuth consent screen -> Test users -> Add.

After that first grant, the token is cached in token_hu_email.json (gitignored)
and every later send is silent — no browser popup, no re-login.

Scope is gmail.send ONLY, on purpose. This cannot read your inbox, contacts,
or anything else — only send mail as you. Reply-detection for the follow-up
engine needs a broader scope later; that will be its own explicit re-auth
when that gets built, not bundled into this.
"""

import base64
import json
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

_HERE = Path(__file__).parent
CREDENTIALS_PATH = _HERE / "credentials.json"

# gmail.send alone cannot answer "did they write back?", and without that the
# follow-up engine would keep chasing people who already replied — the single
# worst thing an automated sequence can do.
#
# gmail.metadata, NOT gmail.readonly: metadata grants message headers and
# thread structure only, never message bodies. That is exactly enough to see
# that a thread now contains an inbound message, and nothing more — the actual
# reply still gets read in Gmail like any other mail. Adding this scope
# invalidates the existing send-only tokens, so get_service() below detects
# the mismatch and re-runs the consent flow once per account instead of
# failing with an opaque refresh error.
SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.metadata",
]

# Multiple sending identities, each its own OAuth grant (one credentials.json
# OAuth *client* is shared, but each account needs its own token file — a
# token is tied to whichever Google account you log in as during the
# authorization popup, not to this file). Add an account here, then run
# gmail_send.py directly once to trigger that account's one-time login.
ACCOUNTS = {
    "landscale": {
        "email": "milan@landscale.net",
        "token_path": _HERE / "token_hu_email.json",
        "label": "Landscale (milan@landscale.net)",
        # Paid Workspace on its own domain — can keep ramping (the warmup
        # schedule for each type lives in hu_email_app._RAMP_SCHEDULES).
        "type": "workspace",
    },
    "gmail": {
        "email": "landscale.agency@gmail.com",
        "token_path": _HERE / "token_hu_email_gmail.json",
        "label": "Gmail (landscale.agency@gmail.com)",
        # Free consumer Gmail — the safe cold-email ceiling is ~25/day
        # FOREVER, not just during warmup, and a spam-complaint pattern can
        # lock the whole account with no admin console to appeal through.
        "type": "free_gmail",
    },
}
DEFAULT_ACCOUNT = "landscale"

# Kept for any old caller that still imports these directly.
SENDER_EMAIL = ACCOUNTS[DEFAULT_ACCOUNT]["email"]
TOKEN_PATH = ACCOUNTS[DEFAULT_ACCOUNT]["token_path"]


def _granted_scopes(token_path) -> set:
    """Scopes actually written in the token file.

    Read from the JSON directly rather than through
    Credentials.from_authorized_user_file(path, SCOPES): that constructor
    OVERRIDES the stored scopes with whatever list you hand it, so the
    resulting object always reports exactly the scopes you asked about and
    creds.has_scopes(SCOPES) comes back True even for a token that was never
    granted them. Checking the file is the only way to learn what the grant
    really covers."""
    try:
        with open(token_path, "r", encoding="utf-8") as fh:
            return set(json.load(fh).get("scopes") or [])
    except Exception:
        return set()


def missing_scopes(account: str = DEFAULT_ACCOUNT) -> bool:
    """True if this account has no token yet, or holds one granted before
    gmail.metadata was added. Lets the UI say 'this account needs a one-time
    re-login' instead of blowing up mid-send."""
    token_path = ACCOUNTS[account]["token_path"]
    if not token_path.exists():
        return True
    return not set(SCOPES).issubset(_granted_scopes(token_path))


def token_revoked(account: str = DEFAULT_ACCOUNT) -> bool:
    """True if a token FILE exists with the right scopes but the grant behind
    it is dead (revoked, or expired past recovery).

    missing_scopes() cannot see this. It reads the scope list out of the JSON,
    and a revoked token still lists exactly the scopes it was granted — so a
    dead token looks perfectly healthy on disk and reauth_email.py skipped it
    with "ezek mar keszen vannak" while every send failed with
    'invalid_grant: Token has been expired or revoked.'

    The only way to know is to ask Google. This does a real refresh against
    the token endpoint; anything that comes back as a refresh failure means a
    fresh consent flow is required.
    """
    token_path = ACCOUNTS[account]["token_path"]
    if not token_path.exists():
        return False          # nothing to be revoked; missing_scopes covers it
    try:
        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)
    except Exception:
        return True           # unreadable/corrupt -> needs re-auth
    if not creds.refresh_token:
        return True
    try:
        creds.refresh(Request())
        return False
    except Exception:
        return True


def needs_reauth(account: str = DEFAULT_ACCOUNT) -> bool:
    """Single question the re-login script should ask: does this account need
    a browser consent round? Covers both 'no/old token' and 'dead token'."""
    return missing_scopes(account) or token_revoked(account)


def get_service(account: str = DEFAULT_ACCOUNT):
    acct = ACCOUNTS[account]
    token_path, sender_email = acct["token_path"], acct["email"]

    creds = None
    if token_path.exists():
        try:
            creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)
        except Exception:
            creds = None
    # A token issued for a SMALLER scope set refreshes fine but then 403s on
    # the first metadata call, which is a confusing place to discover the
    # problem. Treat it as "no token" so the consent flow runs now instead.
    # Checked against the FILE, not creds.has_scopes — see _granted_scopes.
    if creds and missing_scopes(account):
        print(f"{sender_email}: token predates the reply-detection scope — "
              f"one-time re-authorization needed.")
        creds = None
    if not creds or not creds.valid:
        refreshed = False
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                refreshed = True
            except Exception as e:
                # invalid_grant means Google killed the grant server-side; the
                # local JSON is intact but worthless. This used to propagate,
                # which made the failure unrecoverable BY DESIGN: the browser
                # consent flow lived in the else-branch, so the one code path
                # that could fix a dead token was the one path a dead token
                # could never reach. reauth_email.py would correctly identify
                # the account as needing re-login, open nothing, and report
                # invalid_grant. Drop the dead credentials and fall through.
                print(f"{sender_email}: stored authorization is no longer valid "
                      f"({str(e)[:90]}). Starting a fresh login.")
                creds = None
        if not refreshed:
            if not CREDENTIALS_PATH.exists():
                raise FileNotFoundError(
                    f"{CREDENTIALS_PATH} not found — download the OAuth client JSON "
                    "from Google Cloud Console and save it there as credentials.json.")
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_PATH), SCOPES)
            print(f"Opening browser for one-time Gmail authorization — "
                  f"log in as {sender_email} (not any other account).")
            creds = flow.run_local_server(port=0)
        token_path.write_text(creds.to_json())
    return build("gmail", "v1", credentials=creds)


def _text_to_html(plain_body: str) -> str:
    """Turn the '\\n\\n'-separated blocks _hu_email_body() builds into plain
    <p> tags. Deliberately undecorated — a heavily styled HTML email reads as
    marketing, which is exactly what a cold first-touch email should NOT look
    like next to the hero photo."""
    parts = [p.replace("\n", "<br>") for p in plain_body.split("\n\n") if p.strip()]
    inner = "".join(f"<p style='margin:0 0 16px;'>{p}</p>" for p in parts)
    # margin:0 auto centers the block itself in the message pane (the Charlie
    # Morgan look — a column of text floating in the middle with space on
    # both sides), same as the hero <img> below needs the matching change.
    return ("<div style=\"font-family:Arial,Helvetica,sans-serif;font-size:14px;"
            f"color:#111111;line-height:1.5;max-width:520px;margin:0 auto;\">{inner}</div>")


def build_message(to: str, subject: str, plain_body: str,
                   hero_image: bytes | None = None, sender: str = SENDER_EMAIL,
                   in_reply_to: str = ""):
    """Builds a TRUE plain-text message when there's no hero image — which is
    now every real send (see hu_email_gen.process_lead).

    This used to always build multipart/alternative with an HTML part beside
    the text, because the hero image needed HTML to render inline. With the
    image gone that HTML part bought nothing and cost real deliverability: a
    styled HTML body is one of the signals Gmail sorts on for the Promotions
    tab, and plain text consistently beats HTML-with-styling on inbox
    placement and reply rate for cold first-touch mail. A one-to-one cold
    email from a person should look like exactly that at the MIME level too,
    not just in its wording — multipart/alternative + a <div> with inline
    CSS is what bulk tooling emits, and it's visible to filters even when the
    rendered result looks identical to the human reading it.

    The hero_image branch below is kept intact and still works, so a later
    email in a sequence (after they've replied and engaged, when the
    deliverability maths is completely different) can still send the
    screenshot inline without rebuilding this."""
    def _thread_headers(m):
        # Passing threadId to the API alone makes GMAIL group the messages,
        # but the recipient's client threads on these RFC headers — Outlook
        # in particular ignores subject-matching. A follow-up that opens as
        # a brand new message instead of under the original is just a second
        # cold email, which is the whole thing this is trying not to be.
        if in_reply_to:
            m["In-Reply-To"] = in_reply_to
            m["References"] = in_reply_to
        return m

    if not hero_image:
        msg = MIMEText(plain_body, "plain", "utf-8")
        msg["To"] = to
        msg["From"] = sender
        msg["Subject"] = subject
        return _thread_headers(msg)

    msg = MIMEMultipart("related")
    msg["To"] = to
    msg["From"] = sender
    msg["Subject"] = subject
    _thread_headers(msg)

    alt = MIMEMultipart("alternative")
    msg.attach(alt)
    alt.attach(MIMEText(plain_body, "plain", "utf-8"))

    html = _text_to_html(plain_body)
    if hero_image:
        # The image used to be a sibling of the centered text div, so it had
        # no width limit of its own and stretched full-width instead of
        # lining up with the centered column beneath it. Wrapping both in one
        # 520px/margin:auto shell keeps the photo and the text the same width
        # and centered together.
        img_tag = ("<img src='cid:heroimg' style='width:100%;border-radius:6px;"
                   "margin-bottom:18px;display:block;' alt=''>")
        html = (f"<div style=\"max-width:520px;margin:0 auto;\">{img_tag}{html}</div>")
    alt.attach(MIMEText(html, "html", "utf-8"))

    if hero_image:
        img = MIMEImage(hero_image, _subtype="jpeg")
        img.add_header("Content-ID", "<heroimg>")
        img.add_header("Content-Disposition", "inline", filename="hero.jpg")
        msg.attach(img)

    return msg


def send_email(to: str, subject: str, plain_body: str,
               hero_image: bytes | None = None, account: str = DEFAULT_ACCOUNT,
               thread_id: str = "", in_reply_to: str = "") -> dict:
    """Sends immediately via the Gmail API, from whichever account was picked
    in the UI. Caller (hu_email_app.py's /send route) is what turns this into
    a single explicit per-email action — this function itself has no
    confirmation step, so it should only ever be called from a user-clicked
    Send.

    Returns a dict, not a bare id: the follow-up engine needs the threadId to
    ask later whether anyone replied, and the RFC Message-ID to thread the
    follow-up under this message in the recipient's client. Old callers that
    only wanted the id keep working via result["id"].

    thread_id/in_reply_to are what MAKE it a follow-up rather than a second
    cold email — pass both when continuing an existing conversation."""
    sender = ACCOUNTS[account]["email"]
    msg = build_message(to, subject, plain_body, hero_image, sender, in_reply_to)
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    body = {"raw": raw}
    if thread_id:
        body["threadId"] = thread_id

    service = get_service(account)
    sent = _execute(service.users().messages().send(userId="me", body=body))
    out = {"id": sent.get("id", ""), "thread_id": sent.get("threadId", ""),
           "message_id": ""}

    # Gmail assigns the real Message-ID itself; it isn't in the send response,
    # so read it back. Non-fatal: without it follow-ups still thread on
    # threadId for Gmail recipients, just less reliably for everyone else.
    try:
        meta = service.users().messages().get(
            userId="me", id=out["id"], format="metadata",
            metadataHeaders=["Message-ID"]).execute()
        for h in meta.get("payload", {}).get("headers", []):
            if h.get("name", "").lower() == "message-id":
                out["message_id"] = h.get("value", "")
                break
    except Exception as e:
        print(f"    (Message-ID lookup failed, threading may be weaker: {str(e)[:80]})")
    return out


def thread_subject(thread_id: str, account: str = DEFAULT_ACCOUNT) -> str:
    """Subject line of the thread's FIRST message.

    Read back rather than stored alongside the thread id, for two reasons:
    the subject is chosen at random per lead so it cannot be regenerated, and
    keeping it only in the spreadsheet would mean a hand-edited cell could
    silently break threading. Gmail already holds the authoritative copy and
    the metadata scope can see it."""
    if not thread_id:
        return ""
    service = get_service(account)
    thread = service.users().threads().get(
        userId="me", id=thread_id, format="metadata",
        metadataHeaders=["Subject"]).execute()
    for m in thread.get("messages", [])[:1]:
        for h in m.get("payload", {}).get("headers", []):
            if h.get("name", "").lower() == "subject":
                return h.get("value", "")
    return ""


# Senders that mean "this never reached a human". Matched as substrings of
# the From header, which is where every mainstream MTA puts them.
_BOUNCE_SENDERS = ("mailer-daemon", "postmaster@", "no-reply@dmarcreport",
                    "mail delivery subsystem", "mailerdaemon")


def _execute(req):
    """Run a Gmail API request, retrying the transient rate-limit answers.

    The reply sweep walks every open thread, and the UI used to re-run that
    whole sweep after each send, so a run of follow-ups fired hundreds of
    threads.get calls within a minute and hit "Quota exceeded for quota
    metric 'Total Query Cost'" (HTTP 403, reason rateLimitExceeded).

    That is a THROTTLE, not an auth failure: the identical call succeeds a
    second later. It matters that they are told apart, because
    send_followup() deliberately fails CLOSED on a status-check error, so an
    ordinary burst used to look like a broken login and block sending.

    Only quota/rate answers are retried. A 403 for missing scopes is a real
    failure and still raises immediately, so re-auth is never masked.
    """
    import random
    import time
    from googleapiclient.errors import HttpError

    delay = 1.0
    for attempt in range(4):
        try:
            return req.execute()
        except HttpError as e:
            status = getattr(getattr(e, "resp", None), "status", None)
            body = (e.content or b"")
            if isinstance(body, str):
                body = body.encode("utf-8", "replace")
            throttled = status in (429, 500, 503) or (
                status == 403 and (b"ateLimit" in body or b"uota" in body))
            if not throttled or attempt == 3:
                raise
            time.sleep(delay + random.uniform(0, 0.4))
            delay *= 2


def thread_status(thread_id: str, account: str = DEFAULT_ACCOUNT) -> str:
    """One pass over a thread's From headers -> 'reply', 'bounce' or ''.

    Both answers come from the same call on purpose: they are decided by the
    same data, and asking twice would double the API traffic for every open
    lead on every check.

    A real reply outranks a bounce if somehow both appear — a human writing
    back is the outcome that matters, and a stale delivery warning shouldn't
    bury it.

    Uses format='metadata', the only format the gmail.metadata scope permits:
    headers and structure, never message bodies."""
    if not thread_id:
        return ""
    sender = ACCOUNTS[account]["email"].lower()
    service = get_service(account)
    thread = _execute(service.users().threads().get(
        userId="me", id=thread_id, format="metadata",
        metadataHeaders=["From"]))

    bounced = False
    for m in thread.get("messages", []):
        frm = ""
        for h in m.get("payload", {}).get("headers", []):
            if h.get("name", "").lower() == "from":
                frm = h.get("value", "").lower()
                break
        if not frm or sender in frm:
            continue                      # our own first touch or follow-up
        if any(b in frm for b in _BOUNCE_SENDERS):
            bounced = True
            continue
        return "reply"
    return "bounce" if bounced else ""


def thread_has_reply(thread_id: str, account: str = DEFAULT_ACCOUNT) -> bool:
    """True if a human wrote back. Thin wrapper kept so existing callers and
    tests don't need to know about the bounce case."""
    return thread_status(thread_id, account) == "reply"


if __name__ == "__main__":
    # Manual smoke test — also how you do each account's one-time OAuth
    # login: run this, pick the account, and the browser popup will ask you
    # to log in as that account's address specifically.
    print("Fiókok:", ", ".join(f"{k} ({v['email']})" for k, v in ACCOUNTS.items()))
    acct = input(f"Melyik fiókkal ({'/'.join(ACCOUNTS)}) [{DEFAULT_ACCOUNT}]: ").strip() or DEFAULT_ACCOUNT
    test_to = input("Send a test email to (your own address): ").strip()
    mid = send_email(
        to=test_to,
        subject="teszt — gmail api",
        plain_body="Ez egy teszt üzenet a Gmail API integrációból.\n\nMásodik bekezdés, hogy lássuk a formázást is.",
        account=acct,
    )
    print(f"Elküldve {ACCOUNTS[acct]['email']}-ról. Message id: {mid}")

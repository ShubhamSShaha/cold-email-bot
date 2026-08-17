import base64
from datetime import datetime, timedelta, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import requests

import config
from auth import get_access_token

GMAIL = "https://gmail.googleapis.com/gmail/v1/users/me"
TIMEOUT = 30


def _headers():
    return {"Authorization": f"Bearer {get_access_token()}"}


def _raise_for_status(resp: requests.Response, context: str):
    if not resp.ok:
        detail = resp.text.strip() or "(empty response body)"
        raise RuntimeError(f"{context} failed [{resp.status_code}]: {detail}")


def _sender() -> str:
    return config.get("email", "sender")


def _to_html(text: str) -> str:
    STYLE = (
        "font-family:Arial,Helvetica,sans-serif;"
        "font-size:12pt;"
        "color:#000000;"
        "line-height:normal;"
        "margin:0;padding:0;"
    )
    blocks = []
    current = []
    for line in text.split("\n"):
        if line.strip() == "":
            if current:
                blocks.append(current)
                current = []
        else:
            current.append(line.strip())
    if current:
        blocks.append(current)

    html_parts = []
    for block in blocks:
        if all(l.startswith("* ") for l in block):
            items = "".join(f"<li>{l[2:]}</li>" for l in block)
            html_parts.append(
                f"<ul style='{STYLE} margin-left:18pt; margin-bottom:12pt;'>"
                f"{items}</ul>"
            )
        else:
            inner = "<br>".join(block)
            html_parts.append(
                f"<p style='{STYLE} margin-bottom:12pt;'>{inner}</p>"
            )
    return f"<div style='{STYLE}'>" + "".join(html_parts) + "</div>"


# ── Building & sending ─────────────────────────────────────────────────────────

def _build_message(to: str, subject: str, body: str) -> MIMEMultipart:
    msg = MIMEMultipart("alternative")
    msg["To"] = to
    msg["Subject"] = subject
    sender = _sender()
    if sender:
        msg["From"] = sender
    msg.attach(MIMEText(body, "plain", "utf-8"))
    msg.attach(MIMEText(_to_html(body), "html", "utf-8"))
    return msg


def _send_mime(msg: MIMEMultipart, thread_id: str = "") -> tuple[str, str]:
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    payload = {"raw": raw}
    if thread_id:
        payload["threadId"] = thread_id
    resp = requests.post(
        f"{GMAIL}/messages/send", headers=_headers(), json=payload, timeout=TIMEOUT
    )
    _raise_for_status(resp, "messages.send")
    data = resp.json()
    return data.get("id", ""), data.get("threadId", "")


def send_email(to: str, subject: str, body: str) -> tuple[str, str]:
    """Send a new email. Returns (message_id, thread_id).

    Unlike Graph's sendMail, this returns both IDs synchronously in the
    response — no polling Sent Items afterward to recover them.
    """
    return _send_mime(_build_message(to, subject, body))


# ── Reply in the same thread ──────────────────────────────────────────────────

def _thread_last_message_id_header(thread_id: str) -> str:
    resp = requests.get(
        f"{GMAIL}/threads/{thread_id}",
        headers=_headers(),
        params={"format": "metadata", "metadataHeaders": "Message-Id"},
        timeout=TIMEOUT,
    )
    _raise_for_status(resp, "threads.get")
    messages = resp.json().get("messages", [])
    if not messages:
        return ""
    for h in messages[-1].get("payload", {}).get("headers", []):
        if h.get("name", "").lower() == "message-id":
            return h.get("value", "")
    return ""


def reply_to_message(thread_id: str, body: str, to_email: str, subject: str = "") -> tuple[str, str]:
    """Send a follow-up as a true reply in the same Gmail thread."""
    if not thread_id:
        raise RuntimeError("reply_to_message requires a Gmail thread_id — run option 7 to backfill.")

    msgid_header = _thread_last_message_id_header(thread_id)
    reply_subject = subject if subject.lower().startswith("re:") else f"Re: {subject}"

    msg = MIMEMultipart("alternative")
    msg["To"] = to_email
    msg["Subject"] = reply_subject
    sender = _sender()
    if sender:
        msg["From"] = sender
    if msgid_header:
        msg["In-Reply-To"] = msgid_header
        msg["References"] = msgid_header
    msg.attach(MIMEText(body, "plain", "utf-8"))
    msg.attach(MIMEText(_to_html(body), "html", "utf-8"))

    return _send_mime(msg, thread_id=thread_id)


# ── Find a thread by recipient and subject ─────────────────────────────────────

def find_thread_id(to: str, subject: str) -> str:
    """Search Sent for a message matching to+subject. Returns threadId, or ''."""
    q = f'in:sent to:{to} subject:"{subject}"'
    resp = requests.get(
        f"{GMAIL}/messages", headers=_headers(), params={"q": q, "maxResults": 5}, timeout=TIMEOUT
    )
    if not resp.ok:
        return ""
    messages = resp.json().get("messages", [])
    return messages[0]["threadId"] if messages else ""


# ── Inbox scanning for replies ────────────────────────────────────────────────

_BOUNCE_MARKERS = [
    "undeliverable", "delivery failed", "delivery failure", "mail delivery",
    "returned mail", "bounce", "non-delivery", "could not be delivered", "failure notice",
]


def get_inbox_thread_ids(since_days: int = 30) -> set[str]:
    """Return Gmail threadIds seen in the inbox within the last since_days days. Skips bounces."""
    cutoff = (datetime.now(timezone.utc) - timedelta(days=since_days)).strftime("%Y/%m/%d")
    terms = " OR ".join(f'"{t}"' if " " in t else t for t in _BOUNCE_MARKERS)
    q = f"in:inbox after:{cutoff} -subject:({terms})"

    thread_ids = set()
    page_token = None
    while True:
        params = {"q": q, "maxResults": 100}
        if page_token:
            params["pageToken"] = page_token
        resp = requests.get(f"{GMAIL}/messages", headers=_headers(), params=params, timeout=TIMEOUT)
        _raise_for_status(resp, "messages.list")
        data = resp.json()
        thread_ids.update(m["threadId"] for m in data.get("messages", []))
        page_token = data.get("nextPageToken")
        if not page_token:
            break
    return thread_ids


# ── Connectivity probe / identity ──────────────────────────────────────────────

def get_profile() -> dict:
    """GET .../profile — doubles as connectivity probe and identity check."""
    resp = requests.get(f"{GMAIL}/profile", headers=_headers(), timeout=TIMEOUT)
    if not resp.ok:
        return {"ok": False, "detail": f"HTTP {resp.status_code}: {resp.text.strip() or '(empty body)'}"}
    data = resp.json()
    return {
        "ok": True,
        "email": data.get("emailAddress", ""),
        "messages_total": data.get("messagesTotal", "?"),
    }

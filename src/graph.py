import requests
from auth import get_access_token

GRAPH = "https://graph.microsoft.com/v1.0"


def _headers():
    return {
        "Authorization": f"Bearer {get_access_token()}",
        "Content-Type": "application/json",
    }


def _raise_for_status(resp: requests.Response, context: str):
    if not resp.ok:
        raise RuntimeError(f"{context} failed [{resp.status_code}]: {resp.text}")


def _to_html(text: str) -> str:
    OUTLOOK_STYLE = (
        "font-family:Aptos,Calibri,sans-serif;"
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
                f"<ul style='{OUTLOOK_STYLE} margin-left:18pt; margin-bottom:12pt;'>"
                f"{items}</ul>"
            )
        else:
            inner = "<br>".join(block)
            html_parts.append(
                f"<p style='{OUTLOOK_STYLE} margin-bottom:12pt;'>{inner}</p>"
            )
    return (
        f"<div style='{OUTLOOK_STYLE}'>"
        + "".join(html_parts)
        + "</div>"
    )


# ── Send a brand-new email ────────────────────────────────────────────────────

def send_email(to: str, subject: str, body: str) -> tuple[str, str]:
    """Send a new email. Returns (message_id, conversation_id)."""
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": _to_html(body)},
            "toRecipients": [{"emailAddress": {"address": to}}],
        },
        "saveToSentItems": True,
    }
    resp = requests.post(f"{GRAPH}/me/sendMail", headers=_headers(), json=payload)
    _raise_for_status(resp, f"send_email to {to}")
    msg_id, conv_id = _find_sent_message(to, subject)
    return msg_id, conv_id


def _find_sent_message(to: str, subject: str) -> tuple[str, str]:
    """Returns (message_id, conversation_id) of the most recently sent message."""
    params = {
        "$top": 20,
        "$select": "id,toRecipients,subject,conversationId",
    }
    resp = requests.get(
        f"{GRAPH}/me/mailFolders/SentItems/messages",
        headers=_headers(),
        params=params,
    )
    _raise_for_status(resp, "find_sent_message")
    items = resp.json().get("value", [])
    for item in items:
        if item.get("subject", "").strip() != subject.strip():
            continue
        for r in item.get("toRecipients", []):
            addr = r.get("emailAddress", {}).get("address", "")
            if addr.lower() == to.lower():
                return item["id"], item.get("conversationId", "")
    raise RuntimeError(f"Could not locate sent message to {to} with subject '{subject}'")


# ── Schedule an email ─────────────────────────────────────────────────────────

def schedule_email(to: str, subject: str, body: str, send_at: str) -> tuple[str, str]:
    """Schedule an email. Returns (draft_id, conversation_id)."""
    draft_payload = {
        "subject": subject,
        "body": {"contentType": "HTML", "content": _to_html(body)},
        "toRecipients": [{"emailAddress": {"address": to}}],
    }
    resp = requests.post(
        f"{GRAPH}/me/messages",
        headers=_headers(),
        json=draft_payload,
    )
    _raise_for_status(resp, f"create_draft to {to}")
    draft    = resp.json()
    draft_id = draft["id"]
    conv_id  = draft.get("conversationId", "")

    resp = requests.patch(
        f"{GRAPH}/me/messages/{draft_id}",
        headers=_headers(),
        json={"singleValueExtendedProperties": [{
            "id": "SystemTime 0x3FEF",
            "value": send_at + "Z"
        }]},
    )
    _raise_for_status(resp, f"patch_schedule to {to}")

    resp = requests.post(
        f"{GRAPH}/me/messages/{draft_id}/send",
        headers=_headers(),
    )
    _raise_for_status(resp, f"schedule_send to {to}")

    return draft_id, conv_id


# ── Find a valid message ID to reply to ──────────────────────────────────────

def _find_replyable_message(message_id: str, conversation_id: str) -> str:
    """
    Find a valid message ID to use for createReply.
    First tries the stored message_id directly.
    If that fails, pages through all Sent Items to find the conversation.
    """
    # Try stored message ID first
    if message_id:
        resp = requests.get(
            f"{GRAPH}/me/messages/{message_id}",
            headers=_headers(),
            params={"$select": "id"},
        )
        if resp.ok:
            return message_id

    # Page through all Sent Items to find by conversation_id
    if conversation_id:
        url = f"{GRAPH}/me/mailFolders/SentItems/messages"
        params = {
            "$top": 100,
            "$select": "id,conversationId,sentDateTime",
        }
        all_items = []
        while url:
            resp = requests.get(url, headers=_headers(), params=params)
            if not resp.ok:
                break
            data = resp.json()
            all_items.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
            params = {}

        matches = [
            i for i in all_items
            if i.get("conversationId") == conversation_id
        ]
        if matches:
            return max(matches, key=lambda x: x.get("sentDateTime", ""))["id"]

    raise RuntimeError(
        f"Could not find a valid message to reply to for conversation {conversation_id}"
    )


# ── Reply in the same thread ──────────────────────────────────────────────────

def reply_to_message(message_id: str, body: str, to_email: str, conversation_id: str = ""):
    """
    Send a follow-up as a true reply in the same thread.
    Finds the actual sent message from Sent Items using conversation_id
    if the stored message_id has expired (scheduled send case).
    """
    # Step 1: Find a valid replyable message ID
    reply_to_id = _find_replyable_message(message_id, conversation_id)

    # Step 2: Create reply draft
    resp = requests.post(
        f"{GRAPH}/me/messages/{reply_to_id}/createReply",
        headers=_headers(),
    )
    _raise_for_status(resp, "createReply")
    draft_id = resp.json()["id"]

    # Step 3: Patch body and recipient
    patch = {
        "body": {"contentType": "HTML", "content": _to_html(body)},
        "toRecipients": [{"emailAddress": {"address": to_email}}],
    }
    resp = requests.patch(
        f"{GRAPH}/me/messages/{draft_id}",
        headers=_headers(),
        json=patch,
    )
    _raise_for_status(resp, "patch draft reply")

    # Step 4: Send
    resp = requests.post(
        f"{GRAPH}/me/messages/{draft_id}/send",
        headers=_headers(),
    )
    _raise_for_status(resp, "send draft reply")


# ── Find conversation ID by recipient and subject ─────────────────────────────

def find_conversation_id(to: str, subject: str) -> str:
    """Search Sent Items for a message matching to+subject. Returns conversationId."""
    params = {
        "$top": 50,
        "$select": "id,toRecipients,subject,conversationId",
    }
    resp = requests.get(
        f"{GRAPH}/me/mailFolders/SentItems/messages",
        headers=_headers(),
        params=params,
    )
    if not resp.ok:
        return ""
    for item in resp.json().get("value", []):
        if item.get("subject", "").strip() != subject.strip():
            continue
        for r in item.get("toRecipients", []):
            addr = r.get("emailAddress", {}).get("address", "")
            if addr.lower() == to.lower():
                return item.get("conversationId", "")
    return ""


# ── Inbox scanning for replies ────────────────────────────────────────────────

def get_inbox_conversation_ids(since_days: int = 30) -> set[str]:
    """Return conversationIds found in inbox within last since_days days. Skips bounces."""
    from datetime import datetime, timedelta, timezone
    cutoff = (datetime.now(timezone.utc) - timedelta(days=since_days)).strftime(
        "%Y-%m-%dT%H:%M:%SZ"
    )
    conv_ids = set()
    url = f"{GRAPH}/me/mailFolders/Inbox/messages"
    params = {
        "$filter": f"receivedDateTime ge {cutoff}",
        "$select": "conversationId,subject",
        "$top": 100,
    }
    while url:
        resp = requests.get(url, headers=_headers(), params=params)
        _raise_for_status(resp, "get_inbox_conversation_ids")
        data = resp.json()
        for msg in data.get("value", []):
            subj    = msg.get("subject", "").lower()
            conv_id = msg.get("conversationId", "")
            if _is_bounce(subj) or not conv_id:
                continue
            conv_ids.add(conv_id)
        url = data.get("@odata.nextLink")
        params = {}
    return conv_ids


def _is_bounce(subject_lower: str) -> bool:
    bounce_markers = [
        "undeliverable", "delivery failed", "delivery failure",
        "mail delivery", "returned mail", "bounce", "non-delivery",
        "could not be delivered", "failure notice",
    ]
    return any(m in subject_lower for m in bounce_markers)
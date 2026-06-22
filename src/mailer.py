import configparser
import logging
import random
import time

import excel
import graph

CONFIG_PATH = "config.ini"


def _cfg():
    c = configparser.ConfigParser()
    c.read(CONFIG_PATH)
    return c


def _delay():
    cfg = _cfg()
    lo = int(cfg["settings"]["min_delay_seconds"])
    hi = int(cfg["settings"]["max_delay_seconds"])
    secs = random.randint(lo, hi)
    logging.info(f"Waiting {secs}s before next send...")
    time.sleep(secs)


# ── 1. Send new emails ────────────────────────────────────────────────────────

def send_new_emails():
    contacts = excel.get_pending_contacts()
    if not contacts:
        print("No pending contacts found.")
        logging.info("send_new_emails: no pending contacts")
        return

    print(f"Found {len(contacts)} pending contact(s). Starting sends...\n")

    for i, contact in enumerate(contacts):
        name    = contact["contact_name"]
        email   = contact["email"]
        subject = contact["subject"]
        body    = contact["initial_body"]
        row     = contact["row"]

        if not email or not subject or not body:
            logging.warning(f"Row {row} ({name}) missing email/subject/body — skipping")
            print(f"  ⚠  Skipping {name} — missing email, subject, or body")
            continue

        try:
            print(f"  → Sending to {name} <{email}>...")
            msg_id, conv_id = graph.send_email(to=email, subject=subject, body=body)
            excel.mark_sent(row=row, message_id=msg_id, conversation_id=conv_id)
            logging.info(f"SENT | {name} | {email} | row {row} | msg_id {msg_id}")
            print(f"     ✓ Sent. Message ID stored.")
        except Exception as e:
            logging.error(f"SEND_FAILED | {name} | {email} | {e}")
            print(f"     ✗ Failed: {e}")

        if i < len(contacts) - 1:
            _delay()

    print("\nDone sending.")


# ── 2. Schedule sends ─────────────────────────────────────────────────────────

def schedule_sends(start_dt_str: str):
    from datetime import datetime, timedelta, timezone
    import datetime as dt

    cfg = _cfg()
    min_gap = int(cfg["settings"]["schedule_min_gap_seconds"])
    max_gap = int(cfg["settings"]["schedule_max_gap_seconds"])

    try:
        local_dt = datetime.strptime(start_dt_str.strip(), "%Y-%m-%d %H:%M")
    except ValueError:
        print("Invalid format. Use YYYY-MM-DD HH:MM (e.g. 2026-05-20 09:00)")
        return

    local_dt_aware = local_dt.replace(
        tzinfo=dt.datetime.now(timezone.utc).astimezone().tzinfo
    )
    utc_dt = local_dt_aware.astimezone(timezone.utc).replace(tzinfo=None)

    contacts = excel.get_pending_contacts()
    if not contacts:
        print("No pending contacts found.")
        return

    print(f"\nScheduling {len(contacts)} email(s) starting at {start_dt_str}...\n")

    current_utc = utc_dt
    scheduled = []

    for i, contact in enumerate(contacts):
        name    = contact["contact_name"]
        email   = contact["email"]
        subject = contact["subject"]
        body    = contact["initial_body"]
        row     = contact["row"]

        if not email or not subject or not body:
            print(f"  ⚠  Skipping {name} — missing email, subject, or body")
            continue

        local_display = (
            current_utc.replace(tzinfo=timezone.utc)
            .astimezone()
            .strftime("%Y-%m-%d %I:%M %p")
        )
        send_at = current_utc.strftime("%Y-%m-%dT%H:%M:%S")

        try:
            msg_id, conv_id = graph.schedule_email(
                to=email, subject=subject, body=body, send_at=send_at
            )
            excel.mark_sent(row=row, message_id=msg_id, conversation_id=conv_id)
            logging.info(f"SCHEDULED | {name} | {email} | {local_display}")
            print(f"  ✓ {name} <{email}> → {local_display}")
            scheduled.append(local_display)
        except Exception as e:
            logging.error(f"SCHEDULE_FAILED | {name} | {email} | {e}")
            print(f"  ✗ Failed for {name}: {e}")

        gap = random.randint(min_gap, max_gap)
        current_utc += timedelta(seconds=gap)

    if scheduled:
        print(f"\n✓ {len(scheduled)} email(s) scheduled.")
        print(f"  First : {scheduled[0]}")
        print(f"  Last  : {scheduled[-1]}")


# ── 3. Reply detection ────────────────────────────────────────────────────────

def detect_replies() -> int:
    wb_data = excel._load()
    ws = wb_data[excel.ACTIVE_SHEET]

    active = []
    for row in range(2, ws.max_row + 1):
        conv_id = excel._cell_value(ws, row, excel.COL["conversation_id"])
        status  = str(excel._cell_value(ws, row, excel.COL["status"]) or "").strip().lower()
        email   = excel._cell_value(ws, row, excel.COL["email"])
        name    = excel._cell_value(ws, row, excel.COL["contact_name"])
        if conv_id and status not in ("", "pending"):
            active.append({"row": row, "conv_id": conv_id, "email": email, "name": name})

    if not active:
        logging.info("detect_replies: no active contacts with conversation IDs")
        return 0

    inbox_conv_ids = graph.get_inbox_conversation_ids(since_days=30)
    replied_count = 0

    for contact in sorted(active, key=lambda x: x["row"], reverse=True):
        if contact["conv_id"] not in inbox_conv_ids:
            continue
        try:
            excel.move_to_replied(row=contact["row"])
            logging.info(f"REPLY_DETECTED | {contact['email']} | row {contact['row']} → Replied tab")
            print(f"  ↩  Reply detected from {contact['email']} — moved to Replied tab")
            replied_count += 1
        except Exception as e:
            logging.error(f"MOVE_TO_REPLIED_FAILED | {contact['email']} | {e}")
            print(f"  ✗ Could not move {contact['email']} to Replied: {e}")

    return replied_count


# ── 4. Send follow-ups ────────────────────────────────────────────────────────

def send_followups():
    cfg = _cfg()
    max_followups = int(cfg["settings"]["max_follow_ups"])

    candidates = excel.get_followup_candidates()
    if not candidates:
        print("No follow-ups due.")
        logging.info("send_followups: none due")
        return

    print(f"Found {len(candidates)} follow-up(s) due.\n")

    for contact in candidates:
        name    = contact["contact_name"]
        email   = contact["email"]
        row     = contact["row"]
        count   = contact["followup_count"]
        msg_id  = contact["message_id"]
        conv_id = contact.get("conversation_id") or ""
        body    = contact["followup_body"]

        if not conv_id and not msg_id:
            logging.warning(f"Row {row} ({name}) has no message_id or conversation_id — skipping")
            print(f"  ⚠  Skipping {name} — run Option 7 to backfill conversation IDs first")
            continue

        if not body:
            logging.warning(f"Row {row} ({name}) has no follow-up body — skipping")
            print(f"  ⚠  Skipping {name} — follow-up body is empty")
            continue

        try:
            print(f"  → Follow-up #{count + 1} to {name} <{email}>...")
            graph.reply_to_message(
                message_id=msg_id,
                body=body,
                to_email=email,
                conversation_id=conv_id,
            )
            excel.mark_followup_sent(row=row)
            logging.info(f"FOLLOWUP_SENT | {name} | {email} | count={count + 1} | row {row}")
            print(f"     ✓ Follow-up #{count + 1} sent.")

            if count + 1 >= max_followups:
                print(f"     → Max follow-ups reached for {name}. Archiving...")
                excel.move_to_archived(row=row)
                logging.info(f"ARCHIVED | {name} | {email} | row {row}")
                print(f"     ✓ Moved to Archived tab.")

        except Exception as e:
            logging.error(f"FOLLOWUP_FAILED | {name} | {email} | {e}")
            print(f"     ✗ Failed: {e}")

    print("\nDone with follow-ups.")


# ── 5. Daily check ────────────────────────────────────────────────────────────

def run_daily_check():
    print("── Step 1: Scanning inbox for replies ──")
    n = detect_replies()
    print(f"   {n} reply(ies) detected.\n")
    print("── Step 2: Sending due follow-ups ──")
    send_followups()


# ── 6. Fix email address ──────────────────────────────────────────────────────

def fix_email_for_contact(search_term: str):
    matches = excel.find_contact(search_term)
    if not matches:
        print(f"No contacts found matching '{search_term}'.")
        return

    print(f"\nFound {len(matches)} match(es):\n")
    for i, c in enumerate(matches):
        print(f"  [{i}] {c['company']} | {c['contact_name']} | {c['email']} | Status: {c['status']}")

    choice = input("\nEnter index to select (or 'q' to cancel): ").strip()
    if choice.lower() == "q":
        return

    try:
        selected = matches[int(choice)]
    except (ValueError, IndexError):
        print("Invalid selection.")
        return

    print(f"\nCurrent email: {selected['email']}")
    new_email = input("Enter correct email address: ").strip()
    if not new_email:
        print("No email entered. Cancelled.")
        return

    excel.fix_email(row=selected["row"], new_email=new_email)
    logging.info(
        f"EMAIL_FIXED | {selected['contact_name']} | {selected['email']} → {new_email} | row {selected['row']}"
    )
    print(f"✓ Email updated to '{new_email}' and row reset to Pending.")


# ── 7. Backfill missing conversation IDs ─────────────────────────────────────

def backfill_conversation_ids():
    rows = excel.get_empty_conversation_ids()
    if not rows:
        print("All contacts already have conversation IDs.")
        return

    print(f"Found {len(rows)} contact(s) with missing conversation ID.\n")
    filled = 0
    failed = 0

    for contact in rows:
        name    = contact["contact_name"]
        email   = contact["email"]
        subject = contact["subject"]
        row     = contact["row"]

        conv_id = graph.find_conversation_id(to=email, subject=subject)
        if conv_id:
            excel.backfill_conversation_id(row=row, conv_id=conv_id)
            logging.info(f"BACKFILL | {name} | {email} | conv_id={conv_id}")
            print(f"  ✓ {name} <{email}> — backfilled")
            filled += 1
        else:
            logging.warning(f"BACKFILL_FAILED | {name} | {email} — not found in Sent Items")
            print(f"  ✗ {name} <{email}> — not found in Sent Items (may not have sent yet)")
            failed += 1

    print(f"\n✓ Backfilled: {filled}  |  Not found: {failed}")
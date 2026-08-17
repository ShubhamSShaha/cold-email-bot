import logging
import random
import time
from datetime import datetime, timedelta

import config
import excel
import gmail


def _delay():
    lo = config.get_int("settings", "min_delay_seconds", 30)
    hi = config.get_int("settings", "max_delay_seconds", 60)
    secs = random.randint(min(lo, hi), max(lo, hi))
    logging.info(f"Waiting {secs}s before next send...")
    time.sleep(secs)


# ── Connection check ──────────────────────────────────────────────────────────

def check_connection() -> bool:
    """Verify login and mailbox access before sending anything."""
    try:
        profile = gmail.get_profile()
    except Exception as e:
        print(f"  ✗ {e}")
        return False

    configured = config.get("email", "sender")

    if not profile["ok"]:
        print(f"  ✗ {profile['detail']}")
        print("\n  Could not reach Gmail. Common causes:")
        print("    • Consent wasn't completed with all requested scopes granted —")
        print("      run option 9 (sign out), then retry to re-authenticate from scratch.")
        print("    • A Testing-mode Google Cloud app's refresh token expired (7-day limit) —")
        print("      option 9 then retry will prompt a fresh device-code login.")
        return False

    print(f"  Signed in as   : {profile['email']}")
    print(f"  Mailbox        : ✓ reachable ({profile['messages_total']} messages)")

    if configured and profile["email"] and configured.lower() != profile["email"].lower():
        print(
            f"  ⚠  config.ini email.sender is {configured} but you are signed in as "
            f"{profile['email']} — mail will be sent from the signed-in account. "
            "If this account has Gmail 'Send As' aliases configured, double-check "
            "which address you expect to appear in outgoing mail."
        )

    return True


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
            msg_id, thread_id = gmail.send_email(to=email, subject=subject, body=body)
            excel.mark_sent(row=row, message_id=msg_id, conversation_id=thread_id)
            logging.info(f"SENT | {name} | {email} | row {row} | msg_id {msg_id}")
            print(f"     ✓ Sent.")
        except Exception as e:
            logging.error(f"SEND_FAILED | {name} | {email} | {e}")
            print(f"     ✗ Failed: {e}")

        if i < len(contacts) - 1:
            _delay()

    print("\nDone sending.")


# ── 2. Schedule sends (local scheduler — Gmail has no server-side delay) ──────

def schedule_sends(start_dt_str: str):
    min_gap = config.get_int("settings", "schedule_min_gap_seconds", 60)
    max_gap = config.get_int("settings", "schedule_max_gap_seconds", 180)

    try:
        local_dt = datetime.strptime(start_dt_str.strip(), "%Y-%m-%d %H:%M")
    except ValueError:
        print("Invalid format. Use YYYY-MM-DD HH:MM (e.g. 2026-05-20 09:00)")
        return

    local_tz = datetime.now().astimezone().tzinfo
    target = local_dt.replace(tzinfo=local_tz)

    contacts = excel.get_pending_contacts()
    if not contacts:
        print("No pending contacts found.")
        return

    plan = []
    current = target
    for contact in contacts:
        plan.append((contact, current))
        current += timedelta(seconds=random.randint(min_gap, max_gap))

    print(f"\nScheduling {len(plan)} email(s) starting {start_dt_str}.")
    print(f"Estimated finish: {plan[-1][1].strftime('%Y-%m-%d %I:%M %p')}")
    print("Gmail has no server-side delayed delivery, so this script sends each email")
    print("itself when its time comes — keep this terminal open and the machine awake")
    print("for the whole window (on macOS: caffeinate -s python3 src/main.py).\n")

    sent_count = 0
    for contact, target_dt in plan:
        name    = contact["contact_name"]
        email   = contact["email"]
        subject = contact["subject"]
        body    = contact["initial_body"]
        row     = contact["row"]

        if not email or not subject or not body:
            print(f"  ⚠  Skipping {name} — missing email, subject, or body")
            continue

        wait = (target_dt - datetime.now().astimezone()).total_seconds()
        if wait > 0:
            print(f"  ⏳ Waiting until {target_dt.strftime('%Y-%m-%d %I:%M %p')} for {name}...")
            time.sleep(wait)

        try:
            print(f"  → Sending to {name} <{email}>...")
            msg_id, thread_id = gmail.send_email(to=email, subject=subject, body=body)
            excel.mark_sent(
                row=row, message_id=msg_id, conversation_id=thread_id,
                sent_date=datetime.now().date(),
            )
            logging.info(f"SCHEDULED_SENT | {name} | {email} | row {row}")
            print("     ✓ Sent.")
            sent_count += 1
        except Exception as e:
            logging.error(f"SCHEDULE_SEND_FAILED | {name} | {email} | {e}")
            print(f"     ✗ Failed: {e}")

    print(f"\n✓ {sent_count} email(s) sent.")


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

    inbox_thread_ids = gmail.get_inbox_thread_ids(since_days=30)
    replied_count = 0

    for contact in sorted(active, key=lambda x: x["row"], reverse=True):
        if contact["conv_id"] not in inbox_thread_ids:
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
    max_followups = config.get_int("settings", "max_follow_ups", 5)

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
        conv_id = contact.get("conversation_id") or ""
        subject = contact["subject"]
        body    = contact["followup_body"]

        if not conv_id:
            logging.warning(f"Row {row} ({name}) has no conversation_id — skipping")
            print(f"  ⚠  Skipping {name} — run Option 7 to backfill conversation IDs first")
            continue

        if not body:
            logging.warning(f"Row {row} ({name}) has no follow-up body — skipping")
            print(f"  ⚠  Skipping {name} — follow-up body is empty")
            continue

        try:
            print(f"  → Follow-up #{count + 1} to {name} <{email}>...")
            gmail.reply_to_message(
                thread_id=conv_id,
                body=body,
                to_email=email,
                subject=subject,
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

        thread_id = gmail.find_thread_id(to=email, subject=subject)
        if thread_id:
            excel.backfill_conversation_id(row=row, conv_id=thread_id)
            logging.info(f"BACKFILL | {name} | {email} | thread_id={thread_id}")
            print(f"  ✓ {name} <{email}> — backfilled")
            filled += 1
        else:
            logging.warning(f"BACKFILL_FAILED | {name} | {email} — not found in Sent")
            print(f"  ✗ {name} <{email}> — not found in Sent (may not have sent yet)")
            failed += 1

    print(f"\n✓ Backfilled: {filled}  |  Not found: {failed}")

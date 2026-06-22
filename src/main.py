import configparser
import logging
import os
import sys

sys.path.insert(0, os.path.dirname(__file__))

import mailer
import excel

CONFIG_PATH = "config.ini"


def _setup_logging():
    cfg = configparser.ConfigParser()
    cfg.read(CONFIG_PATH)
    log_path = cfg["settings"]["log_path"]
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )


def _print_menu():
    print("\n" + "=" * 50)
    print("  Cold Email Bot")
    print("=" * 50)
    print("  1.  Send new emails       (Pending contacts)")
    print("  1b. Schedule sends        (Pick start date/time)")
    print("  2.  Run daily check       (Replies + Follow-ups)")
    print("  3.  Status summary")
    print("  4.  Fix email for a contact")
    print("  5.  Import contacts from file")
    print("  7.  Backfill conversation IDs")
    print("  8.  Exit")
    print("=" * 50)


def _show_summary():
    s = excel.get_status_summary()
    print("\n── Outreach Summary ──────────────────────────")
    print(f"  Active contacts  : {s['active']}")
    print(f"    └ Pending      : {s['pending']}")
    print(f"    └ Sent         : {s['sent']}")
    print(f"    └ Follow-up    : {s['followup_sent']}")
    print(f"  Replied          : {s['replied']}")
    print(f"  Archived         : {s['archived']}")
    print(f"  Follow-ups due   : {s['due_today']}")
    print("─" * 46)


def _import_contacts():
    print("\n── Import Contacts from File ─────────────────")
    cfg = configparser.ConfigParser()
    cfg.read(CONFIG_PATH)
    path = cfg["settings"]["import_path"]
    if not os.path.exists(path):
        print(f"File not found: {path}")
        print(f"Make sure '{path}' is in the cold-email-bot/ root folder.")
        return
    result = excel.import_from_file(path)
    print(f"\n  ✓ Imported : {result['imported']} contact(s)")
    print(f"  ⚠ Skipped  : {result['skipped']} contact(s)")
    for r in result["reasons"]:
        print(f"    └ {r}")
    logging.info(
        f"IMPORT | imported={result['imported']} skipped={result['skipped']} | {path}"
    )


def main():
    if not os.path.exists(CONFIG_PATH):
        print(f"ERROR: config.ini not found at {CONFIG_PATH}")
        print("Make sure you are running this from the cold-email-bot/ root directory.")
        sys.exit(1)

    _setup_logging()
    logging.info("Cold Email Bot started")

    while True:
        _print_menu()
        choice = input("Select option: ").strip()

        if choice == "1":
            print("\n── Send New Emails ───────────────────────────")
            mailer.send_new_emails()

        elif choice == "1b":
            print("\n── Schedule Sends ────────────────────────────")
            start = input("Enter start date and time (YYYY-MM-DD HH:MM): ").strip()
            mailer.schedule_sends(start)

        elif choice == "2":
            print("\n── Daily Check ───────────────────────────────")
            mailer.run_daily_check()

        elif choice == "3":
            _show_summary()

        elif choice == "4":
            print("\n── Fix Email Address ─────────────────────────")
            search = input("Enter company or contact name to search: ").strip()
            if search:
                mailer.fix_email_for_contact(search)

        elif choice == "5":
            _import_contacts()

        elif choice == "7":
            print("\n── Backfill Conversation IDs ─────────────────")
            mailer.backfill_conversation_ids()

        elif choice == "8":
            logging.info("Cold Email Bot exited")
            print("Goodbye.")
            break

        else:
            print("Invalid option. Enter 1, 1b, 2, 3, 4, 5, 7, or 8.")


if __name__ == "__main__":
    main()
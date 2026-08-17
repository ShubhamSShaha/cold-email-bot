import logging
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import auth
import config
import excel
import mailer


def _setup_logging():
    log_path = config.path("log_path", "outreach.log")
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
    print("  0.  Check connection      (Verify login + mailbox)")
    print("  1.  Send new emails       (Pending contacts)")
    print("  1b. Schedule sends        (Pick start date/time)")
    print("  2.  Run daily check       (Replies + Follow-ups)")
    print("  3.  Status summary")
    print("  4.  Fix email for a contact")
    print("  5.  Import contacts from file")
    print("  7.  Backfill conversation IDs")
    print("  8.  Exit")
    print("  9.  Sign out              (Clear cached login)")
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
    path = config.path("import_path", "latestList.xlsx")
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
    if not config.exists():
        print(f"ERROR: config.ini not found at {config.CONFIG_PATH}")
        print("Copy config.ini.template to config.ini and fill in your values.")
        sys.exit(1)

    _setup_logging()
    logging.info("Cold Email Bot started")

    while True:
        _print_menu()
        choice = input("Select option: ").strip().lower()

        if choice == "0":
            print("\n── Connection Check ──────────────────────────")
            mailer.check_connection()

        elif choice == "1":
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

        elif choice == "9":
            auth.sign_out()
            logging.info("Signed out — token cache cleared")
            print("Cached login cleared. The next action will prompt for a device-code login.")

        else:
            print("Invalid option. Enter 0, 1, 1b, 2, 3, 4, 5, 7, 8, or 9.")


if __name__ == "__main__":
    main()
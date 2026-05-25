#!/usr/bin/env python3
"""
Timetable Bot - Test Mode
Creates a throwaway test spreadsheet, starts the bot polling (finding nothing),
then releases a fake timetable sheet after a short delay — simulating 7pm release.
Shows exactly how fast the bot detects and books the slot.
"""

import os
import sys
import time
import threading
from datetime import datetime, timedelta

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import timetable_bot as bot
from timetable_bot import (
    get_service, list_sheets, find_target_sheet,
    attempt_booking, get_date, get_ward, MY_NAME,
)


# ─── Test spreadsheet helpers ─────────────────────────────────────────────────

def create_test_spreadsheet(service):
    """Create a blank test spreadsheet with one placeholder sheet."""
    body = {
        "properties": {"title": "TIMETABLE BOT TEST (safe to delete)"},
        "sheets": [{"properties": {"title": "PLACEHOLDER"}}],
    }
    result = service.spreadsheets().create(body=body).execute()
    return result["spreadsheetId"], result["spreadsheetUrl"]


def build_timetable_data(anchor_date):
    """
    Build a list-of-rows matching the real timetable structure:
    - Night Shift date header near the top
    - Empty WARD 8PA / 8PB / 7PA rows further down (where the bot writes)
    """
    dates = [anchor_date + timedelta(days=i) for i in range(7)]
    date_cells = [f"{d.day}/{d.month}/{str(d.year)[2:]}" for d in dates]

    rows = []
    # Row 1 — top info (ignored by bot)
    rows.append(["", "TEST TIMETABLE"])
    rows.append(["", "AM shift: 7AM-5PM"])
    rows.append(["", "PM shift: 7AM-8PM", "MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"])
    # Row 4 — Night Shift header with dates (this is what the bot uses to find columns)
    rows.append(["", "Night Shift: 8PM-9AM"] + date_cells)
    rows.append([])

    # Filler rows to simulate the real sheet's depth
    for _ in range(43):
        rows.append([])

    # Night Shift REQUEST section (row ~49 onward)
    rows.append(["", "NIGHT SHIFT"])          # triggers bot's night-shift search
    rows.append(["", "WARD 8PA"] + [""] * 7)  # all empty — bot fills here
    rows.append(["", "WARD 8PB"] + [""] * 7)
    rows.append(["", "WARD 7PA"] + [""] * 7)

    return rows


def add_timetable_sheet(service, spreadsheet_id, anchor_date):
    """Add the timetable sheet tab to the test spreadsheet."""
    start = anchor_date
    end   = anchor_date + timedelta(days=6)
    sheet_title = f"{start.day}/{start.month} - {end.day}/{end.month}"

    # Create the new tab
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": [{"addSheet": {"properties": {"title": sheet_title}}}]},
    ).execute()

    # Populate it
    data = build_timetable_data(anchor_date)
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"'{sheet_title}'!A1",
        valueInputOption="USER_ENTERED",
        body={"values": data},
    ).execute()

    return sheet_title


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    print("""
+--------------------------------------------------------------+
|        TIMETABLE BOT - TEST MODE                             |
|   Simulates the 7pm timetable release in real time          |
+--------------------------------------------------------------+
""")

    print("Connecting to Google...")
    service = get_service()
    print("   Connected!\n")

    # Create throwaway spreadsheet
    print("Creating temporary test spreadsheet...")
    test_id, test_url = create_test_spreadsheet(service)
    print(f"   Created: {test_url}\n")

    # Preferences
    print("--- FIRST CHOICE ---")
    date1 = get_date("  Date to book (DD/MM or DD/MM/YYYY): ")
    ward1 = get_ward("  Ward (8PA / 8PB / 7PA):              ")
    print("\n--- SECOND CHOICE (fallback) ---")
    date2 = get_date("  Fallback date:  ")
    ward2 = get_ward("  Fallback ward:  ")

    print(f"""
  Booking:   {MY_NAME}
  1st choice: WARD {ward1}  on  {date1.strftime('%d/%m/%Y')}
  2nd choice: WARD {ward2}  on  {date2.strftime('%d/%m/%Y')}
""")

    # Override the spreadsheet ID so the bot targets our test sheet
    bot.SPREADSHEET_ID = test_id

    # ── Release time (mirrors real bot behaviour) ─────────────────────────────
    print("What time should the test sheet 'release'? (HH:MM, 24-hour)")
    print("   Enter a time 1-2 minutes from now to test the full countdown.")
    print("   Press Enter to release in 5 seconds instead.")
    raw_time = input("\n   Release time: ").strip()

    release_at = None
    if raw_time:
        try:
            h, m = map(int, raw_time.replace(".", ":").split(":"))
            now  = datetime.now()
            release_at = now.replace(hour=h, minute=m, second=0, microsecond=0)
            if release_at <= now:
                print("   That time has already passed — releasing in 5 seconds instead.")
                release_at = None
        except Exception:
            print("   Could not read that — releasing in 5 seconds instead.")

    slow_start   = (release_at - timedelta(minutes=5)) if release_at else None
    release_time = None   # set by background thread when sheet actually drops

    # ── Wait until 5 min before release ──────────────────────────────────────
    if slow_start and slow_start > datetime.now():
        print(f"\n   Waiting until {slow_start.strftime('%H:%M')} "
              f"(5 min before release) ...  Ctrl+C to cancel\n")
        try:
            while datetime.now() < slow_start:
                left       = (slow_start - datetime.now()).total_seconds()
                h2, rem    = divmod(int(left), 3600)
                m2, s2     = divmod(rem, 60)
                label      = f"{h2}h {m2:02d}m {s2:02d}s" if h2 else f"{m2:02d}m {s2:02d}s"
                sys.stdout.write(f"\r   {label} until warm-up ...   ")
                sys.stdout.flush()
                time.sleep(1)
            print("\r   Warm-up started — polling every 10s.              ")
        except KeyboardInterrupt:
            print("\n   Cancelled.")
            sys.exit(0)

    def release_sheet():
        nonlocal release_time
        if release_at:
            # Wait until the exact release time
            while datetime.now() < release_at:
                time.sleep(0.5)
        else:
            # No time given — count down 5 seconds from now
            for i in range(5, 0, -1):
                sys.stdout.write(f"\r  >> Sheet releasing in {i}s ...   ")
                sys.stdout.flush()
                time.sleep(1)

        print("\r  >> SHEET RELEASED — adding to spreadsheet now!   ")
        release_time = datetime.now()
        try:
            thread_service = get_service()
            title = add_timetable_sheet(thread_service, test_id, date1)
            print(f"  >> Sheet '{title}' is live.\n")
        except Exception as exc:
            print(f"  >> Error releasing sheet: {exc}")

    threading.Thread(target=release_sheet, daemon=True).start()

    # ── Polling loop ──────────────────────────────────────────────────────────
    if release_at:
        rapid_at = release_at - timedelta(minutes=2)
        print(f"  BOT POLLING — slow (10s) until {rapid_at.strftime('%H:%M')}, rapid (2s) from then on.\n")
    else:
        print("  BOT POLLING every 2 seconds...\n")

    known_titles: set = set()
    try:
        for s in list_sheets(service):
            known_titles.add(s["properties"]["title"])
    except Exception:
        pass

    attempt = 0
    done    = False
    loop_start = datetime.now()

    while not done:
        try:
            attempt += 1
            ts = datetime.now().strftime("%H:%M:%S.%f")[:-3]

            sheets = list_sheets(service)
            titles = {s["properties"]["title"] for s in sheets}

            rapid_at = (release_at - timedelta(minutes=2)) if release_at else None
            interval = 2 if (rapid_at is None or datetime.now() >= rapid_at) else 10

            new = titles - known_titles
            if new:
                detected_at = datetime.now()
                lag = (detected_at - release_time).total_seconds() if release_time else "?"
                print(f"  [{ts}]  NEW SHEET DETECTED: {new}  (detected {lag:.2f}s after release)")
                known_titles = titles
                interval     = 2   # snap to rapid the moment sheet appears

            target = find_target_sheet(sheets)
            sheet_name = target["properties"]["title"]

            if sheet_name == "PLACEHOLDER":
                mode = "RAPID" if interval == 2 else "slow "
                sys.stdout.write(f"\r  [{ts}]  #{attempt:>3}  [{mode}]  Waiting for timetable to drop ...")
                sys.stdout.flush()
                time.sleep(interval)
                continue

            sys.stdout.write(f"\r  [{ts}]  #{attempt:>3}  Sheet found — attempting booking ...   ")
            sys.stdout.flush()

            status, detail = attempt_booking(service, sheet_name, ward1, date1)

            if status == "success":
                booked_at = datetime.now()
                total = (booked_at - release_time).total_seconds() if release_time else "?"
                print(f"\n\n  SUCCESS!  WARD {ward1}  on  {date1.strftime('%d/%m/%Y')}")
                print(f"  Cell written: {detail}")
                print(f"  Time from release to booked: {total:.2f} seconds")
                done = True

            elif status == "taken":
                print(f"\n  1st choice taken — trying fallback ...")
                status2, detail2 = attempt_booking(service, sheet_name, ward2, date2)
                if status2 == "success":
                    total = (datetime.now() - release_time).total_seconds() if release_time else "?"
                    print(f"  SUCCESS (fallback)!  WARD {ward2}  ->  {detail2}  ({total:.2f}s)")
                    done = True
                elif status2 == "taken":
                    print("  Both slots taken in test — check your dates.")
                    done = True

            elif status == "error":
                print(f"\n  Error: {detail}")

            if not done:
                time.sleep(1)

        except KeyboardInterrupt:
            print("\n  Stopped.")
            done = True
        except Exception as exc:
            print(f"\n  Unexpected error: {exc}")
            time.sleep(2)

    print(f"\n  Test complete!")
    print(f"  View result: {test_url}")
    print("  Delete the test spreadsheet from Google Drive when done.\n")
    input("Press Enter to exit...")


if __name__ == "__main__":
    main()

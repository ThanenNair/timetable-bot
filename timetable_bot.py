#!/usr/bin/env python3
"""
Night Shift Timetable Bot
Automatically fills in 'THANEN' in the weekly night shift request timetable
the moment the new sheet becomes available.
"""

import os
import sys
import time
import pickle
import re
from datetime import datetime, timedelta

# Force UTF-8 output on Windows so box-drawing chars and arrows print correctly
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# ─────────────────────────────────────────────────────────────────────────────
SPREADSHEET_ID = "1sGtcCSvpKwK8ONgV0uy9ZXl_Cwp8C-YbaXx-fNfbbGY"
MY_NAME        = "THANEN"
SCOPES         = ["https://www.googleapis.com/auth/spreadsheets"]
WARDS          = ["8PA", "8PB", "7PA"]
TOKEN_FILE     = os.path.join(os.path.dirname(os.path.abspath(__file__)), "token.pickle")
CREDS_FILE     = os.path.join(os.path.dirname(os.path.abspath(__file__)), "credentials.json")
# ─────────────────────────────────────────────────────────────────────────────


def ensure_packages():
    try:
        import googleapiclient  # noqa: F401
        import google_auth_oauthlib  # noqa: F401
    except ImportError:
        print("📦  Installing required packages (first time only)...")
        os.system(f'"{sys.executable}" -m pip install --quiet '
                  'google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client')
        print("✅  Packages installed!\n")


def setup_instructions():
    print("""
╔══════════════════════════════════════════════════════════════╗
║          FIRST-TIME SETUP: Google API Credentials           ║
╚══════════════════════════════════════════════════════════════╝

You need a  credentials.json  file in the same folder as this bot.
This is a one-time setup that takes about 5 minutes.

STEP 1 ─ Open Google Cloud Console
  → https://console.cloud.google.com/

STEP 2 ─ Create a project
  • Click the project dropdown at the top → "New Project"
  • Name it anything (e.g. TimetableBot) → Create

STEP 3 ─ Enable Google Sheets API
  • Left menu: "APIs & Services" → "Enable APIs and Services"
  • Search "Google Sheets API" → Click it → Enable

STEP 4 ─ Create credentials
  • Left menu: "APIs & Services" → "Credentials"
  • Click "+ CREATE CREDENTIALS" → "OAuth client ID"
  • Application type: "Desktop app"
  • Name: anything → Create
  • Click the download icon (⬇) next to your new credential
  • Save the file as  credentials.json

STEP 5 ─ Place the file
  • Move credentials.json to:
    C:\\Users\\PC\\TimetableBot\\credentials.json

STEP 6 ─ Run the bot again
  • A browser window will open — log in with your hospital Google account
  • That's it! The bot will remember your login for future runs.
""")


def get_service():
    ensure_packages()
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build

    creds = None
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "rb") as f:
            creds = pickle.load(f)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDS_FILE):
                setup_instructions()
                sys.exit(1)
            flow = InstalledAppFlow.from_client_secrets_file(CREDS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "wb") as f:
            pickle.dump(creds, f)

    return build("sheets", "v4", credentials=creds)


# ─── Sheet helpers ────────────────────────────────────────────────────────────

def list_sheets(service):
    info = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    return info.get("sheets", [])


def find_target_sheet(sheets):
    """
    Pick the most current timetable sheet by parsing dates from sheet names.
    Sheet names are expected to contain dates like '1/6', '1/6/26', '01/06' etc.
    Returns the sheet with the start date closest to today (including future sheets
    up to 60 days ahead, to catch newly released upcoming weeks).
    Falls back to the highest-index sheet if no dates can be parsed.
    """
    today = datetime.now()
    candidates = []

    for sheet in sheets:
        title = sheet["properties"]["title"]
        match = re.search(r'(\d{1,2})[/\-](\d{1,2})(?:[/\-](\d{2,4}))?', title)
        if match:
            try:
                day   = int(match.group(1))
                month = int(match.group(2))
                yr_raw = match.group(3)
                yr = today.year
                if yr_raw:
                    yr = int(yr_raw)
                    if yr < 100:
                        yr += 2000
                sheet_date = datetime(yr, month, day)
                candidates.append((sheet_date, sheet))
            except ValueError:
                pass

    if not candidates:
        return sorted(sheets, key=lambda s: s["properties"]["index"])[-1]

    # Sort newest-first
    candidates.sort(key=lambda x: x[0], reverse=True)

    # Prefer the most recent sheet that started no more than 60 days in the future
    future_cutoff = today + timedelta(days=60)
    valid = [(d, s) for d, s in candidates if d <= future_cutoff]

    if valid:
        return valid[0][1]   # most recent within window

    return candidates[-1][1]  # everything is in the future; pick earliest


def read_sheet(service, sheet_name):
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!A1:AZ200",
    ).execute()
    return result.get("values", [])


def col_letter(idx):
    """0-based column index → A1 letter (supports up to ZZ)."""
    result = ""
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        result = chr(65 + rem) + result
    return result


# ─── Timetable parsing ────────────────────────────────────────────────────────

def find_ward_row(data, ward_code):
    """
    Flexibly locate the booking row for a given ward (e.g. 8PA).

    Strategy:
    1. Collect every row index where the ward code appears as a standalone
       token inside the first two columns (handles "WARD 8PA", "8PA",
       "8PA (2-3)", "Ward 8PA", etc.).
    2. Collect every row that looks like a Night-Shift section header
       (contains "night" anywhere — case-insensitive).
    3. Prefer ward-candidate rows that appear AFTER the last night-shift
       header; if none, prefer the lower half of the sheet; final fallback
       is the last occurrence.
    4. The header_row is the last night-shift row at or before the ward row
       (used as the starting point for date-column search).

    Returns (header_row_idx, ward_row_idx), either of which may be None.
    """
    ward_key = ward_code.upper()  # e.g. "8PA"
    # Regex: ward code as a standalone token (not glued to other letters/digits)
    ward_pattern = re.compile(
        r'(?<![A-Z\d])' + re.escape(ward_key) + r'(?![A-Z\d])',
        re.IGNORECASE,
    )

    night_rows:  list[int] = []
    ward_candidates: list[int] = []

    for i, row in enumerate(data):
        if not row:
            continue

        # Night-shift header detection — any row containing "night"
        full_text = " ".join(str(c) for c in row)
        if re.search(r'night', full_text, re.IGNORECASE):
            night_rows.append(i)

        # Ward detection — only look in first two columns to avoid matching
        # individual shift-assignment cells deeper in the row
        first_two = " ".join(str(row[j]) for j in range(min(2, len(row))))
        if ward_pattern.search(first_two):
            ward_candidates.append(i)

    if not ward_candidates:
        return (night_rows[-1] if night_rows else None), None

    # Pick the best ward row
    last_night = night_rows[-1] if night_rows else -1

    # 1st preference: first ward row that comes after the last night-shift header
    after_night = [r for r in ward_candidates if r > last_night]
    if after_night:
        ward_row = after_night[0]
    else:
        # 2nd preference: ward row in the lower half of the sheet
        mid = len(data) // 2
        lower = [r for r in ward_candidates if r >= mid]
        ward_row = lower[0] if lower else ward_candidates[-1]

    # Header row = last night-shift row at or before the ward row
    header_row = None
    for nr in reversed(night_rows):
        if nr <= ward_row:
            header_row = nr
            break

    return header_row, ward_row


def date_matches_cell(cell_text, target):
    """
    Return True if cell_text looks like the given datetime date.
    Handles many common formats: 25/5, 25-5, 25 May, MON 25/5, 25, etc.
    """
    cell = str(cell_text).strip()
    if not cell:
        return False

    day, month = target.day, target.month

    # Specific patterns (day + month) — safe, no false positives
    patterns = [
        f"{day}/{month}",
        f"{day}-{month}",
        f"{day:02d}/{month:02d}",
        f"{day:02d}-{month:02d}",
        target.strftime("%d %b").lstrip("0"),
        target.strftime("%d %b"),
        target.strftime("%d %B"),
    ]
    for p in patterns:
        if p.lower() in cell.lower():
            return True

    # Last resort: cell is ONLY the day number (e.g. a header cell that just says "6")
    if cell.strip() == str(day):
        return True

    return False


def find_date_col(data, header_row_idx, target_date):
    """
    Find the column index for target_date.
    First searches near the header row, then falls back to scanning the whole
    sheet — needed because the date header row may be far above the ward rows.
    """
    # Pass 1: rows near the ward section header
    if header_row_idx is not None:
        search_rows = [header_row_idx + offset for offset in range(-3, 4)
                       if 0 <= header_row_idx + offset < len(data)]
        for ri in search_rows:
            for ci, cell in enumerate(data[ri]):
                if date_matches_cell(cell, target_date):
                    return ci

    # Pass 2: full-sheet scan (handles sheets where the date header is far away)
    for ri, row in enumerate(data):
        for ci, cell in enumerate(row):
            if date_matches_cell(cell, target_date):
                return ci

    return None


# ─── Booking ─────────────────────────────────────────────────────────────────

def attempt_booking(service, sheet_name, ward, date):
    """
    Try to write MY_NAME into the cell for (ward, date).
    Returns ("success", cell_ref) | ("taken", current_name)
             | ("not_found", reason) | ("error", msg)
    """
    try:
        data = read_sheet(service, sheet_name)
        header_row_idx, ward_row_idx = find_ward_row(data, ward)

        if ward_row_idx is None:
            return "not_found", f"Ward {ward} not in sheet '{sheet_name}'"

        col_idx = find_date_col(data, header_row_idx, date)
        if col_idx is None:
            return "not_found", (f"Date {date.strftime('%d/%m/%Y')} not found in sheet "
                                 f"(header row {header_row_idx})")

        # Check current occupant
        current = ""
        if ward_row_idx < len(data) and col_idx < len(data[ward_row_idx]):
            current = str(data[ward_row_idx][col_idx]).strip()
        if current and current.upper() not in ("", "N/A", "-"):
            return "taken", current

        # Write!
        cell_ref = f"'{sheet_name}'!{col_letter(col_idx)}{ward_row_idx + 1}"
        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=cell_ref,
            valueInputOption="USER_ENTERED",
            body={"values": [[MY_NAME]]},
        ).execute()
        return "success", cell_ref

    except Exception as exc:
        return "error", str(exc)


# ─── Debug / preview ─────────────────────────────────────────────────────────

def debug_sheet(service):
    """Print the detected structure of the current sheet so you can verify."""
    print("\n🔍  DEBUG MODE — reading current sheet structure...\n")
    sheets = list_sheets(service)
    if not sheets:
        print("  No sheets found.")
        return

    # Show every sheet in the spreadsheet
    print(f"  All sheets found ({len(sheets)} total):")
    for s in sorted(sheets, key=lambda x: x["properties"]["index"]):
        print(f"    [{s['properties']['index']}]  {s['properties']['title']}")

    target = find_target_sheet(sheets)
    name = target["properties"]["title"]
    print(f"\n  ✅  Bot will use sheet: '{name}'\n")

    data = read_sheet(service, name)
    print(f"  Total rows read: {len(data)}\n")
    print("  First 15 rows (truncated to 10 cols):")
    for i, row in enumerate(data[:15]):
        print(f"    Row {i+1:>2}: {row[:10]}")

    print()
    for ward in WARDS:
        hr, wr = find_ward_row(data, ward)
        print(f"  WARD {ward}  →  header_row={hr}  ward_row={wr}")
        if wr is not None and wr < len(data):
            print(f"    Content: {data[wr][:10]}")
    print()


# ─── CLI helpers ─────────────────────────────────────────────────────────────

def get_date(prompt):
    while True:
        raw = input(prompt).strip()
        for fmt in ("%d/%m/%Y", "%d/%m/%y", "%d/%m", "%d-%m-%Y", "%d-%m"):
            try:
                if fmt in ("%d/%m", "%d-%m"):
                    raw_with_year = raw + f"/{datetime.now().year}"
                    return datetime.strptime(raw_with_year, fmt + "/%Y")
                return datetime.strptime(raw, fmt)
            except ValueError:
                pass
        print("  ⚠️   Format not recognised. Try:  25/05  or  25/05/2026")


def get_ward(prompt):
    while True:
        raw = input(prompt).strip().upper().replace("WARD ", "").strip()
        if raw in WARDS:
            return raw
        print(f"  ⚠️   Please enter one of: {' / '.join(WARDS)}")


# ─── Main ────────────────────────────────────────────────────────────────────

BANNER = """
+--------------------------------------------------------------+
|        NIGHT SHIFT TIMETABLE BOT  v1.0                       |
|   Books your shift the instant the weekly form goes live     |
+--------------------------------------------------------------+
"""

def main():
    ensure_packages()

    if "--debug" in sys.argv:
        print(BANNER)
        print("🔐  Connecting to Google Sheets...")
        svc = get_service()
        print("   ✅  Connected!\n")
        debug_sheet(svc)
        input("Press Enter to exit...")
        return

    print(BANNER)

    # ── Auth ──────────────────────────────────────────────────────────────────
    print("🔐  Connecting to Google Sheets...")
    try:
        service = get_service()
        print("   ✅  Connected!\n")
    except SystemExit:
        raise
    except Exception as exc:
        print(f"   ❌  Failed: {exc}")
        sys.exit(1)

    # ── Preferences ───────────────────────────────────────────────────────────
    print("━━━  FIRST CHOICE  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    date1 = get_date("  Date (DD/MM or DD/MM/YYYY): ")
    ward1 = get_ward("  Ward (8PA / 8PB / 7PA):     ")

    print("\n━━━  SECOND CHOICE  (used if first slot is already taken)  ━━━━━━")
    date2 = get_date("  Date (DD/MM or DD/MM/YYYY): ")
    ward2 = get_ward("  Ward (8PA / 8PB / 7PA):     ")

    print(f"""
━━━  SUMMARY  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  Booking name : {MY_NAME}
  1st choice   : WARD {ward1}  on  {date1.strftime('%d %b %Y')}
  2nd choice   : WARD {ward2}  on  {date2.strftime('%d %b %Y')}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
""")

    # ── Release time ──────────────────────────────────────────────────────────
    print("What time is the timetable releasing today? (HH:MM, 24-hour)")
    print("   Examples:  19:00  for 7pm  |  20:00  for 8pm")
    print("   Press Enter to skip and start immediately.")
    raw_time = input("\n   Release time: ").strip()

    release_at = None
    if raw_time:
        try:
            h, m = map(int, raw_time.replace(".", ":").split(":"))
            now  = datetime.now()
            release_at = now.replace(hour=h, minute=m, second=0, microsecond=0)
            if release_at <= now:
                print("   That time has already passed — starting immediately.")
                release_at = None
        except Exception:
            print("   Could not read that time — starting immediately.")

    slow_start = (release_at - timedelta(minutes=5)) if release_at else None

    # ── Wait until 5 minutes before release ──────────────────────────────────
    if slow_start and slow_start > datetime.now():
        print(f"\n   Waiting until {slow_start.strftime('%H:%M')} "
              f"(5 min before release) ...  Ctrl+C to cancel\n")
        try:
            while datetime.now() < slow_start:
                left = (slow_start - datetime.now()).total_seconds()
                h2, rem = divmod(int(left), 3600)
                m2, s2  = divmod(rem, 60)
                label   = f"{h2}h {m2:02d}m {s2:02d}s" if h2 else f"{m2:02d}m {s2:02d}s"
                sys.stdout.write(f"\r   {label} until warm-up ...   ")
                sys.stdout.flush()
                time.sleep(1)
            print("\r   Warm-up started — polling every 10 seconds.   ")
        except KeyboardInterrupt:
            print("\n   Cancelled.")
            sys.exit(0)

    # ── Snapshot of current sheets ────────────────────────────────────────────
    known_sheets: set = set()
    try:
        for s in list_sheets(service):
            known_sheets.add(s["properties"]["title"])
    except Exception:
        pass

    # ── Racing loop ───────────────────────────────────────────────────────────
    print(f"\nBOT RUNNING — watching for WARD {ward1} on {date1.strftime('%d/%m')}"
          f"  (fallback: WARD {ward2} on {date2.strftime('%d/%m')})")
    if release_at:
        rapid_at = release_at - timedelta(minutes=2)
        print(f"   Slow (10s) until {rapid_at.strftime('%H:%M')}, rapid (2s) from then on.")
    print("Press Ctrl+C to stop.\n")

    attempt = 0
    done    = False

    while not done:
        try:
            attempt += 1
            ts = datetime.now().strftime("%H:%M:%S")

            # Poll speed: slow until 2 min before release, rapid from then on
            now      = datetime.now()
            rapid_at = (release_at - timedelta(minutes=2)) if release_at else None
            interval = 2 if (rapid_at is None or now >= rapid_at) else 10

            # ── Detect new sheet ──────────────────────────────────────────────
            sheets = list_sheets(service)
            titles = {s["properties"]["title"] for s in sheets}
            new    = titles - known_sheets
            if new:
                print(f"\n   [{ts}]  NEW SHEET DETECTED: {new}")
                known_sheets = titles
                interval     = 2   # always go rapid the moment a new sheet appears

            sheet_name = find_target_sheet(sheets)["properties"]["title"]

            sys.stdout.write(
                f"\r   [{ts}]  #{attempt:>4}  sheet='{sheet_name}'  "
                f"[{'RAPID' if interval == 2 else 'slow '}]   "
            )
            sys.stdout.flush()

            # ── Try first choice ──────────────────────────────────────────────
            status, detail = attempt_booking(service, sheet_name, ward1, date1)

            if status == "success":
                print(f"\n\n   BOOKED!  WARD {ward1}  on  {date1.strftime('%d/%m/%Y')}  ->  {detail}")
                done = True

            elif status == "taken":
                print(f"\n   [{ts}]  1st choice taken ({detail}) — trying 2nd choice ...")
                status2, detail2 = attempt_booking(service, sheet_name, ward2, date2)
                if status2 == "success":
                    print(f"   BOOKED!  WARD {ward2}  on  {date2.strftime('%d/%m/%Y')}  ->  {detail2}  (2nd choice)")
                    done = True
                elif status2 == "taken":
                    print(f"   Both choices are taken. Please check the timetable manually.")
                    done = True

            elif status == "error":
                print(f"\n   [{ts}]  Error: {detail}")

            # "not_found" → target date not in this sheet yet, keep polling

            if not done:
                time.sleep(2)

        except KeyboardInterrupt:
            print("\n\n   Bot stopped by user.")
            done = True
        except Exception as exc:
            print(f"\n   Unexpected error: {exc}")
            time.sleep(10)

    print(f"\nDone!  Check the spreadsheet to confirm your booking.")
    print(f"   Link: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit\n")
    input("Press Enter to exit...")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""
Timetable Bot - Cloud version (GitHub Actions).
Reads preferences from environment variables.
Uses a service account instead of OAuth — no browser login needed.
"""

import os
import sys
import time
import json
import re
from datetime import datetime, timedelta

# ─── Config ───────────────────────────────────────────────────────────────────
SPREADSHEET_ID = "1sGtcCSvpKwK8ONgV0uy9ZXl_Cwp8C-YbaXx-fNfbbGY"
MY_NAME        = "THANEN"
SCOPES         = ["https://www.googleapis.com/auth/spreadsheets"]
WARDS          = ["8PA", "8PB", "7PA"]
# ─────────────────────────────────────────────────────────────────────────────


def get_service():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    sa_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT")
    if not sa_json:
        print("ERROR: GOOGLE_SERVICE_ACCOUNT secret is not set.")
        sys.exit(1)

    sa_info = json.loads(sa_json)
    creds   = service_account.Credentials.from_service_account_info(sa_info, scopes=SCOPES)
    return build("sheets", "v4", credentials=creds)


# ─── Sheet helpers ────────────────────────────────────────────────────────────

def list_sheets(service):
    info = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    return info.get("sheets", [])


def read_sheet(service, sheet_name):
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!A1:AZ200",
    ).execute()
    return result.get("values", [])


def col_letter(idx):
    result = ""
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        result = chr(65 + rem) + result
    return result


def find_target_sheet(sheets):
    today      = datetime.now()
    candidates = []

    for sheet in sheets:
        title = sheet["properties"]["title"]
        match = re.search(r'(\d{1,2})[/\-](\d{1,2})(?:[/\-](\d{2,4}))?', title)
        if match:
            try:
                day    = int(match.group(1))
                month  = int(match.group(2))
                yr_raw = match.group(3)
                yr     = today.year
                if yr_raw:
                    yr = int(yr_raw)
                    if yr < 100:
                        yr += 2000
                candidates.append((datetime(yr, month, day), sheet))
            except ValueError:
                pass

    if not candidates:
        return sorted(sheets, key=lambda s: s["properties"]["index"])[-1]

    candidates.sort(key=lambda x: x[0], reverse=True)
    future_cutoff = today + timedelta(days=60)
    valid = [(d, s) for d, s in candidates if d <= future_cutoff]
    return valid[0][1] if valid else candidates[-1][1]


# ─── Timetable parsing ────────────────────────────────────────────────────────

def find_ward_row(data, ward_code):
    ward_key    = ward_code.upper()
    ward_pattern = re.compile(
        r'(?<![A-Z\d])' + re.escape(ward_key) + r'(?![A-Z\d])',
        re.IGNORECASE,
    )

    night_rows:      list = []
    ward_candidates: list = []

    for i, row in enumerate(data):
        if not row:
            continue
        full_text = " ".join(str(c) for c in row)
        if re.search(r'night', full_text, re.IGNORECASE):
            night_rows.append(i)
        first_two = " ".join(str(row[j]) for j in range(min(2, len(row))))
        if ward_pattern.search(first_two):
            ward_candidates.append(i)

    if not ward_candidates:
        return (night_rows[-1] if night_rows else None), None

    last_night  = night_rows[-1] if night_rows else -1
    after_night = [r for r in ward_candidates if r > last_night]
    if after_night:
        ward_row = after_night[0]
    else:
        mid      = len(data) // 2
        lower    = [r for r in ward_candidates if r >= mid]
        ward_row = lower[0] if lower else ward_candidates[-1]

    header_row = None
    for nr in reversed(night_rows):
        if nr <= ward_row:
            header_row = nr
            break

    return header_row, ward_row


def date_matches_cell(cell_text, target):
    cell = str(cell_text).strip()
    if not cell:
        return False

    day, month = target.day, target.month

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

    if cell.strip() == str(day):
        return True

    return False


def find_date_col(data, header_row_idx, target_date):
    if header_row_idx is not None:
        search_rows = [header_row_idx + o for o in range(-3, 4)
                       if 0 <= header_row_idx + o < len(data)]
        for ri in search_rows:
            for ci, cell in enumerate(data[ri]):
                if date_matches_cell(cell, target_date):
                    return ci

    for ri, row in enumerate(data):
        for ci, cell in enumerate(row):
            if date_matches_cell(cell, target_date):
                return ci

    return None


# ─── Booking ─────────────────────────────────────────────────────────────────

def attempt_booking(service, sheet_name, ward, date):
    try:
        data                      = read_sheet(service, sheet_name)
        header_row_idx, ward_row_idx = find_ward_row(data, ward)

        if ward_row_idx is None:
            return "not_found", f"Ward {ward} not in sheet '{sheet_name}'"

        col_idx = find_date_col(data, header_row_idx, date)
        if col_idx is None:
            return "not_found", f"Date {date.strftime('%d/%m/%Y')} not found"

        current = ""
        if ward_row_idx < len(data) and col_idx < len(data[ward_row_idx]):
            current = str(data[ward_row_idx][col_idx]).strip()
        if current and current.upper() not in ("", "N/A", "-"):
            return "taken", current

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


# ─── Main ─────────────────────────────────────────────────────────────────────

def parse_date(s):
    for fmt in ("%d/%m/%Y", "%d/%m/%y", "%d/%m", "%d-%m-%Y", "%d-%m"):
        try:
            if fmt in ("%d/%m", "%d-%m"):
                return datetime.strptime(s + f"/{datetime.now().year}", fmt + "/%Y")
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    raise ValueError(f"Cannot parse date: {s!r}")


def main():
    date1_str    = os.environ.get("BOOKING_DATE1", "").strip()
    ward1        = os.environ.get("BOOKING_WARD1", "").strip().upper()
    date2_str    = os.environ.get("BOOKING_DATE2", "").strip()
    ward2        = os.environ.get("BOOKING_WARD2", "").strip().upper()
    release_str  = os.environ.get("RELEASE_TIME",  "").strip()

    # Validate
    missing = [n for n, v in [("BOOKING_DATE1", date1_str), ("BOOKING_WARD1", ward1),
                               ("BOOKING_DATE2", date2_str), ("BOOKING_WARD2", ward2)] if not v]
    if missing:
        print(f"ERROR: Missing required inputs: {', '.join(missing)}")
        sys.exit(1)

    try:
        date1 = parse_date(date1_str)
        date2 = parse_date(date2_str)
    except ValueError as e:
        print(f"ERROR: {e}")
        sys.exit(1)

    release_at = None
    if release_str:
        try:
            h, m       = map(int, release_str.replace(".", ":").split(":"))
            now        = datetime.now()
            release_at = now.replace(hour=h, minute=m, second=0, microsecond=0)
            if release_at <= now:
                print(f"Release time {release_str} already passed — starting immediately.")
                release_at = None
        except Exception:
            print(f"Could not parse release time {release_str!r} — starting immediately.")

    rapid_at   = (release_at - timedelta(minutes=2)) if release_at else None
    slow_start = (release_at - timedelta(minutes=5)) if release_at else None

    print(f"Booking:    {MY_NAME}")
    print(f"1st choice: WARD {ward1}  on  {date1.strftime('%d %b %Y')}")
    print(f"2nd choice: WARD {ward2}  on  {date2.strftime('%d %b %Y')}")
    if release_at:
        print(f"Release:    {release_at.strftime('%H:%M')}  "
              f"(rapid from {rapid_at.strftime('%H:%M')})")
    print()

    print("Connecting to Google Sheets...")
    service = get_service()
    print("Connected.\n")

    # Wait until 5 min before release
    if slow_start and slow_start > datetime.now():
        print(f"Waiting until {slow_start.strftime('%H:%M')} to start warm-up polling ...")
        while datetime.now() < slow_start:
            left    = (slow_start - datetime.now()).total_seconds()
            m2, s2  = divmod(int(left), 60)
            print(f"  {m2:02d}:{s2:02d} remaining ...", flush=True)
            time.sleep(30)   # print status every 30s so Actions log stays alive
        print("Warm-up started.\n")

    known_sheets: set = set()
    for s in list_sheets(service):
        known_sheets.add(s["properties"]["title"])

    attempt = 0
    done    = False

    while not done:
        try:
            attempt += 1
            ts       = datetime.now().strftime("%H:%M:%S")
            now      = datetime.now()
            interval = 2 if (rapid_at is None or now >= rapid_at) else 10

            sheets = list_sheets(service)
            titles = {s["properties"]["title"] for s in sheets}
            new    = titles - known_sheets
            if new:
                print(f"[{ts}]  NEW SHEET: {new}")
                known_sheets = titles
                interval     = 2

            sheet_name = find_target_sheet(sheets)["properties"]["title"]
            print(f"[{ts}]  #{attempt:>4}  sheet='{sheet_name}'  "
                  f"[{'RAPID' if interval == 2 else 'slow '}]", flush=True)

            status, detail = attempt_booking(service, sheet_name, ward1, date1)

            if status == "success":
                print(f"\nBOOKED!  WARD {ward1}  on  {date1.strftime('%d/%m/%Y')}  ->  {detail}")
                done = True
            elif status == "taken":
                print(f"1st choice taken ({detail}) — trying fallback ...")
                status2, detail2 = attempt_booking(service, sheet_name, ward2, date2)
                if status2 == "success":
                    print(f"BOOKED (fallback)!  WARD {ward2}  ->  {detail2}")
                    done = True
                elif status2 == "taken":
                    print("Both choices taken. Check the timetable manually.")
                    done = True
            elif status == "error":
                print(f"Error: {detail}")

            if not done:
                time.sleep(interval)

        except KeyboardInterrupt:
            print("Stopped.")
            done = True
        except Exception as exc:
            print(f"Unexpected error: {exc}")
            time.sleep(10)

    print(f"\nDone. Spreadsheet: "
          f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit")


if __name__ == "__main__":
    main()

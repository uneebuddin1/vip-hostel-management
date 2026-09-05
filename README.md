# VIP Hostel Management System — Corrected Hostel Edition

This application is designed around the supplied **VIP Hostel IUK** workbook and the actual hostel workflow: residents, rooms, beds, allotments, exchanges, fees, vacation/gate passes, ID cards, staff occupancy, reports and session reset.

## Main dashboard
- Separate **Boys** and **Girls** occupancy figures.
- Students living, total beds, occupied beds, vacant beds and rooms.
- University-wise student counts, split into Boys/Girls.
- Room status and bed-level occupancy.
- No misleading "partial/full allotment" labels: occupancy is tracked by individual bed; rooms use operational statuses such as Vacant, Occupied, Staff Occupied, Maintenance or Closed.

## Records
- Students: add/edit/delete/search.
- Staff/non-student occupants: Teacher, Warden, Office Staff, Hostel Management, Cleaning Staff, Cooking Staff, Security and other categories. They do not require an ID number; search works by category/name.
- Rooms: add/edit/delete; change number of beds; bed labels are maintained automatically.
- Allotment: assign a student or staff occupant to a vacant bed with occupation date, fees and payment status.
- Exchange: move an occupant to another vacant bed and retain exchange history; printable exchange receipt with warden signature.
- Fees: hostel, mess, security and other payments; printable receipt.
- Vacation: closes the active allotment, makes the bed vacant, changes the student's status to Left, and automatically creates a record in **Hostel Left**. A printable gate-pass/vacation receipt is generated.
- ID cards: printable PDF hostel ID card.

## Reports / printing
- Weekly, monthly and yearly summary reports.
- Student list, room occupancy, allotment register, fee register and hostel-left register.
- Export to Excel and print/save to PDF.
- Allotment, exchange, payment and vacation/gate-pass documents include a **Warden / Hostel In-charge** signature line.

## New session
Administrator can use **Import / Export → Start New Session**. The system first creates an automatic backup and then clears operational records: students, staff occupants, allotments, payments, leave history, ID cards, exchange history and hostel-left history. Rooms, beds, users and settings are retained for the next session.

## Excel reference
The original workbook is retained unchanged at `data/source_reference.xlsx`.
The importer reads the workbook's Students, Rooms and Student_Leave sheets and records import problems in `import_issues` rather than silently discarding them. Existing records are preserved when IDs already exist.

## Installation on Windows
1. Install Python 3.11+ from https://www.python.org/downloads/ and tick **Add Python to PATH**.
2. If Windows reports the Microsoft Store Python message, use `py` instead of `python`.
3. Open this folder in Command Prompt.
4. Run:
   `py -m pip install -r requirements.txt`
5. Run:
   `py app.py`
   or double-click `launch.bat`.

## Default accounts
- Administrator: `admin` / `admin123`
- Data Entry User: `dataentry` / `data123`
- Viewer: `viewer` / `view123`

Change administrator credentials after first use.

## Data safety
- SQLite database: `data/hostel.db`
- Backups can be created from the application.
- Before restore, the current database is copied to a safety backup.
- New-session reset automatically creates a timestamped backup.
- SQL operations use parameterized queries.

## Source analysis
The supplied workbook contains a VIP Hostel IUK structure, not a school academic database. Important source sheets include Students (660 rows / 18 columns), Rooms (203 rows), Beds (780 rows), Student_Leave (129 rows), Other_Campus (75 rows), RAW_Boys, RAW_Girls, RAW_OtherCampus and Data_Quality_Report. The source includes non-student occupants such as clinic, teacher/staff and cooking/facility entries; the application therefore supports staff/non-student room occupation as a first-class record type.

## v4 fixes
- Fixed student-list refresh crash caused by referencing a non-existent `students.payment_status` field.
- Dashboard student counts now mean students with an active hostel allotment (students actually living in the hostel).
- New-session reset is transactional, backs up first, clears all operational records including import issues, resets every bed/room to vacant, and refreshes the live dashboard.
- Repaired the previous release's broken SQLite `payments -> allotments_old` foreign-key migration.
- Hostel Left records now retain selectable record IDs and include Group and Semester in display/printing.
- Student, allotment, hostel-left reports include University, Group and Semester where applicable.
- Printable PDFs can include the selected hostel logo and can be sent to the Windows default printer after PDF creation.
- Settings now includes a Choose Logo button; PNG logos also appear on the login/welcome screen.


## v5 fixes
- Full New Session clears students, staff, rooms, beds, allotments, payments, hostel-left, ID cards, exchanges and import issues while preserving users/settings and recreating Boys/Girls hostels.
- Reset Accommodation removes rooms and beds only, preserving student/staff master records.
- Added Manage Beds for adding/deleting vacant beds directly.
- Added Delete All Students with automatic backup.
- Added Choose Hostel Logo on the welcome/loading screen.
- Fixed stale dashboard/reset behavior and one-time source import so a deliberate reset is not repopulated on restart.
- Repaired legacy payments foreign-key migration.

## v8 changes
- Staff occupied room count is shown as a separate live dashboard metric.
- Student fee information is maintained directly in Students: hostel fee, mess fee, security, amount paid, balance and status calculate automatically.
- Separate fee/payment-entry workflow and payment receipts are not required for normal use.
- Student ID Card tab is removed from the main navigation.
- Fee reports read directly from student records.
- Staff Data Excel export is included.


## v9 fixes
- Fixed Staff Data Excel export: it now exports staff/non-student records with current hostel, room, bed and status instead of falling through to Hostel Left data.
- Removed the Welcome command-center tab from the main application navigation; login remains unchanged.
- Added/confirmed administrator-only Delete Selected controls for both Students and Staff Occupants.
- Student and staff deletion now safely removes dependent payments, exchanges, allotments and relevant history while releasing occupied beds, preventing SQLite foreign-key errors.
- Excel exports now create headers even when there are no records, freeze the header row, enable filters and auto-size columns.

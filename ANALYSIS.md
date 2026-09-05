# Supplied Excel Analysis — VIP Hostel IUK

The supplied workbook is a **VIP Hostel IUK management workbook**, not a school academic database.

## Key sheets and structures

- **Students** — 660 rows / 18 columns. Fields include Admission_No, Name, University, Semester, Group_No, Hostel, Room, Gender, Occupant_Type, Security, Hostel_Fee, Mess_Fee, Fee_Status, source lineage and data-completeness formula.
- **Rooms** — 203 rows / 9 columns. Room number, floor, capacity, formula-driven occupied/vacant counts, occupancy percentage and room status.
- **Beds** — 780 rows / 10 columns. Bed row, hostel, room, status, occupant type, admission number, name, source lineage and notes.
- **Student_Leave** — 129 rows / 10 columns. Admission number, student, university, semester, group, room, date of leaving, payment status and notes.
- **Other_Campus** — 75 rows / 6 columns; useful for university/campus visibility but does not have complete bed occupancy tracking.
- **RAW_Boys / RAW_Girls / RAW_OtherCampus / RAW_StudentLeave** — raw source/reference data.
- **Data_Quality_Report** — 123 source issues, including missing admission numbers, missing university/semester values and occupied beds without IDs.
- **Dashboard / Search / How_To / README** — workbook documentation, formulas and user guidance.

## Important source facts

The Students sheet identifies **618 Student occupants**, **33 rows with missing admission number**, and **9 Non-Student (Staff/Facility) occupants**. Examples of non-student occupants include Clinic, teacher/doctor entries and cooking/facility staff.

The Beds sheet contains both occupied and vacancy rows. Therefore the application tracks occupancy at **bed level**. It intentionally does not use misleading `PARTIALLY OCCUPIED` or `FULL` allotment labels. A room has an operational status (Vacant, Occupied, Staff Occupied, Maintenance, Closed), while the dashboard separately shows occupied and vacant beds.

## Improved normalized design

`Hostels → Rooms → Beds → Allotments`

`Students → Allotments → Payments`

`Staff Occupants → Allotments → Payments`

`Students → Hostel Left / ID Cards / Room Exchanges`

Additional support tables: Users, Settings, Import Issues, Audit Log.

## Data preservation approach

The original workbook is kept unchanged as `data/source_reference.xlsx`. Source lineage (`source_sheet`, `source_row`, `raw_json`) is retained for student imports. Duplicate IDs are not silently overwritten. Source rows that cannot be safely imported are logged as import issues.

## Hostel-specific workflow now implemented

1. Add/edit/delete students or non-student staff occupants.
2. Add/edit/delete rooms and change bed count without deleting occupied beds.
3. Allot a vacant bed to a student or staff category.
4. Print allotment receipt with Warden/Hostel In-charge signature line.
5. Exchange a room/bed while preserving exchange history and print exchange receipt.
6. Record hostel/mess/security payments and print payment receipt.
7. Vacate a student: active allotment closes, bed becomes vacant, student status becomes `Left`, and a **Hostel Left** record is created automatically.
8. Print vacation/gate-pass receipt with signature line.
9. Generate/print hostel student ID card.
10. View separate Boys/Girls dashboard and university-wise dashboard.
11. Generate weekly/monthly/yearly reports and export/print them.
12. Administrator can create a timestamped backup and reset operational data for a new session.

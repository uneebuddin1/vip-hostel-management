# VIP Hostel Online — Basic User & Cloud Guide

## What this version does
- Mobile-responsive web application.
- Multiple users can log in at the same time.
- Central PostgreSQL-ready database via `DATABASE_URL`.
- Roles: Administrator, Data Entry User, Viewer.
- Dashboard, Students, Staff Occupants, Rooms & Beds, Allotment.
- Administrator-only delete for students and staff occupants.
- CSV export for students and staff.

## Test accounts
- Administrator: `admin` / `admin123`
- Data Entry User: `dataentry` / `data123`
- Viewer: `viewer` / `view123`

Change these passwords before real use.

## Recommended cloud setup
1. Create a GitHub repository and upload this folder.
2. Create a Render account.
3. Create a PostgreSQL database and a Web Service from the repository.
4. Set `DATABASE_URL` to the PostgreSQL connection string and `SECRET_KEY` to a long random value.
5. Build: `pip install -r requirements-web.txt`
6. Start: `gunicorn webapp:app`
7. Open the generated `onrender.com` URL on a phone or computer.
8. Users use the same URL with their own accounts.

For production hostel records, use a paid database/hosting plan with backups rather than a temporary/free database. Do not put real student data into a public Git repository.

## Basic user guide
1. Open the web URL.
2. Log in with your assigned username/password.
3. Dashboard shows students, staff, beds and availability.
4. Students: search or add students. Administrators see Delete.
5. Staff Occupants: search or add staff. Administrators see Delete.
6. Rooms & Beds: add hostel rooms and number of beds.
7. Allotment: choose Student/Staff, choose person and a vacant bed, then allot.
8. Export CSV from Students or Staff when needed.
9. Logout when finished.

## Important
The web version is a new online interface; it does not automatically copy an existing desktop SQLite database into PostgreSQL. If the desktop database later contains real records, use a controlled migration/import step before production use.

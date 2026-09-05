import os,sys
sys.path.insert(0,os.path.dirname(__file__))
import app

db=app.DBX()
wb=app.openpyxl.load_workbook(app.SRC,data_only=True)
# Use a small headless version of the importer.
if 'Students' in wb.sheetnames:
 ws=wb['Students']
 for i,row in enumerate(ws.iter_rows(min_row=2,values_only=True),2):
  adm=app.n(row[1]); name=app.n(row[2])
  if not adm or not name: continue
  if db.one('select 1 from students where admission_no=?',(adm,)): continue
  db.x('insert into students(admission_no,name,university,semester,group_no,gender,status,source_sheet,source_row,raw_json) values(?,?,?,?,?,?,?,?,?,?)',(adm,name,app.n(row[3]),app.n(row[4]),app.n(row[6]),app.n(row[9]),'Active','Students',i,repr(row)))
if 'Rooms' in wb.sheetnames:
 ws=wb['Rooms']
 for row in ws.iter_rows(min_row=2,values_only=True):
  hostel=app.n(row[1]);room=app.n(row[2]);cap=int(row[4] or 0)
  if not hostel or not room:continue
  h=db.one('select id from hostels where name=?',(hostel,))
  if not h:
   db.x('insert or ignore into hostels(name) values(?)',(hostel,));h=db.one('select id from hostels where name=?',(hostel,))
  rr=db.one('select id from rooms where hostel_id=? and room_number=?',(h['id'],room))
  if not rr:db.x('insert into rooms(hostel_id,room_number,floor,capacity) values(?,?,?,?)',(h['id'],room,app.n(row[3]),cap))
  rr=db.one('select id from rooms where hostel_id=? and room_number=?',(h['id'],room))
  existing=db.one('select count(*) n from beds where room_id=?',(rr['id'],))['n']
  for i in range(existing+1,cap+1):db.x('insert into beds(room_id,bed_label,status) values(?,?,?)',(rr['id'],str(i),'Vacant'))
if 'Beds' in wb.sheetnames:
 ws=wb['Beds']
 for i,row in enumerate(ws.iter_rows(min_row=2,values_only=True),2):
  hostel=app.n(row[1]);room=app.n(row[2]);status='Occupied' if app.n(row[3]).lower()=='occupied' else 'Vacant';adm=app.n(row[5]);name=app.n(row[6])
  h=db.one('select id from hostels where name=?',(hostel,));rr=db.one('select id from rooms where hostel_id=? and room_number=?',(h['id'],room)) if h else None
  if not rr:continue
  label=str(row[0]);b=db.one('select id from beds where room_id=? and bed_label=?',(rr['id'],label))
  if not b:b=db.one('select id from beds where room_id=? and bed_label=?',(rr['id'],str((i-1))))
  if not b:
   bid=db.x('insert into beds(room_id,bed_label,status,notes) values(?,?,?,?)',(rr['id'],label,status,'Imported from source Beds sheet'))
  else:bid=b['id'];db.x('update beds set status=? where id=?',(status,bid))
  if status=='Occupied' and adm:
   s=db.one('select id from students where admission_no=?',(adm,))
   if s and not db.one("select 1 from allotments where student_id=? and status='Active'",(s['id'],)) and not db.one("select 1 from allotments where bed_id=? and status='Active'",(bid,)):
    db.x('insert into allotments(student_id,bed_id,occupation_date,payment_status,remarks) values(?,?,?,?,?)',(s['id'],bid,'2025-09-01','Pending','Imported from Excel current occupancy'));db.x("update beds set status='Occupied' where id=?",(bid,))
if 'Student_Leave' in wb.sheetnames:
 ws=wb['Student_Leave']
 for i,row in enumerate(ws.iter_rows(min_row=2,values_only=True),2):
  adm=app.n(row[1]);dt=row[7]
  if not dt:continue
  if hasattr(dt,'strftime'):dt=dt.strftime('%Y-%m-%d')
  s=db.one('select id,name from students where admission_no=?',(adm,)) if adm else None
  db.x('insert into leave_records(student_id,student_name_snapshot,start_date,end_date,leave_days,leave_type,application_date,approval_status,remarks,source_sheet,source_row) values(?,?,?,?,?,?,?,?,?,?,?)',((s['id'] if s else None),(s['name'] if s else app.n(row[2])),dt,dt,1,'Historical source leave',dt,'Unknown/Historical',app.n(row[9]),'Student_Leave',i))
print('Database:',app.DB)
for table in ['students','hostels','rooms','beds','allotments','leave_records']:
 print(table,db.one('select count(*) n from '+table)['n'])

import os, csv, io, secrets, hashlib
from datetime import date, datetime
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import func

app=Flask(__name__)
app.secret_key=os.environ.get('SECRET_KEY','change-this-secret-key')
dburl=os.environ.get('DATABASE_URL','sqlite:///hostel_web.db')
if dburl.startswith('postgres://'): dburl=dburl.replace('postgres://','postgresql://',1)
app.config['SQLALCHEMY_DATABASE_URI']=dburl
app.config['SQLALCHEMY_TRACK_MODIFICATIONS']=False
db=SQLAlchemy(app)

class User(db.Model):
    id=db.Column(db.Integer,primary_key=True); username=db.Column(db.String(80),unique=True,nullable=False); password_hash=db.Column(db.String(128),nullable=False); salt=db.Column(db.String(64),nullable=False); role=db.Column(db.String(30),default='Viewer'); full_name=db.Column(db.String(120)); active=db.Column(db.Boolean,default=True)
class Hostel(db.Model):
    id=db.Column(db.Integer,primary_key=True); name=db.Column(db.String(80),unique=True); gender=db.Column(db.String(20)); active=db.Column(db.Boolean,default=True)
class Room(db.Model):
    id=db.Column(db.Integer,primary_key=True); hostel_id=db.Column(db.Integer,db.ForeignKey('hostel.id')); room_number=db.Column(db.String(40)); floor=db.Column(db.String(30)); capacity=db.Column(db.Integer,default=0); status=db.Column(db.String(30),default='Vacant'); hostel=db.relationship('Hostel')
class Bed(db.Model):
    id=db.Column(db.Integer,primary_key=True); room_id=db.Column(db.Integer,db.ForeignKey('room.id')); bed_label=db.Column(db.String(40)); status=db.Column(db.String(30),default='Vacant'); room=db.relationship('Room')
class Student(db.Model):
    id=db.Column(db.Integer,primary_key=True); admission_no=db.Column(db.String(80),unique=True,nullable=False); name=db.Column(db.String(160),nullable=False); gender=db.Column(db.String(20)); university=db.Column(db.String(160)); semester=db.Column(db.String(50)); contact_number=db.Column(db.String(60)); address=db.Column(db.Text); admission_date=db.Column(db.String(20)); status=db.Column(db.String(30),default='Active'); created_at=db.Column(db.DateTime,default=datetime.utcnow)
class Staff(db.Model):
    id=db.Column(db.Integer,primary_key=True); name=db.Column(db.String(160),nullable=False); category=db.Column(db.String(80),nullable=False); gender=db.Column(db.String(20)); contact_number=db.Column(db.String(60)); address=db.Column(db.Text); notes=db.Column(db.Text); active=db.Column(db.Boolean,default=True)
class Allotment(db.Model):
    id=db.Column(db.Integer,primary_key=True); student_id=db.Column(db.Integer,db.ForeignKey('student.id')); staff_id=db.Column(db.Integer,db.ForeignKey('staff.id')); bed_id=db.Column(db.Integer,db.ForeignKey('bed.id')); occupant_category=db.Column(db.String(30),default='Student'); occupation_date=db.Column(db.String(20),nullable=False); status=db.Column(db.String(20),default='Active'); student=db.relationship('Student'); staff=db.relationship('Staff'); bed=db.relationship('Bed')

def phash(p,salt=None):
    salt=salt or secrets.token_hex(16); return hashlib.pbkdf2_hmac('sha256',p.encode(),salt.encode(),120000).hex(),salt
def verify(p,h,s): return secrets.compare_digest(phash(p,s)[0],h)
def current_user(): return User.query.get(session['uid']) if session.get('uid') else None
def editable(): return current_user() and current_user().role!='Viewer'
def admin(): return current_user() and current_user().role=='Administrator'
@app.context_processor
def ctx(): return {'me':current_user(),'editable':editable(),'admin':admin()}

@app.before_request
def guard():
    if request.endpoint not in ('login','static') and not session.get('uid'): return redirect(url_for('login'))

@app.route('/login',methods=['GET','POST'])
def login():
    if request.method=='POST':
        u=User.query.filter_by(username=request.form['username'].strip(),active=True).first()
        if u and verify(request.form['password'],u.password_hash,u.salt): session['uid']=u.id; return redirect(url_for('dashboard'))
        flash('Invalid username or password.','danger')
    return render_template('login.html')
@app.route('/logout')
def logout(): session.clear(); return redirect(url_for('login'))

@app.route('/')
@app.route('/dashboard')
def dashboard():
    students=Student.query.count(); staff=Staff.query.filter_by(active=True).count(); rooms=Room.query.count(); beds=Bed.query.count(); occupied=Allotment.query.filter_by(status='Active').count()
    return render_template('dashboard.html',students=students,staff=staff,rooms=rooms,beds=beds,occupied=occupied,available=max(beds-occupied,0))

@app.route('/students',methods=['GET','POST'])
def students():
    if request.method=='POST' and editable():
        try:
            s=Student(admission_no=request.form['admission_no'].strip(),name=request.form['name'].strip(),gender=request.form.get('gender'),university=request.form.get('university'),semester=request.form.get('semester'),contact_number=request.form.get('contact_number'),address=request.form.get('address'),admission_date=request.form.get('admission_date'),status='Active')
            db.session.add(s); db.session.commit(); flash('Student added.','success')
        except Exception as e: db.session.rollback(); flash('Could not add student: '+str(e),'danger')
        return redirect(url_for('students'))
    q=request.args.get('q','').strip(); query=Student.query
    if q: query=query.filter((Student.name.ilike(f'%{q}%'))|(Student.admission_no.ilike(f'%{q}%'))|(Student.university.ilike(f'%{q}%')))
    return render_template('students.html',rows=query.order_by(Student.name).all(),q=q)
@app.post('/students/<int:id>/delete')
def student_delete(id):
    if not admin(): flash('Administrator access required.','danger'); return redirect(url_for('students'))
    s=Student.query.get_or_404(id)
    for a in Allotment.query.filter_by(student_id=id).all():
        if a.status=='Active': a.bed.status='Vacant'
        db.session.delete(a)
    db.session.delete(s); db.session.commit(); flash('Student deleted and active bed released.','success'); return redirect(url_for('students'))

@app.route('/staff',methods=['GET','POST'])
def staff():
    if request.method=='POST' and editable():
        s=Staff(name=request.form['name'].strip(),category=request.form['category'].strip(),gender=request.form.get('gender'),contact_number=request.form.get('contact_number'),address=request.form.get('address'),notes=request.form.get('notes'),active=True); db.session.add(s); db.session.commit(); flash('Staff occupant added.','success'); return redirect(url_for('staff'))
    q=request.args.get('q','').strip(); query=Staff.query.filter_by(active=True)
    if q: query=query.filter((Staff.name.ilike(f'%{q}%'))|(Staff.category.ilike(f'%{q}%')))
    return render_template('staff.html',rows=query.order_by(Staff.name).all(),q=q)
@app.post('/staff/<int:id>/delete')
def staff_delete(id):
    if not admin(): flash('Administrator access required.','danger'); return redirect(url_for('staff'))
    s=Staff.query.get_or_404(id)
    for a in Allotment.query.filter_by(staff_id=id).all():
        if a.status=='Active': a.bed.status='Vacant'
        db.session.delete(a)
    db.session.delete(s); db.session.commit(); flash('Staff occupant deleted and active bed released.','success'); return redirect(url_for('staff'))

@app.route('/rooms',methods=['GET','POST'])
def rooms():
    if request.method=='POST' and editable():
        h=Hostel.query.get(int(request.form['hostel_id'])); r=Room(hostel_id=h.id,room_number=request.form['room_number'],floor=request.form.get('floor'),capacity=int(request.form.get('capacity') or 0)); db.session.add(r); db.session.flush()
        for i in range(1,r.capacity+1): db.session.add(Bed(room_id=r.id,bed_label=f'Bed {i}'))
        db.session.commit(); flash('Room and beds added.','success'); return redirect(url_for('rooms'))
    return render_template('rooms.html',hostels=Hostel.query.all(),rows=Room.query.order_by(Room.room_number).all())

@app.route('/allot',methods=['GET','POST'])
def allot():
    if request.method=='POST' and editable():
        cat=request.form['category']; bed=Bed.query.get(int(request.form['bed_id']))
        if bed.status=='Occupied': flash('Bed is already occupied.','danger'); return redirect(url_for('allot'))
        kwargs=dict(bed_id=bed.id,occupant_category=cat,occupation_date=request.form.get('occupation_date') or str(date.today()),status='Active')
        if cat=='Student': kwargs['student_id']=int(request.form['person_id'])
        else: kwargs['staff_id']=int(request.form['person_id'])
        db.session.add(Allotment(**kwargs)); bed.status='Occupied'; db.session.commit(); flash('Bed allotted successfully.','success'); return redirect(url_for('allot'))
    return render_template('allot.html',students=Student.query.filter_by(status='Active').order_by(Student.name).all(),staff=Staff.query.filter_by(active=True).order_by(Staff.name).all(),beds=Bed.query.filter_by(status='Vacant').join(Room).all())

@app.route('/export/<kind>')
def export(kind):
    if kind not in ('students','staff'): return redirect(url_for('dashboard'))
    out=io.StringIO(); w=csv.writer(out)
    if kind=='students':
        w.writerow(['Admission No','Name','Gender','University','Semester','Contact','Status']); [w.writerow([s.admission_no,s.name,s.gender,s.university,s.semester,s.contact_number,s.status]) for s in Student.query.order_by(Student.name).all()]
        fn='students.csv'
    else:
        w.writerow(['Name','Category','Gender','Contact','Address','Active']); [w.writerow([s.name,s.category,s.gender,s.contact_number,s.address,s.active]) for s in Staff.query.order_by(Staff.name).all()]
        fn='staff.csv'
    return send_file(io.BytesIO(out.getvalue().encode('utf-8-sig')),as_attachment=True,download_name=fn,mimetype='text/csv')

with app.app_context():
    db.create_all()
    if not User.query.first():
        for un,pw,role in [('admin','admin123','Administrator'),('dataentry','data123','Data Entry User'),('viewer','view123','Viewer')]:
            h,s=phash(pw); db.session.add(User(username=un,password_hash=h,salt=s,role=role,full_name=un.title()))
    if not Hostel.query.first(): db.session.add_all([Hostel(name='Boys',gender='Male'),Hostel(name='Girls',gender='Female')])
    db.session.commit()

if __name__=='__main__': app.run(host='0.0.0.0',port=int(os.environ.get('PORT',5000)),debug=True)

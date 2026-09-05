import os, sqlite3, hashlib, secrets, shutil, re, json, csv
from datetime import datetime, date, timedelta
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

try:
    import openpyxl
except Exception:
    openpyxl = None
try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet
except Exception:
    SimpleDocTemplate = None

APP = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(APP, 'data')
DB = os.path.join(DATA, 'hostel.db')
SRC = os.path.join(DATA, 'source_reference.xlsx')
os.makedirs(DATA, exist_ok=True)

# ---------------- database ----------------
def phash(password, salt=None):
    salt = salt or secrets.token_hex(16)
    return hashlib.pbkdf2_hmac('sha256', password.encode(), salt.encode(), 120000).hex(), salt

def verify(password, stored, salt):
    return secrets.compare_digest(phash(password, salt)[0], stored)

def dvalid(s, allow_blank=True):
    if not s and allow_blank: return True
    try: datetime.strptime(s, '%Y-%m-%d'); return True
    except Exception: return False

def clean(v):
    return '' if v is None else str(v).strip()

def iso(v):
    if hasattr(v, 'strftime'): return v.strftime('%Y-%m-%d')
    return clean(v)

SCHEMA = r'''
PRAGMA foreign_keys=ON;
CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY AUTOINCREMENT,username TEXT UNIQUE NOT NULL,password_hash TEXT NOT NULL,salt TEXT NOT NULL,role TEXT NOT NULL,full_name TEXT,active INTEGER NOT NULL DEFAULT 1,created_at TEXT DEFAULT CURRENT_TIMESTAMP);
CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY,value TEXT);
CREATE TABLE IF NOT EXISTS hostels(id INTEGER PRIMARY KEY AUTOINCREMENT,name TEXT UNIQUE NOT NULL,gender TEXT,active INTEGER DEFAULT 1);
CREATE TABLE IF NOT EXISTS rooms(id INTEGER PRIMARY KEY AUTOINCREMENT,hostel_id INTEGER NOT NULL,room_number TEXT NOT NULL,floor TEXT,capacity INTEGER NOT NULL DEFAULT 0,status TEXT DEFAULT 'Vacant',notes TEXT,UNIQUE(hostel_id,room_number),FOREIGN KEY(hostel_id) REFERENCES hostels(id));
CREATE TABLE IF NOT EXISTS students(id INTEGER PRIMARY KEY AUTOINCREMENT,admission_no TEXT UNIQUE NOT NULL,name TEXT NOT NULL,gender TEXT,university TEXT,semester TEXT,group_no TEXT,contact_number TEXT,address TEXT,father_name TEXT,mother_name TEXT,guardian_name TEXT,admission_date TEXT,status TEXT NOT NULL DEFAULT 'Active',hostel_fee REAL DEFAULT 0,mess_fee REAL DEFAULT 0,security_amount REAL DEFAULT 0,payment_status TEXT DEFAULT 'Pending',fee_balance REAL DEFAULT 0,source_sheet TEXT,source_row INTEGER,raw_json TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP,updated_at TEXT DEFAULT CURRENT_TIMESTAMP);
CREATE TABLE IF NOT EXISTS staff_occupants(id INTEGER PRIMARY KEY AUTOINCREMENT,name TEXT NOT NULL,category TEXT NOT NULL,gender TEXT,contact_number TEXT,address TEXT,notes TEXT,active INTEGER NOT NULL DEFAULT 1,created_at TEXT DEFAULT CURRENT_TIMESTAMP);
CREATE TABLE IF NOT EXISTS beds(id INTEGER PRIMARY KEY AUTOINCREMENT,room_id INTEGER NOT NULL,bed_label TEXT NOT NULL,status TEXT NOT NULL DEFAULT 'Vacant',notes TEXT,UNIQUE(room_id,bed_label),FOREIGN KEY(room_id) REFERENCES rooms(id) ON DELETE CASCADE);
CREATE TABLE IF NOT EXISTS allotments(id INTEGER PRIMARY KEY AUTOINCREMENT,student_id INTEGER,staff_id INTEGER,occupant_category TEXT NOT NULL DEFAULT 'Student',bed_id INTEGER NOT NULL,occupation_date TEXT NOT NULL,vacation_date TEXT,payment_status TEXT DEFAULT 'Pending',security_amount REAL DEFAULT 0,monthly_fee REAL DEFAULT 0,mess_fee REAL DEFAULT 0,status TEXT NOT NULL DEFAULT 'Active',remarks TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP,FOREIGN KEY(student_id) REFERENCES students(id),FOREIGN KEY(staff_id) REFERENCES staff_occupants(id),FOREIGN KEY(bed_id) REFERENCES beds(id));
CREATE UNIQUE INDEX IF NOT EXISTS ux_active_student_allotment ON allotments(student_id) WHERE status='Active' AND student_id IS NOT NULL;
CREATE UNIQUE INDEX IF NOT EXISTS ux_active_staff_allotment ON allotments(staff_id) WHERE status='Active' AND staff_id IS NOT NULL;
CREATE UNIQUE INDEX IF NOT EXISTS ux_active_bed_allotment ON allotments(bed_id) WHERE status='Active';
CREATE TABLE IF NOT EXISTS room_exchanges(id INTEGER PRIMARY KEY AUTOINCREMENT,student_id INTEGER,staff_id INTEGER,from_bed_id INTEGER,to_bed_id INTEGER,exchange_date TEXT NOT NULL,reason TEXT,remarks TEXT,FOREIGN KEY(student_id) REFERENCES students(id),FOREIGN KEY(staff_id) REFERENCES staff_occupants(id),FOREIGN KEY(from_bed_id) REFERENCES beds(id),FOREIGN KEY(to_bed_id) REFERENCES beds(id));
CREATE TABLE IF NOT EXISTS payments(id INTEGER PRIMARY KEY AUTOINCREMENT,student_id INTEGER,staff_id INTEGER,allotment_id INTEGER,payment_date TEXT NOT NULL,fee_month TEXT NOT NULL,fee_type TEXT NOT NULL,amount REAL NOT NULL,discount REAL DEFAULT 0,balance REAL DEFAULT 0,status TEXT DEFAULT 'Paid',receipt_no TEXT UNIQUE,remarks TEXT,FOREIGN KEY(student_id) REFERENCES students(id),FOREIGN KEY(staff_id) REFERENCES staff_occupants(id),FOREIGN KEY(allotment_id) REFERENCES allotments(id));
CREATE TABLE IF NOT EXISTS leave_records(id INTEGER PRIMARY KEY AUTOINCREMENT,student_id INTEGER,student_name_snapshot TEXT,room_snapshot TEXT,start_date TEXT NOT NULL,end_date TEXT NOT NULL,leave_days INTEGER NOT NULL,leave_type TEXT,reason TEXT,application_date TEXT,approval_status TEXT DEFAULT 'Pending',remarks TEXT,source_sheet TEXT,source_row INTEGER,FOREIGN KEY(student_id) REFERENCES students(id) ON DELETE SET NULL);
CREATE TABLE IF NOT EXISTS left_hostel(id INTEGER PRIMARY KEY AUTOINCREMENT,student_id INTEGER,student_name TEXT NOT NULL,admission_no TEXT,university TEXT,hostel TEXT,room TEXT,bed TEXT,occupation_date TEXT,vacation_date TEXT,reason TEXT,payment_status TEXT,remarks TEXT,left_at TEXT DEFAULT CURRENT_TIMESTAMP,FOREIGN KEY(student_id) REFERENCES students(id) ON DELETE SET NULL);
CREATE TABLE IF NOT EXISTS id_cards(id INTEGER PRIMARY KEY AUTOINCREMENT,student_id INTEGER UNIQUE NOT NULL,card_no TEXT UNIQUE NOT NULL,issue_date TEXT NOT NULL,expiry_date TEXT,active INTEGER DEFAULT 1,FOREIGN KEY(student_id) REFERENCES students(id));
CREATE TABLE IF NOT EXISTS import_issues(id INTEGER PRIMARY KEY AUTOINCREMENT,batch TEXT,source_sheet TEXT,source_row INTEGER,field TEXT,value TEXT,problem TEXT,recommended_action TEXT,status TEXT DEFAULT 'Pending');
CREATE TABLE IF NOT EXISTS audit_log(id INTEGER PRIMARY KEY AUTOINCREMENT,username TEXT,action TEXT,entity TEXT,entity_id TEXT,details TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP);
CREATE INDEX IF NOT EXISTS idx_student_name ON students(name); CREATE INDEX IF NOT EXISTS idx_student_adm ON students(admission_no); CREATE INDEX IF NOT EXISTS idx_payment_month ON payments(fee_month); CREATE INDEX IF NOT EXISTS idx_leave_date ON leave_records(start_date,end_date);
'''

class DBX:
    def __init__(self):
        self.c = sqlite3.connect(DB)
        self.c.row_factory = sqlite3.Row
        self.c.execute('PRAGMA foreign_keys=ON')
        self.migrate()
        self.seed()
    def migrate(self):
        # Migrate the previous release before applying indexes that reference the new staff_id column.
        old_cols = {r['name'] for r in self.q('PRAGMA table_info(allotments)')} if self.one("SELECT name FROM sqlite_master WHERE type='table' AND name='allotments'") else set()
        if old_cols and 'staff_id' not in old_cols:
            self.c.execute('''CREATE TABLE IF NOT EXISTS staff_occupants(id INTEGER PRIMARY KEY AUTOINCREMENT,name TEXT NOT NULL,category TEXT NOT NULL,gender TEXT,contact_number TEXT,address TEXT,notes TEXT,active INTEGER NOT NULL DEFAULT 1,created_at TEXT DEFAULT CURRENT_TIMESTAMP)''')
            self.c.execute('ALTER TABLE allotments RENAME TO allotments_old')
            self.c.execute('''CREATE TABLE allotments(id INTEGER PRIMARY KEY AUTOINCREMENT,student_id INTEGER,staff_id INTEGER,occupant_category TEXT NOT NULL DEFAULT 'Student',bed_id INTEGER NOT NULL,occupation_date TEXT NOT NULL,vacation_date TEXT,payment_status TEXT DEFAULT 'Pending',security_amount REAL DEFAULT 0,monthly_fee REAL DEFAULT 0,mess_fee REAL DEFAULT 0,status TEXT NOT NULL DEFAULT 'Active',remarks TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP,FOREIGN KEY(student_id) REFERENCES students(id),FOREIGN KEY(staff_id) REFERENCES staff_occupants(id),FOREIGN KEY(bed_id) REFERENCES beds(id))''')
            self.c.execute('''INSERT INTO allotments(id,student_id,bed_id,occupation_date,vacation_date,payment_status,security_amount,monthly_fee,mess_fee,status,remarks,created_at) SELECT id,student_id,bed_id,occupation_date,vacation_date,payment_status,security_amount,monthly_fee,mess_fee,status,remarks,created_at FROM allotments_old''')
            self.c.execute('DROP TABLE allotments_old')
            self.c.commit()
        # Repair databases from releases that left Payments pointing at allotments_old.
        psql = self.one("SELECT sql FROM sqlite_master WHERE type='table' AND name='payments'")
        if psql and 'allotments_old' in (psql['sql'] or ''):
            self.c.execute('ALTER TABLE payments RENAME TO payments_broken')
            self.c.execute('''CREATE TABLE payments(id INTEGER PRIMARY KEY AUTOINCREMENT,student_id INTEGER,staff_id INTEGER,allotment_id INTEGER,payment_date TEXT NOT NULL,fee_month TEXT NOT NULL,fee_type TEXT NOT NULL,amount REAL NOT NULL,discount REAL DEFAULT 0,balance REAL DEFAULT 0,status TEXT DEFAULT 'Paid',receipt_no TEXT UNIQUE,remarks TEXT,FOREIGN KEY(student_id) REFERENCES students(id),FOREIGN KEY(staff_id) REFERENCES staff_occupants(id),FOREIGN KEY(allotment_id) REFERENCES allotments(id))''')
            cols={r['name'] for r in self.q('PRAGMA table_info(payments_broken)')}
            wanted=['id','student_id','allotment_id','payment_date','fee_month','fee_type','amount','discount','balance','status','receipt_no','remarks']
            usable=[x for x in wanted if x in cols]
            if usable:
                colsql=','.join(usable)
                self.c.execute(f'INSERT INTO payments({colsql}) SELECT {colsql} FROM payments_broken')
            self.c.execute('DROP TABLE payments_broken')
            self.c.commit()
        self.c.executescript(SCHEMA)
        # Repair databases created by earlier releases. CREATE TABLE IF NOT EXISTS
        # does not add newly introduced columns, so explicitly add any missing
        # columns before the UI starts issuing INSERT/UPDATE statements.
        migrations = {
            'rooms': {
                'notes': "ALTER TABLE rooms ADD COLUMN notes TEXT",
            },
            'beds': {
                'notes': "ALTER TABLE beds ADD COLUMN notes TEXT",
            },
            'staff_occupants': {
                'notes': "ALTER TABLE staff_occupants ADD COLUMN notes TEXT",
                'active': "ALTER TABLE staff_occupants ADD COLUMN active INTEGER NOT NULL DEFAULT 1",
                'created_at': "ALTER TABLE staff_occupants ADD COLUMN created_at TEXT DEFAULT CURRENT_TIMESTAMP",
            },
            'students': {
                'hostel_fee': "ALTER TABLE students ADD COLUMN hostel_fee REAL DEFAULT 0",
                'mess_fee': "ALTER TABLE students ADD COLUMN mess_fee REAL DEFAULT 0",
                'security_amount': "ALTER TABLE students ADD COLUMN security_amount REAL DEFAULT 0",
                'payment_status': "ALTER TABLE students ADD COLUMN payment_status TEXT DEFAULT 'Pending'",
                'fee_balance': "ALTER TABLE students ADD COLUMN fee_balance REAL DEFAULT 0",
                'group_no': "ALTER TABLE students ADD COLUMN group_no TEXT",
                'semester': "ALTER TABLE students ADD COLUMN semester TEXT",
                'updated_at': "ALTER TABLE students ADD COLUMN updated_at TEXT DEFAULT CURRENT_TIMESTAMP",
            },
        }
        for table, cols in migrations.items():
            existing = {r['name'] for r in self.q(f'PRAGMA table_info({table})')}
            for col, sql in cols.items():
                if col not in existing:
                    try:
                        self.c.execute(sql)
                    except sqlite3.OperationalError:
                        pass
        self.c.execute("UPDATE beds SET status='Vacant' WHERE lower(status) IN ('vacancy','available','free')")
        self.c.commit()
    def seed(self):
        if not self.one('SELECT 1 FROM users LIMIT 1'):
            for u,p,r in [('admin','admin123','Administrator'),('dataentry','data123','Data Entry User'),('viewer','view123','Viewer')]:
                h,s=phash(p); self.x('INSERT INTO users(username,password_hash,salt,role,full_name) VALUES(?,?,?,?,?)',(u,h,s,r,u.title()))
        for k,v in [('institution_name','VIP Hostel IUK'),('app_title','VIP Hostel Management System'),('warden_name','Warden / Hostel In-charge'),('logo_path',''),('initial_import_done','0')]:
            self.x("INSERT OR IGNORE INTO settings(key,value) VALUES(?,?)",(k,v))
        for x,g in [('Boys','Male'),('Girls','Female')]: self.x('INSERT OR IGNORE INTO hostels(name,gender) VALUES(?,?)',(x,g))
        self.c.commit()
    def setting_value(self,k):
        r=self.c.execute('SELECT value FROM settings WHERE key=?',(k,)).fetchone(); return r['value'] if r else ''
    def commit(self): self.c.commit()
    def q(self,s,a=()): return self.c.execute(s,a).fetchall()
    def one(self,s,a=()): return self.c.execute(s,a).fetchone()
    def x(self,s,a=()): cur=self.c.execute(s,a); self.c.commit(); return cur.lastrowid
    def exec(self,s,a=()): self.c.execute(s,a); self.c.commit()
    def log(self,u,a,e='',i='',d=''): self.x('INSERT INTO audit_log(username,action,entity,entity_id,details) VALUES(?,?,?,?,?)',(u,a,e,str(i),d))

# ---------------- application ----------------
class App:
    def __init__(self):
        self.db=DBX(); self.user=None
        self.root=tk.Tk(); self.root.title('VIP Hostel Management System'); self.root.geometry('1450x900'); self.root.minsize(1180,760)
        self.style=ttk.Style(); self.style.theme_use('clam')
        self.style.configure('Title.TLabel',font=('Segoe UI',22,'bold')); self.style.configure('Sub.TLabel',font=('Segoe UI',11)); self.style.configure('Card.TLabel',font=('Segoe UI',22,'bold'))
        self.root.protocol('WM_DELETE_WINDOW', self.root.destroy)
        # Never auto-import the reference workbook. A blank/new session must remain
        # blank after restart; importing source data is always an explicit user action.
        self.show_login()
    def clear(self):
        for w in self.root.winfo_children(): w.destroy()
    def setting(self,k):
        r=self.db.one('SELECT value FROM settings WHERE key=?',(k,)); return r['value'] if r else ''
    def clock(self,label):
        if label.winfo_exists(): label.config(text=datetime.now().strftime('%d %B %Y  %H:%M:%S')); self.root.after(1000,lambda:self.clock(label))
    def show_login(self):
        self.clear(); outer=ttk.Frame(self.root,padding=40); outer.place(relx=.5,rely=.5,anchor='center')
        logo=self.setting('logo_path')
        if logo and os.path.exists(logo) and logo.lower().endswith('.png'):
            try:
                self.logo_img=tk.PhotoImage(file=logo); ttk.Label(outer,image=self.logo_img,relief='solid').grid(row=0,column=0,columnspan=2,pady=10)
            except Exception:
                ttk.Label(outer,text='HOSTEL LOGO',font=('Segoe UI',18,'bold'),relief='solid',width=18,anchor='center').grid(row=0,column=0,columnspan=2,pady=10,ipady=20)
        else:
            ttk.Label(outer,text='HOSTEL LOGO',font=('Segoe UI',18,'bold'),relief='solid',width=18,anchor='center').grid(row=0,column=0,columnspan=2,pady=10,ipady=20)
        ttk.Label(outer,text=self.setting('institution_name'),style='Title.TLabel').grid(row=1,column=0,columnspan=2)
        ttk.Label(outer,text=self.setting('app_title'),font=('Segoe UI',15)).grid(row=2,column=0,columnspan=2,pady=(3,25))
        ttk.Label(outer,text='Username').grid(row=3,column=0,sticky='w',pady=6); self.eu=ttk.Entry(outer,width=34); self.eu.grid(row=3,column=1)
        ttk.Label(outer,text='Password').grid(row=4,column=0,sticky='w',pady=6); self.ep=ttk.Entry(outer,width=34,show='*'); self.ep.grid(row=4,column=1)
        self.lclock=ttk.Label(outer); self.lclock.grid(row=5,column=0,columnspan=2,pady=12); self.clock(self.lclock)
        ttk.Button(outer,text='Choose Hostel Logo',command=self.choose_login_logo).grid(row=6,column=0,pady=8); ttk.Button(outer,text='Login',command=self.login).grid(row=6,column=1,pady=8); ttk.Button(outer,text='Exit',command=self.root.destroy).grid(row=7,column=0,columnspan=2,pady=4)
        self.root.bind('<Return>',lambda e:self.login()); self.eu.focus_set()
    def choose_login_logo(self):
        p=filedialog.askopenfilename(filetypes=[('Image Files','*.png *.jpg *.jpeg *.gif')])
        if not p:return
        try:
            self.db.exec('INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value',('logo_path',p))
            messagebox.showinfo('Logo Saved','Hostel logo saved. The welcome screen and printable documents will use it.')
            self.show_login()
        except Exception as e:
            messagebox.showerror('Logo','Unable to save logo setting.\n'+str(e))

    def login(self):
        r=self.db.one('SELECT * FROM users WHERE username=? AND active=1',(self.eu.get().strip(),))
        if not r or not verify(self.ep.get(),r['password_hash'],r['salt']): messagebox.showerror('Login failed','Invalid username or password.'); return
        self.user=dict(r); self.db.log(self.user['username'],'LOGIN'); self.main()
    def allow_edit(self): return self.user['role']!='Viewer'
    def allow_admin(self): return self.user['role']=='Administrator'
    def main(self):
        self.clear(); self.root.unbind('<Return>')
        top=ttk.Frame(self.root,padding=10); top.pack(fill='x'); ttk.Label(top,text=self.setting('app_title'),style='Title.TLabel').pack(side='left'); ttk.Label(top,text=f"  {self.user['full_name']} • {self.user['role']}").pack(side='left')
        ttk.Button(top,text='Logout',command=self.logout).pack(side='right')
        self.nb=ttk.Notebook(self.root); self.nb.pack(fill='both',expand=True,padx=8,pady=8); self.tabs={}
        for name,fn in [('Dashboard',self.dashboard),('Students',self.students),('Rooms & Beds',self.rooms),('Allotment',self.allotments),('Staff Occupants',self.staff),('Hostel Left',self.left_tab),('Reports & Printing',self.reports),('Import / Export',self.io)]: self.addtab(name,fn)
        if self.allow_admin(): self.addtab('Users / Settings',self.settings)
        self.refresh_all()
    def addtab(self,n,fn): f=ttk.Frame(self.nb,padding=10); self.nb.add(f,text=n); self.tabs[n]=f; fn(f)
    def logout(self): self.db.log(self.user['username'],'LOGOUT'); self.user=None; self.show_login()
    def tree(self,p,cols,height=15):
        w=ttk.Frame(p); w.pack(fill='both',expand=True); t=ttk.Treeview(w,columns=cols,show='headings',height=height); t.grid(row=0,column=0,sticky='nsew'); sy=ttk.Scrollbar(w,orient='vertical',command=t.yview); sy.grid(row=0,column=1,sticky='ns'); sx=ttk.Scrollbar(w,orient='horizontal',command=t.xview); sx.grid(row=1,column=0,sticky='ew'); t.configure(yscrollcommand=sy.set,xscrollcommand=sx.set); w.rowconfigure(0,weight=1); w.columnconfigure(0,weight=1)
        for c in cols: t.heading(c,text=c); t.column(c,width=max(90,min(190,10*len(c))))
        return t
    def btn(self,p,text,cmd,enabled=True): return ttk.Button(p,text=text,command=cmd,state=('normal' if enabled else 'disabled'))

    # Welcome / command center
    def welcome(self,f):
        ttk.Label(f,text='Welcome — One Screen Hostel Control',style='Title.TLabel').pack(anchor='w'); ttk.Label(f,text='Search, add, edit, delete, print and manage hostel records from here.',style='Sub.TLabel').pack(anchor='w',pady=(0,12))
        search=ttk.Frame(f); search.pack(fill='x',pady=5); ttk.Label(search,text='Search student / staff / ID / room:').pack(side='left'); self.global_search=tk.StringVar(); ttk.Entry(search,textvariable=self.global_search,width=45).pack(side='left',padx=8); ttk.Button(search,text='Search',command=self.global_search_run).pack(side='left')
        self.global_tree=self.tree(f,['Type','ID / Category','Name','University','Hostel','Room','Bed','Status'],12)
        self.global_tree.bind('<Double-1>',lambda e:self.global_edit_selected())
        actions=ttk.Frame(f); actions.pack(fill='x',pady=10)
        for txt,cmd in [('Add Student',self.student_form),('Add Staff/Teacher',self.staff_form),('Add Room',self.room_form),('Allot Room/Bed',self.allot_form),('Room Exchange',self.exchange_form),('Vacation / Gate Pass',self.vacate_selected),('Print Reports',lambda:self.nb.select(self.tabs['Reports & Printing']))]: self.btn(actions,txt,cmd,self.allow_edit()).pack(side='left',padx=3)
        self.btn(actions,'Delete Selected',self.delete_global,self.allow_admin()).pack(side='left',padx=3)
    def global_edit_selected(self):
        sel=self.global_tree.selection();
        if not sel:return
        v=self.global_tree.item(sel[0],'values');
        if v[0]=='Student':
            r=self.db.one('SELECT id FROM students WHERE admission_no=?',(v[1],));
            if r:self.student_form(r['id'])
        else:
            r=self.db.one('SELECT id FROM staff_occupants WHERE category=? AND name=?',(v[1],v[2]));
            if r:self.staff_form(r['id'])
    def global_search_run(self):
        if not hasattr(self,'global_tree'): return
        q='%'+self.global_search.get().strip()+'%'; self.global_tree.delete(*self.global_tree.get_children())
        rows=self.db.q('''SELECT 'Student' typ,s.admission_no ident,s.name,s.university,h.name hostel,r.room_number,b.bed_label,s.status
          FROM students s LEFT JOIN allotments a ON a.student_id=s.id AND a.status='Active' LEFT JOIN beds b ON b.id=a.bed_id LEFT JOIN rooms r ON r.id=b.room_id LEFT JOIN hostels h ON h.id=r.hostel_id
          WHERE s.admission_no LIKE ? OR s.name LIKE ? OR COALESCE(s.university,'') LIKE ? OR COALESCE(s.contact_number,'') LIKE ? OR COALESCE(r.room_number,'') LIKE ?
          UNION ALL SELECT 'Staff',st.category,st.name,NULL,h.name,r.room_number,b.bed_label,st.active FROM staff_occupants st LEFT JOIN allotments a ON a.staff_id=st.id AND a.status='Active' LEFT JOIN beds b ON b.id=a.bed_id LEFT JOIN rooms r ON r.id=b.room_id LEFT JOIN hostels h ON h.id=r.hostel_id
          WHERE st.name LIKE ? OR st.category LIKE ? OR COALESCE(r.room_number,'') LIKE ? ORDER BY 3''',(q,q,q,q,q,q,q,q))
        for r in rows: self.global_tree.insert('', 'end', values=tuple(r))
    def delete_global(self):
        sel=self.global_tree.selection()
        if not sel: return messagebox.showwarning('Select','Select a record first.')
        vals=self.global_tree.item(sel[0],'values'); typ=vals[0]
        if typ=='Student':
            s=self.db.one('SELECT id FROM students WHERE admission_no=?',(vals[1],));
            if s: self.delete_student_id(s['id'])
        else:
            s=self.db.one('SELECT id FROM staff_occupants WHERE category=? AND name=?',(vals[1],vals[2]))
            if s and self.allow_admin():
                if hasattr(self,'stftree'):
                    self.stftree.selection_set(str(s['id']))
                    self.delete_staff()

    # Dashboard
    def dashboard(self,f):
        head=ttk.Frame(f); head.pack(fill='x'); ttk.Label(head,text='Live Hostel Dashboard',style='Title.TLabel').pack(side='left'); ttk.Button(head,text='Refresh Live Data',command=self.refresh_dashboard).pack(side='right'); self.dcards=[]; row=ttk.Frame(f); row.pack(fill='x',pady=8)
        for label in ['Students Living','Total Beds','Occupied Beds','Vacant Beds','Rooms','Staff Occupied Rooms','Pending Fees']:
            box=ttk.LabelFrame(row,text=label,padding=14); box.pack(side='left',fill='x',expand=True,padx=4); v=tk.StringVar(value='0'); self.dcards.append(v); ttk.Label(box,textvariable=v,style='Card.TLabel').pack()
        self.dash_hostels={}
        hrow=ttk.Frame(f); hrow.pack(fill='x',pady=5)
        for hname in ['Boys','Girls']:
            box=ttk.LabelFrame(hrow,text=hname+' Hostel',padding=10); box.pack(side='left',fill='x',expand=True,padx=5); v=tk.StringVar(); self.dash_hostels[hname]=v; ttk.Label(box,textvariable=v,font=('Segoe UI',12,'bold')).pack(anchor='w')
        ttk.Label(f,text='University-wise Students Living',font=('Segoe UI',15,'bold')).pack(anchor='w',pady=(12,5)); self.uni_tree=self.tree(f,['University','Students Living','Boys','Girls'],7)
        ttk.Label(f,text='Staff Accommodation — Who is Living in Which Room',font=('Segoe UI',15,'bold')).pack(anchor='w',pady=(12,5)); self.staff_dash=self.tree(f,['Category','Staff Member','Hostel','Room','Bed','Occupation Date'],7)
        ttk.Label(f,text='Room Status',font=('Segoe UI',15,'bold')).pack(anchor='w',pady=(12,5)); self.dash=self.tree(f,['Hostel','Room','Beds','Occupied','Vacant','Room Status'],7)
    def refresh_dashboard(self):
        if not hasattr(self,'dcards'): return
        vals=[self.db.one("SELECT count(*) n FROM allotments WHERE status='Active' AND student_id IS NOT NULL")['n'],self.db.one("SELECT count(*) n FROM beds WHERE status!='Blocked'")['n'],self.db.one("SELECT count(*) n FROM beds WHERE status='Occupied'")['n'],self.db.one("SELECT count(*) n FROM beds WHERE status='Vacant'")['n'],self.db.one("SELECT count(*) n FROM rooms")['n'],self.db.one("SELECT count(DISTINCT r.id) n FROM allotments a JOIN staff_occupants st ON st.id=a.staff_id JOIN beds b ON b.id=a.bed_id JOIN rooms r ON r.id=b.room_id WHERE a.status='Active'")['n'],self.db.one("SELECT coalesce(sum(fee_balance),0) n FROM students WHERE status != 'Deleted'")['n']]
        for v,x in zip(self.dcards,vals): v.set(f'{x:,.0f}')
        for h in ['Boys','Girls']:
            x=self.db.one("SELECT count(*) n FROM allotments a JOIN students s ON s.id=a.student_id JOIN beds b ON b.id=a.bed_id JOIN rooms r ON r.id=b.room_id JOIN hostels h ON h.id=r.hostel_id WHERE h.name=? AND a.status='Active'",(h,))['n']
            beds=self.db.one("SELECT count(*) n FROM beds b JOIN rooms r ON r.id=b.room_id JOIN hostels h ON h.id=r.hostel_id WHERE h.name=? AND b.status!='Blocked'",(h,))['n']
            occ=self.db.one("SELECT count(*) n FROM beds b JOIN rooms r ON r.id=b.room_id JOIN hostels h ON h.id=r.hostel_id WHERE h.name=? AND b.status='Occupied'",(h,))['n']
            staff_count=self.db.one("SELECT count(*) n FROM allotments a JOIN staff_occupants st ON st.id=a.staff_id JOIN beds b ON b.id=a.bed_id JOIN rooms r ON r.id=b.room_id JOIN hostels h ON h.id=r.hostel_id WHERE h.name=? AND a.status='Active'",(h,))['n']
            self.dash_hostels[h].set(f'Students: {x}    Staff: {staff_count}    Beds: {beds}    Occupied: {occ}    Vacant: {max(0,beds-occ)}')
        self.uni_tree.delete(*self.uni_tree.get_children())
        for r in self.db.q("""SELECT COALESCE(s.university,'Unknown') u,count(*) total,sum(CASE WHEN h.name='Boys' THEN 1 ELSE 0 END) boys,sum(CASE WHEN h.name='Girls' THEN 1 ELSE 0 END) girls FROM students s JOIN allotments a ON a.student_id=s.id AND a.status='Active' JOIN beds b ON b.id=a.bed_id JOIN rooms rm ON rm.id=b.room_id JOIN hostels h ON h.id=rm.hostel_id GROUP BY COALESCE(s.university,'Unknown') ORDER BY total DESC"""):
            self.uni_tree.insert('', 'end',values=(r['u'],r['total'],r['boys'],r['girls']))
        self.staff_dash.delete(*self.staff_dash.get_children())
        for r in self.db.q("""SELECT st.category,st.name,h.name hostel,r.room_number,b.bed_label,a.occupation_date FROM allotments a JOIN staff_occupants st ON st.id=a.staff_id JOIN beds b ON b.id=a.bed_id JOIN rooms r ON r.id=b.room_id JOIN hostels h ON h.id=r.hostel_id WHERE a.status='Active' ORDER BY st.category,st.name"""):
            self.staff_dash.insert('', 'end', values=tuple(r))
        self.dash.delete(*self.dash.get_children())
        for r in self.db.q("""SELECT h.name hostel,rm.room_number,count(b.id) beds,sum(CASE WHEN b.status='Occupied' THEN 1 ELSE 0 END) occ,sum(CASE WHEN b.status='Vacant' THEN 1 ELSE 0 END) vac,rm.status FROM rooms rm JOIN hostels h ON h.id=rm.hostel_id LEFT JOIN beds b ON b.room_id=rm.id GROUP BY rm.id ORDER BY h.name,rm.room_number"""):
            self.dash.insert('', 'end',values=(r['hostel'],r['room_number'],r['beds'],r['occ'],r['vac'],r['status']))

    # Students
    def students(self,f):
        # Primary student actions are deliberately kept in the Students tab.
        # Delete is administrator-only, but remains visibly present for non-admin users.
        bar=ttk.Frame(f); bar.pack(fill='x')
        self.btn(bar,'➕ Add Student',self.student_form,self.allow_edit()).pack(side='left',padx=(0,4))
        self.btn(bar,'✏ Edit Selected',self.edit_student,self.allow_edit()).pack(side='left',padx=4)
        self.btn(bar,'🗑 Delete Student',self.delete_student,self.allow_admin()).pack(side='left',padx=4)
        self.btn(bar,'🗑 Delete All Students',self.delete_all_students,self.allow_admin()).pack(side='left',padx=4)
        ttk.Label(bar,text='Search:').pack(side='left',padx=(18,5)); self.ss=tk.StringVar(); ttk.Entry(bar,textvariable=self.ss,width=40).pack(side='left'); ttk.Button(bar,text='Search',command=self.refresh_students).pack(side='left',padx=5); ttk.Button(bar,text='Refresh',command=self.refresh_students).pack(side='left')
        self.str=self.tree(f,['ID','Name','University','Group','Semester','Gender','Hostel','Room','Bed','Occupation Date','Hostel Fee','Mess Fee','Balance','Fee Status','Status'],18); self.str.bind('<Double-1>',lambda e:self.edit_student())
        self.str.bind('<Button-3>',lambda e:self._student_context(e))
        b=ttk.Frame(f); b.pack(fill='x',pady=5); ttk.Button(b,text='Vacation / Gate Pass',command=self.vacate_selected).pack(side='left',padx=4)
        ttk.Label(f,text=('Delete Student is available to Administrators only. Select a student, then click “🗑 Delete Student”.' if self.allow_admin() else 'Student deletion requires Administrator login.'),foreground='gray').pack(anchor='w',pady=(2,0))

    def _student_context(self,event):
        row=self.str.identify_row(event.y)
        if not row:return
        self.str.selection_set(row)
        menu=tk.Menu(self.root,tearoff=0)
        menu.add_command(label='Edit Student',command=self.edit_student,state=('normal' if self.allow_edit() else 'disabled'))
        menu.add_command(label='Delete Student',command=self.delete_student,state=('normal' if self.allow_admin() else 'disabled'))
        menu.tk_popup(event.x_root,event.y_root)
    def refresh_students(self):
        if not hasattr(self,'str'): return
        self.str.delete(*self.str.get_children()); q='%'+self.ss.get().strip()+'%'
        rows=self.db.q("""SELECT s.id,s.admission_no,s.name,s.university,s.group_no,s.semester,s.gender,h.name hostel,r.room_number,b.bed_label,a.occupation_date,COALESCE(a.monthly_fee,s.hostel_fee,0),COALESCE(a.mess_fee,s.mess_fee,0),COALESCE((SELECT COALESCE(SUM(p.balance),0) FROM payments p WHERE p.student_id=s.id AND p.status IN ('Pending','Partial')),s.fee_balance,0),COALESCE(a.payment_status,s.payment_status,'-'),CASE WHEN a.id IS NOT NULL THEN 'Living' ELSE s.status END current_status FROM students s LEFT JOIN allotments a ON a.student_id=s.id AND a.status='Active' LEFT JOIN beds b ON b.id=a.bed_id LEFT JOIN rooms r ON r.id=b.room_id LEFT JOIN hostels h ON h.id=r.hostel_id WHERE s.admission_no LIKE ? OR s.name LIKE ? OR COALESCE(s.father_name,'') LIKE ? OR COALESCE(s.mother_name,'') LIKE ? OR COALESCE(s.university,'') LIKE ? OR COALESCE(s.contact_number,'') LIKE ? OR COALESCE(r.room_number,'') LIKE ? ORDER BY s.name""",(q,q,q,q,q,q,q))
        for r in rows: self.str.insert('', 'end',iid=r['id'],values=tuple(r)[1:])
    def student_form(self,sid=None):
        if sid is None and not self.allow_edit(): return
        w=tk.Toplevel(self.root); w.title('Student Record'); w.geometry('760x700'); v={}; old=self.db.one('SELECT * FROM students WHERE id=?',(sid,)) if sid else None
        fields=[('Student / Admission ID','admission_no'),('Student Name','name'),('Gender','gender'),('University','university'),('Semester','semester'),('Group No','group_no'),('Contact Number','contact_number'),('Address','address'),("Father's Name",'father_name'),("Mother's Name",'mother_name'),('Guardian','guardian_name'),('Admission Date','admission_date'),('Hostel Fee','hostel_fee'),('Mess Fee','mess_fee'),('Security Amount','security_amount'),('Amount Paid','amount_paid'),('Status','status')]
        frm=ttk.Frame(w,padding=15); frm.pack(fill='both',expand=True)
        for i,(lab,key) in enumerate(fields):
            ttk.Label(frm,text=lab).grid(row=i,column=0,sticky='w',padx=5,pady=5)
            if key in ('gender','status'): e=ttk.Combobox(frm,values=(['Male','Female','Other'] if key=='gender' else ['Active','Inactive','Left']),state='readonly',width=34)
            else: e=ttk.Entry(frm,width=37)
            e.grid(row=i,column=1,sticky='w'); v[key]=e; 
            if old and key in old.keys():
                val=clean(old[key])
                if key=='amount_paid':
                    total=float(old['hostel_fee'] or 0)+float(old['mess_fee'] or 0)+float(old['security_amount'] or 0); val=str(max(0,total-float(old['fee_balance'] or 0)))
            else: val='Active' if key=='status' else ''
            e.insert(0,val)
        def save():
            data={k:e.get().strip() for k,e in v.items()};
            if not data['admission_no'] or not data['name']: return messagebox.showerror('Required','Student ID and Name are required.',parent=w)
            if not dvalid(data['admission_date']): return messagebox.showerror('Date','Admission date must be YYYY-MM-DD.',parent=w)
            if data['contact_number'] and not re.fullmatch(r'[0-9+()\- ]{7,20}',data['contact_number']): return messagebox.showerror('Phone','Invalid contact number.',parent=w)
            try:
                for k in ('hostel_fee','mess_fee','security_amount','amount_paid'): data[k]=float(data[k] or 0)
                total_fee=data['hostel_fee']+data['mess_fee']+data['security_amount']; data['fee_balance']=max(0,total_fee-data['amount_paid']); data['payment_status']='Paid' if data['fee_balance']<=0 and total_fee>0 else ('Pending' if data['amount_paid']<=0 else 'Partial') if total_fee>0 else 'Not Applicable'
            except ValueError: return messagebox.showerror('Fee','Fee amounts must be numbers.',parent=w)
            dup=self.db.one('SELECT id FROM students WHERE admission_no=? AND id<>?',(data['admission_no'],sid or 0));
            if dup:return messagebox.showerror('Duplicate','That Student/Admission ID already exists.',parent=w)
            if sid:self.db.x('UPDATE students SET admission_no=?,name=?,gender=?,university=?,semester=?,group_no=?,contact_number=?,address=?,father_name=?,mother_name=?,guardian_name=?,admission_date=?,hostel_fee=?,mess_fee=?,security_amount=?,fee_balance=?,payment_status=?,status=?,updated_at=CURRENT_TIMESTAMP WHERE id=?',(*[data[k] for k in ['admission_no','name','gender','university','semester','group_no','contact_number','address','father_name','mother_name','guardian_name','admission_date','hostel_fee','mess_fee','security_amount','fee_balance','payment_status','status']],sid)); self.db.log(self.user['username'],'UPDATE','Student',sid)
            else:sid2=self.db.x('INSERT INTO students(admission_no,name,gender,university,semester,group_no,contact_number,address,father_name,mother_name,guardian_name,admission_date,hostel_fee,mess_fee,security_amount,fee_balance,payment_status,status) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',tuple(data[k] for k in ['admission_no','name','gender','university','semester','group_no','contact_number','address','father_name','mother_name','guardian_name','admission_date','hostel_fee','mess_fee','security_amount','fee_balance','payment_status','status'])); self.db.log(self.user['username'],'ADD','Student',sid2)
            w.destroy(); self.refresh_all()
            if messagebox.askyesno('Accommodation','Student saved. Allot a hostel room/bed now?'):
                self.allot_form(student_id=(sid if sid else sid2))
        b=ttk.Frame(frm); b.grid(row=len(fields),column=0,columnspan=2,pady=15); ttk.Button(b,text='Save / Update',command=save).pack(side='left'); ttk.Button(b,text='Clear',command=lambda:[e.delete(0,'end') for e in v.values()]).pack(side='left',padx=5); ttk.Button(b,text='Close',command=w.destroy).pack(side='left')
    def edit_student(self):
        sel=self.str.selection();
        if sel:self.student_form(int(sel[0]))
    def delete_student(self):
        sel=self.str.selection();
        if sel:self.delete_student_id(int(sel[0]))
    def delete_all_students(self):
        if not self.allow_admin(): return
        n=self.db.one('SELECT count(*) n FROM students')['n']
        if not n:return messagebox.showinfo('Students','There are no student records to delete.')
        if not messagebox.askyesno('DANGER — Delete All Students',f'Delete ALL {n} student records and their allotments, payments, ID cards and hostel-left history? This cannot be undone. A backup will be created first.'):
            return
        bp=os.path.join(DATA,'backup_before_delete_all_students_'+datetime.now().strftime('%Y%m%d_%H%M%S')+'.db'); self.db.c.commit(); shutil.copy2(DB,bp)
        try:
            self.db.c.execute('BEGIN')
            for t in ['room_exchanges','payments','id_cards','leave_records','left_hostel','allotments','students']:
                self.db.c.execute('DELETE FROM '+t)
            self.db.c.execute("UPDATE beds SET status='Vacant'")
            self.db.c.execute("UPDATE rooms SET status='Vacant'")
            self.db.c.commit()
        except Exception as e:
            self.db.c.rollback(); return messagebox.showerror('Delete failed','No records were deleted.\n\n'+str(e))
        self.db.log(self.user['username'],'DELETE_ALL','Students','ALL',f'backup={bp}')
        self.refresh_all(); messagebox.showinfo('Completed',f'All student records were deleted.\n\nBackup: {bp}')

    def delete_student_id(self,sid):
        if not self.allow_admin(): return
        student=self.db.one('SELECT id,admission_no,name FROM students WHERE id=?',(sid,))
        if not student:return
        if not messagebox.askyesno('Confirm delete',f"Delete student {student['name']} ({student['admission_no']}) and all related records? This cannot be undone."):
            return
        # Release any occupied bed first, then remove dependent records in FK-safe order.
        self.db.c.execute('BEGIN')
        try:
            active=self.db.q("SELECT bed_id FROM allotments WHERE student_id=? AND status='Active'",(sid,))
            for a in active:
                self.db.x("UPDATE beds SET status='Vacant' WHERE id=?",(a['bed_id'],))
            self.db.exec('DELETE FROM payments WHERE student_id=?',(sid,))
            self.db.exec('DELETE FROM room_exchanges WHERE student_id=?',(sid,))
            self.db.exec('DELETE FROM id_cards WHERE student_id=?',(sid,))
            self.db.exec('DELETE FROM leave_records WHERE student_id=?',(sid,))
            self.db.exec('DELETE FROM left_hostel WHERE student_id=?',(sid,))
            self.db.exec('DELETE FROM allotments WHERE student_id=?',(sid,))
            self.db.exec('DELETE FROM students WHERE id=?',(sid,))
            self.db.c.commit()
            self.db.log(self.user['username'],'DELETE','Student',sid, f"admission_no={student['admission_no']}")
        except Exception as e:
            self.db.c.rollback()
            return messagebox.showerror('Delete failed','No records were deleted.\\n\\n'+str(e))
        self.refresh_all()
        messagebox.showinfo('Deleted',f"Student {student['name']} was deleted successfully.")

    # Rooms/beds
    def rooms(self,f):
        bar=ttk.Frame(f); bar.pack(fill='x'); self.btn(bar,'Add Room',self.room_form,self.allow_edit()).pack(side='left'); self.btn(bar,'Edit Room / Beds',self.edit_room,self.allow_edit()).pack(side='left',padx=4); self.btn(bar,'Manage Beds',self.manage_beds,self.allow_edit()).pack(side='left',padx=4); self.btn(bar,'Delete Room',self.delete_room,self.allow_admin()).pack(side='left'); self.btn(bar,'Reset Accommodation',self.reset_accommodation,self.allow_admin()).pack(side='left',padx=4); ttk.Label(bar,text='Search').pack(side='left',padx=(18,5)); self.rs=tk.StringVar(); ttk.Entry(bar,textvariable=self.rs,width=30).pack(side='left'); ttk.Button(bar,text='Search',command=self.refresh_rooms).pack(side='left',padx=5)
        self.rtree=self.tree(f,['Hostel','Room','Floor','Beds','Occupied','Vacant','Room Status','Notes'],18)
    def refresh_rooms(self):
        if not hasattr(self,'rtree'):return
        self.rtree.delete(*self.rtree.get_children()); q='%'+self.rs.get().strip()+'%'
        for r in self.db.q('''SELECT h.name,rm.room_number,rm.floor,count(b.id),sum(CASE WHEN b.status='Occupied' THEN 1 ELSE 0 END),sum(CASE WHEN b.status='Vacant' THEN 1 ELSE 0 END),rm.status,rm.notes,rm.id FROM rooms rm JOIN hostels h ON h.id=rm.hostel_id LEFT JOIN beds b ON b.room_id=rm.id WHERE h.name LIKE ? OR rm.room_number LIKE ? OR rm.status LIKE ? GROUP BY rm.id ORDER BY h.name,rm.room_number''',(q,q,q)): self.rtree.insert('', 'end',iid=r['id'],values=tuple(r)[:-1])
    def room_form(self,rid=None):
        w=tk.Toplevel(self.root);w.title('Room / Bed Setup');w.geometry('560x430');old=self.db.one('SELECT rm.*,h.name hostel FROM rooms rm JOIN hostels h ON h.id=rm.hostel_id WHERE rm.id=?',(rid,)) if rid else None
        v={}; fields=[('Hostel','hostel'),('Room Number','room_number'),('Floor','floor'),('Number of Beds','capacity'),('Room Status','status'),('Notes','notes')]
        frm=ttk.Frame(w,padding=20);frm.pack(fill='both',expand=True)
        for i,(lab,key) in enumerate(fields):
            ttk.Label(frm,text=lab).grid(row=i,column=0,sticky='w',pady=7); vals=['Boys','Girls'] if key=='hostel' else ['Vacant','Occupied','Staff Occupied','Maintenance','Closed'] if key=='status' else None
            e=ttk.Combobox(frm,values=vals,state='readonly',width=30) if vals else ttk.Entry(frm,width=32); e.grid(row=i,column=1);v[key]=e; e.insert(0,clean(old[key]) if old else ('Vacant' if key=='status' else ''))
        def save():
            d={k:e.get().strip() for k,e in v.items()}
            if not d['room_number'] or not d['hostel']:return messagebox.showerror('Required','Hostel and room number are required.',parent=w)
            try: cap=int(d['capacity'])
            except:return messagebox.showerror('Beds','Number of beds must be a whole number.',parent=w)
            if cap<0:return messagebox.showerror('Beds','Number of beds cannot be negative.',parent=w)
            h=self.db.one('SELECT id FROM hostels WHERE name=?',(d['hostel'],))
            if not h:
                return messagebox.showerror('Hostel','Selected hostel does not exist. Please use Settings/hostel setup or recreate the default Boys/Girls hostels.',parent=w)
            if rid:
                used=self.db.one("SELECT count(*) n FROM beds WHERE room_id=? AND status='Occupied'",(rid,))['n']
                if cap<used:return messagebox.showerror('Beds',f'Cannot reduce below {used} occupied beds.',parent=w)
                self.db.x('UPDATE rooms SET hostel_id=?,room_number=?,floor=?,capacity=?,status=?,notes=? WHERE id=?',(h['id'],d['room_number'],d['floor'],cap,d['status'],d['notes'],rid)); self.sync_beds(rid,cap);self.db.log(self.user['username'],'UPDATE','Room',rid)
            else:
                if self.db.one('SELECT 1 FROM rooms WHERE hostel_id=? AND room_number=?',(h['id'],d['room_number'])):return messagebox.showerror('Duplicate','Room already exists.',parent=w)
                rid2=self.db.x('INSERT INTO rooms(hostel_id,room_number,floor,capacity,status,notes) VALUES(?,?,?,?,?,?)',(h['id'],d['room_number'],d['floor'],cap,d['status'],d['notes']));self.sync_beds(rid2,cap);self.db.log(self.user['username'],'ADD','Room',rid2)
            w.destroy();self.refresh_all()
        ttk.Button(frm,text='Save Room',command=save).grid(row=6,column=0,columnspan=2,pady=15)
    def sync_beds(self,rid,cap):
        existing=self.db.q('SELECT id,bed_label,status FROM beds WHERE room_id=? ORDER BY id',(rid,)); n=len(existing)
        for i in range(n+1,cap+1): self.db.x('INSERT INTO beds(room_id,bed_label,status) VALUES(?,?,?)',(rid,str(i),'Vacant'))
        if cap<n:
            for r in existing[cap:]:
                if r['status']!='Occupied': self.db.exec('DELETE FROM beds WHERE id=?',(r['id'],))
    def manage_beds(self):
        sel=self.rtree.selection()
        if not sel:return messagebox.showwarning('Select Room','Select a room first.')
        rid=int(sel[0]); room=self.db.one('SELECT r.*,h.name hostel FROM rooms r JOIN hostels h ON h.id=r.hostel_id WHERE r.id=?',(rid,))
        if not room:return
        w=tk.Toplevel(self.root); w.title(f"Manage Beds — {room['hostel']} / Room {room['room_number']}"); w.geometry('560x470')
        ttk.Label(w,text=f"{room['hostel']} • Room {room['room_number']}",font=('Segoe UI',16,'bold')).pack(anchor='w',padx=15,pady=10)
        tree=self.tree(w,['Bed','Status','Notes'],12)
        def refresh():
            tree.delete(*tree.get_children())
            for b in self.db.q('SELECT id,bed_label,status,notes FROM beds WHERE room_id=? ORDER BY id',(rid,)): tree.insert('', 'end',iid=b['id'],values=(b['bed_label'],b['status'],b['notes'] or ''))
        refresh()
        controls=ttk.Frame(w);controls.pack(fill='x',padx=15,pady=8)
        ttk.Label(controls,text='New bed label').pack(side='left'); lab=tk.StringVar();ttk.Entry(controls,textvariable=lab,width=12).pack(side='left',padx=5)
        def add():
            label=lab.get().strip()
            if not label:return
            try:self.db.x('INSERT INTO beds(room_id,bed_label,status) VALUES(?,?,?)',(rid,label,'Vacant'));lab.set('');refresh();self.refresh_all()
            except sqlite3.IntegrityError:messagebox.showerror('Duplicate','That bed label already exists in this room.',parent=w)
        def delete():
            ss=tree.selection()
            if not ss:return messagebox.showwarning('Select','Select a bed first.',parent=w)
            bid=int(ss[0]); b=self.db.one('SELECT status FROM beds WHERE id=?',(bid,))
            if b and b['status']=='Occupied':return messagebox.showerror('Occupied','Occupied beds cannot be deleted. Vacate/exchange the occupant first.',parent=w)
            if messagebox.askyesno('Delete Bed','Delete this vacant bed?',parent=w):self.db.exec('DELETE FROM beds WHERE id=?',(bid,));refresh();self.refresh_all()
        ttk.Button(controls,text='Add Bed',command=add).pack(side='left',padx=4);ttk.Button(controls,text='Delete Selected Vacant Bed',command=delete).pack(side='left',padx=4)
        ttk.Label(w,text='You can also change the room capacity using Edit Room / Beds. New capacity creates vacant beds automatically.',wraplength=510).pack(anchor='w',padx=15,pady=10)
        ttk.Button(w,text='Close',command=w.destroy).pack(pady=8)

    def edit_room(self):
        s=self.rtree.selection();
        if s:self.room_form(int(s[0]))
    def reset_accommodation(self):
        if not self.allow_admin(): return
        rooms=self.db.one('SELECT count(*) n FROM rooms')['n']; beds=self.db.one('SELECT count(*) n FROM beds')['n']
        if not messagebox.askyesno('RESET ACCOMMODATION',f'Remove ALL rooms and beds ({rooms} rooms / {beds} beds)? All active allotments will also be removed and occupants will become unallotted. Students and staff records will be kept. A backup will be created first.'):
            return
        bp=os.path.join(DATA,'backup_before_reset_accommodation_'+datetime.now().strftime('%Y%m%d_%H%M%S')+'.db'); self.db.c.commit(); shutil.copy2(DB,bp)
        try:
            self.db.c.execute('BEGIN')
            self.db.c.execute('DELETE FROM room_exchanges')
            self.db.c.execute('DELETE FROM allotments')
            self.db.c.execute('DELETE FROM beds')
            self.db.c.execute('DELETE FROM rooms')
            self.db.c.commit()
        except Exception as e:
            self.db.c.rollback(); return messagebox.showerror('Reset failed','Accommodation was not changed.\n\n'+str(e))
        self.db.log(self.user['username'],'RESET','Accommodation','ALL',f'backup={bp}')
        self.refresh_all(); messagebox.showinfo('Accommodation Reset',f'All rooms and beds were removed. You can now add new rooms and bed counts for the new session.\n\nBackup: {bp}')

    def delete_room(self):
        s=self.rtree.selection();
        if not s:return
        rid=int(s[0]); occ=self.db.one("SELECT count(*) n FROM beds WHERE room_id=? AND status='Occupied'",(rid,))['n'];
        if occ:return messagebox.showerror('Cannot delete','Room has occupied beds. Vacate them first.')
        if messagebox.askyesno('Delete','Delete this room and its vacant beds?'):self.db.exec('DELETE FROM rooms WHERE id=?',(rid,));self.db.log(self.user['username'],'DELETE','Room',rid);self.refresh_all()

    # Staff occupants
    def staff(self,f):
        # Staff deletion is exposed directly in the Staff Occupants tab.
        bar=ttk.Frame(f);bar.pack(fill='x')
        self.btn(bar,'➕ Add Staff / Teacher / Worker',self.staff_form,self.allow_edit()).pack(side='left',padx=(0,4))
        self.btn(bar,'✏ Edit Selected',self.edit_staff,self.allow_edit()).pack(side='left',padx=4)
        self.btn(bar,'🗑 Delete Staff / Occupant',self.delete_staff,self.allow_admin()).pack(side='left',padx=4)
        ttk.Label(bar,text='Search').pack(side='left',padx=(18,5));self.stfs=tk.StringVar();ttk.Entry(bar,textvariable=self.stfs,width=35).pack(side='left');ttk.Button(bar,text='Search',command=self.refresh_staff).pack(side='left',padx=5);ttk.Button(bar,text='Refresh',command=self.refresh_staff).pack(side='left',padx=4);self.stftree=self.tree(f,['Category','Name','Gender','Contact','Hostel','Room','Bed','Status'],18)
        self.stftree.bind('<Double-1>',lambda e:self.edit_staff()); self.stftree.bind('<Button-3>',lambda e:self._staff_context(e))
        ttk.Label(f,text=('Delete Staff / Occupant is available to Administrators only. Select a record, then click “🗑 Delete Staff / Occupant”.' if self.allow_admin() else 'Staff deletion requires Administrator login.'),foreground='gray').pack(anchor='w',pady=(2,0))

    def _staff_context(self,event):
        row=self.stftree.identify_row(event.y)
        if not row:return
        self.stftree.selection_set(row)
        menu=tk.Menu(self.root,tearoff=0)
        menu.add_command(label='Edit Staff / Occupant',command=self.edit_staff,state=('normal' if self.allow_edit() else 'disabled'))
        menu.add_command(label='Delete Staff / Occupant',command=self.delete_staff,state=('normal' if self.allow_admin() else 'disabled'))
        menu.tk_popup(event.x_root,event.y_root)
    def refresh_staff(self):
        if not hasattr(self,'stftree'):return
        self.stftree.delete(*self.stftree.get_children());q='%'+self.stfs.get().strip()+'%'
        for r in self.db.q('''SELECT st.id,st.category,st.name,st.gender,st.contact_number,h.name,r.room_number,b.bed_label,CASE WHEN a.id IS NULL THEN 'Not Allotted' ELSE 'Living' END FROM staff_occupants st LEFT JOIN allotments a ON a.staff_id=st.id AND a.status='Active' LEFT JOIN beds b ON b.id=a.bed_id LEFT JOIN rooms r ON r.id=b.room_id LEFT JOIN hostels h ON h.id=r.hostel_id WHERE st.name LIKE ? OR st.category LIKE ? ORDER BY st.category,st.name''',(q,q)):self.stftree.insert('', 'end',iid=r['id'],values=tuple(r)[1:])
    def staff_form(self,sid=None):
        w=tk.Toplevel(self.root);w.title('Staff / Non-Student Occupant');w.geometry('560x430');old=self.db.one('SELECT * FROM staff_occupants WHERE id=?',(sid,)) if sid else None;v={}
        fields=[('Category','category'),('Name','name'),('Gender','gender'),('Contact','contact_number'),('Address','address'),('Notes','notes')];frm=ttk.Frame(w,padding=20);frm.pack(fill='both',expand=True)
        cats=['Teacher','Warden','Office Staff','Hostel Management','Cleaning Staff','Cooking Staff','Security','Other Staff','Guest','Facility']
        for i,(lab,k) in enumerate(fields):
            ttk.Label(frm,text=lab).grid(row=i,column=0,sticky='w',pady=6); e=ttk.Combobox(frm,values=cats,state='readonly',width=30) if k=='category' else ttk.Combobox(frm,values=['Male','Female','Other'],state='readonly',width=30) if k=='gender' else ttk.Entry(frm,width=32);e.grid(row=i,column=1);v[k]=e;e.insert(0,clean(old[k]) if old else '')
        def save():
            d={k:e.get().strip() for k,e in v.items()}
            if not d['category'] or not d['name']:return messagebox.showerror('Required','Category and name are required.',parent=w)
            if sid:self.db.x('UPDATE staff_occupants SET category=?,name=?,gender=?,contact_number=?,address=?,notes=? WHERE id=?',(*[d[k] for k in ['category','name','gender','contact_number','address','notes']],sid));self.db.log(self.user['username'],'UPDATE','Staff',sid)
            else:sid2=self.db.x('INSERT INTO staff_occupants(category,name,gender,contact_number,address,notes) VALUES(?,?,?,?,?,?)',tuple(d[k] for k in ['category','name','gender','contact_number','address','notes']));self.db.log(self.user['username'],'ADD','Staff',sid2)
            w.destroy();self.refresh_all()
        ttk.Button(frm,text='Save',command=save).grid(row=6,column=0,columnspan=2,pady=15)
    def edit_staff(self):
        s=self.stftree.selection();
        if s:self.staff_form(int(s[0]))
    def delete_staff(self):
        s=self.stftree.selection()
        if not s:return
        sid=int(s[0])
        staff=self.db.one('SELECT id,name,category FROM staff_occupants WHERE id=?',(sid,))
        if not staff:return
        if not messagebox.askyesno('Confirm delete',f"Delete {staff['category']} — {staff['name']} and all related accommodation/payment history? This cannot be undone."):
            return
        self.db.c.execute('BEGIN')
        try:
            active=self.db.q("SELECT bed_id FROM allotments WHERE staff_id=? AND status='Active'",(sid,))
            for a in active:
                self.db.x("UPDATE beds SET status='Vacant' WHERE id=?",(a['bed_id'],))
            self.db.exec('DELETE FROM payments WHERE staff_id=?',(sid,))
            self.db.exec('DELETE FROM room_exchanges WHERE staff_id=?',(sid,))
            self.db.exec('DELETE FROM allotments WHERE staff_id=?',(sid,))
            self.db.exec('DELETE FROM staff_occupants WHERE id=?',(sid,))
            self.db.c.commit()
            self.db.log(self.user['username'],'DELETE','Staff',sid,f"category={staff['category']}; name={staff['name']}")
        except Exception as e:
            self.db.c.rollback()
            return messagebox.showerror('Delete failed','No records were deleted.\\n\\n'+str(e))
        self.refresh_all()
        messagebox.showinfo('Deleted',f"{staff['name']} was deleted successfully.")

    # Allotment and exchange
    def allotments(self,f):
        bar=ttk.Frame(f);bar.pack(fill='x');self.btn(bar,'New Allotment',self.allot_form,self.allow_edit()).pack(side='left');self.btn(bar,'Edit Selected',self.edit_allotment,self.allow_edit()).pack(side='left',padx=3);self.btn(bar,'Delete Selected',self.delete_allotment,self.allow_admin()).pack(side='left',padx=3);self.btn(bar,'Print Allotment Receipt',self.print_allotment_selected,True).pack(side='left',padx=3);self.btn(bar,'Exchange Room/Bed',self.exchange_form,self.allow_edit()).pack(side='left',padx=4);self.btn(bar,'Vacation / Gate Pass',self.vacate_selected,self.allow_edit()).pack(side='left');
        self.atree=self.tree(f,['Type','ID / Category','Name','Hostel','Room','Bed','Occupation Date','Vacation','Payment','Status'],18)
    def refresh_allotments(self):
        if not hasattr(self,'atree'):return
        self.atree.delete(*self.atree.get_children())
        for r in self.db.q('''SELECT a.id,CASE WHEN a.student_id IS NOT NULL THEN 'Student' ELSE 'Staff' END typ,COALESCE(s.admission_no,st.category),COALESCE(s.name,st.name),h.name,r.room_number,b.bed_label,a.occupation_date,a.vacation_date,a.payment_status,a.status FROM allotments a LEFT JOIN students s ON s.id=a.student_id LEFT JOIN staff_occupants st ON st.id=a.staff_id JOIN beds b ON b.id=a.bed_id JOIN rooms r ON r.id=b.room_id JOIN hostels h ON h.id=r.hostel_id ORDER BY a.status DESC,a.occupation_date DESC'''):self.atree.insert('', 'end',iid=r['id'],values=tuple(r)[1:])
    def allot_form(self, student_id=None):
        w=tk.Toplevel(self.root);w.title('Room / Bed Allotment');w.geometry('650x560');frm=ttk.Frame(w,padding=18);frm.pack(fill='both',expand=True);v={}
        fields=['Occupant Type','Student / Staff Search','Hostel','Room','Bed','Occupation Date','Security','Hostel Fee','Mess Fee','Payment Status','Remarks'];
        for i,k in enumerate(fields):
            ttk.Label(frm,text=k).grid(row=i,column=0,sticky='w',pady=5)
            if k=='Occupant Type': e=ttk.Combobox(frm,values=['Student','Teacher','Warden','Office Staff','Hostel Management','Cleaning Staff','Cooking Staff','Security','Other Staff','Guest'],state='readonly',width=36);e.set('Student')
            elif k=='Hostel':e=ttk.Combobox(frm,values=['Boys','Girls'],state='readonly',width=36)
            elif k=='Room':e=ttk.Combobox(frm,state='readonly',width=36)
            elif k=='Bed':e=ttk.Combobox(frm,state='readonly',width=36)
            elif k=='Payment Status':e=ttk.Combobox(frm,values=['Paid','Pending','Partial','Not Applicable'],state='readonly',width=36);e.set('Pending')
            else:e=ttk.Entry(frm,width=38)
            e.grid(row=i,column=1);v[k]=e
        v['Occupation Date'].insert(0,date.today().isoformat())
        selected={'id':None,'kind':None}
        if student_id:
            ss=self.db.one('SELECT id,admission_no,name,hostel_fee,mess_fee,security_amount,payment_status FROM students WHERE id=?',(student_id,))
            if ss:
                selected.update(id=ss['id'],kind='student'); v['Student / Staff Search'].insert(0,f"{ss['admission_no']} | {ss['name']}"); v['Security'].insert(0,str(ss['security_amount'] or 0)); v['Hostel Fee'].insert(0,str(ss['hostel_fee'] or 0)); v['Mess Fee'].insert(0,str(ss['mess_fee'] or 0)); v['Payment Status'].set(ss['payment_status'] or 'Pending')
        def search_person():
            q='%'+v['Student / Staff Search'].get().strip()+'%'; typ=v['Occupant Type'].get(); kind='student' if typ=='Student' else 'staff';
            if kind=='student':rows=self.db.q('SELECT id,admission_no,name FROM students WHERE status!=\'Left\' AND (admission_no LIKE ? OR name LIKE ?)',(q,q)); vals=[f"{r['admission_no']} | {r['name']}" for r in rows]
            else:rows=self.db.q('SELECT id,category,name FROM staff_occupants WHERE category=? AND active=1 AND name LIKE ?', (typ,q)); vals=[f"{r['category']} | {r['name']}" for r in rows]
            menu=tk.Toplevel(w);menu.title('Select Person');tt=ttk.Treeview(menu,columns=('a','b'),show='headings');tt.pack(fill='both',expand=True,padx=10,pady=10);tt.heading('a',text='ID/Category');tt.heading('b',text='Name')
            for r in rows:tt.insert('', 'end',iid=r['id'],values=(r['admission_no'] if kind=='student' else r['category'],r['name']))
            def choose():
                s=tt.selection();
                if not s:return
                rid=int(s[0]); rr=self.db.one('SELECT admission_no,name FROM students WHERE id=?',(rid,)) if kind=='student' else self.db.one('SELECT category,name FROM staff_occupants WHERE id=?',(rid,));v['Student / Staff Search'].delete(0,'end');v['Student / Staff Search'].insert(0,f"{rr[0]} | {rr[1]}");selected['id']=rid;selected['kind']=kind;menu.destroy()
            ttk.Button(menu,text='Select',command=choose).pack(pady=5)
        ttk.Button(frm,text='Find',command=search_person).grid(row=1,column=2,padx=5)
        def rooms_change(e=None):
            h=v['Hostel'].get(); rows=self.db.q('SELECT r.id,r.room_number FROM rooms r JOIN hostels h ON h.id=r.hostel_id WHERE h.name=? AND r.status NOT IN (\'Closed\',\'Maintenance\') ORDER BY r.room_number',(h,));v['Room']['values']=[f"{r['id']} | {r['room_number']}" for r in rows];v['Room'].set('');v['Bed'].set('');v['Bed']['values']=[]
        def bed_change(e=None):
            rid=v['Room'].get().split('|')[0].strip() if v['Room'].get() else '';rows=self.db.q("SELECT id,bed_label FROM beds WHERE room_id=? AND status='Vacant' ORDER BY id",(rid,));v['Bed']['values']=[f"{r['id']} | Bed {r['bed_label']}" for r in rows];v['Bed'].set('')
        v['Hostel'].bind('<<ComboboxSelected>>',rooms_change);v['Room'].bind('<<ComboboxSelected>>',bed_change)
        def save():
            if not selected['id']:return messagebox.showerror('Occupant','Find and select the occupant first.',parent=w)
            if not dvalid(v['Occupation Date'].get()):return messagebox.showerror('Date','Occupation date must be YYYY-MM-DD.',parent=w)
            bed=v['Bed'].get().split('|')[0].strip() if v['Bed'].get() else ''; 
            if not bed:return messagebox.showerror('Bed','Select a vacant bed.',parent=w)
            if self.db.one("SELECT 1 FROM allotments WHERE bed_id=? AND status='Active'",(bed,)):return messagebox.showerror('Occupied','That bed is already occupied.',parent=w)
            typ=v['Occupant Type'].get(); student_id=selected['id'] if selected['kind']=='student' else None;staff_id=selected['id'] if selected['kind']=='staff' else None
            aid=self.db.x('INSERT INTO allotments(student_id,staff_id,occupant_category,bed_id,occupation_date,payment_status,security_amount,monthly_fee,mess_fee,status,remarks) VALUES(?,?,?,?,?,?,?,?,?,?,?)',(student_id,staff_id,typ,bed,v['Occupation Date'].get(),v['Payment Status'].get(),float(v['Security'].get() or 0),float(v['Hostel Fee'].get() or 0),float(v['Mess Fee'].get() or 0),'Active',v['Remarks'].get()))
            self.db.exec("UPDATE beds SET status='Occupied' WHERE id=?",(bed,));
            if student_id:self.db.exec("UPDATE students SET status='Active',hostel_fee=?,mess_fee=?,security_amount=?,payment_status=? WHERE id=?",(float(v['Hostel Fee'].get() or 0),float(v['Mess Fee'].get() or 0),float(v['Security'].get() or 0),v['Payment Status'].get(),student_id))
            self.db.log(self.user['username'],'ADD','Allotment',aid);w.destroy();self.refresh_all()
        ttk.Button(frm,text='Save Allotment',command=save).grid(row=len(fields),column=0,columnspan=2,pady=15);ttk.Button(frm,text='Print Allotment Receipt',command=lambda:self.print_selected_allotment(selected)).grid(row=len(fields)+1,column=0,columnspan=2)
    def print_selected_allotment(self,selected):
        if selected['id']: self.print_allotment(selected['id'])
        else: messagebox.showinfo('Receipt','Save the allotment first, then print it from the Allotment tab.')
    def print_allotment_selected(self):
        sel=self.atree.selection() if hasattr(self,'atree') else []
        if sel:self.print_allotment(int(sel[0]))
        else:messagebox.showwarning('Select','Select an allotment first.')
    def print_allotment(self,aid):
        r=self.db.one('''SELECT a.*,COALESCE(s.name,st.name) name,COALESCE(s.admission_no,st.category) ident,COALESCE(s.university,'') university,COALESCE(s.group_no,'') group_no,COALESCE(s.semester,'') semester,h.name hostel,r.room_number,b.bed_label FROM allotments a LEFT JOIN students s ON s.id=a.student_id LEFT JOIN staff_occupants st ON st.id=a.staff_id JOIN beds b ON b.id=a.bed_id JOIN rooms r ON r.id=b.room_id JOIN hostels h ON h.id=r.hostel_id WHERE a.id=?''',(aid,))
        if not r:return
        self.pdf_receipt('HOSTEL ROOM / BED ALLOTMENT RECEIPT',[('Occupant',r['name']),('ID / Category',r['ident']),('University',r['university'] or '-'),('Group',r['group_no'] or '-'),('Semester',r['semester'] or '-'),('Occupant Type',r['occupant_category']),('Hostel',r['hostel']),('Room',r['room_number']),('Bed',r['bed_label']),('Occupation Date',r['occupation_date']),('Hostel Fee',f"{r['monthly_fee'] or 0:.2f}"),('Mess Fee',f"{r['mess_fee'] or 0:.2f}"),('Security',f"{r['security_amount'] or 0:.2f}"),('Payment Status',r['payment_status'])],True)
    def edit_allotment(self):
        sel=self.atree.selection();
        if not sel:return
        aid=int(sel[0]);a=self.db.one('SELECT * FROM allotments WHERE id=?',(aid,));
        if not a:return
        w=tk.Toplevel(self.root);w.title('Edit Allotment');frm=ttk.Frame(w,padding=18);frm.pack();v={};fields=[('Occupation Date','occupation_date'),('Vacation Date','vacation_date'),('Security','security_amount'),('Hostel Fee','monthly_fee'),('Mess Fee','mess_fee'),('Payment Status','payment_status'),('Remarks','remarks')]
        for i,(lab,k) in enumerate(fields):
            ttk.Label(frm,text=lab).grid(row=i,column=0,sticky='w',pady=5); e=ttk.Combobox(frm,values=['Paid','Pending','Partial','Not Applicable'],state='readonly',width=30) if k=='payment_status' else ttk.Entry(frm,width=32);e.grid(row=i,column=1);v[k]=e;e.insert(0,clean(a[k]))
        def save():
            if not dvalid(v['occupation_date'].get()) or not dvalid(v['vacation_date'].get()):return messagebox.showerror('Date','Use YYYY-MM-DD dates.',parent=w)
            self.db.exec('UPDATE allotments SET occupation_date=?,vacation_date=?,security_amount=?,monthly_fee=?,mess_fee=?,payment_status=?,remarks=? WHERE id=?',(v['occupation_date'].get(),v['vacation_date'].get() or None,float(v['security_amount'].get() or 0),float(v['monthly_fee'].get() or 0),float(v['mess_fee'].get() or 0),v['payment_status'].get(),v['remarks'].get(),aid));w.destroy();self.refresh_all()
        ttk.Button(frm,text='Save Changes',command=save).grid(row=len(fields),column=0,columnspan=2,pady=12)
    def delete_allotment(self):
        sel=self.atree.selection();
        if not sel:return
        aid=int(sel[0]);a=self.db.one('SELECT * FROM allotments WHERE id=?',(aid,));
        if a and messagebox.askyesno('Delete','Delete this allotment record and make its bed vacant?'):
            self.db.exec("UPDATE beds SET status='Vacant' WHERE id=?",(a['bed_id'],));self.db.exec('DELETE FROM allotments WHERE id=?',(aid,));self.db.log(self.user['username'],'DELETE','Allotment',aid);self.refresh_all()
    def exchange_form(self):
        s=self.atree.selection() if hasattr(self,'atree') else []
        if not s:return messagebox.showwarning('Select','Select an active allotment first.')
        aid=int(s[0]); a=self.db.one('SELECT * FROM allotments WHERE id=? AND status=\'Active\'',(aid,));
        if not a:return messagebox.showwarning('Select','Only active allotments can be exchanged.')
        w=tk.Toplevel(self.root);w.title('Room / Bed Exchange');w.geometry('560x390');frm=ttk.Frame(w,padding=20);frm.pack(fill='both',expand=True)
        person=self.db.one('SELECT COALESCE(s.name,st.name) name FROM allotments a LEFT JOIN students s ON s.id=a.student_id LEFT JOIN staff_occupants st ON st.id=a.staff_id WHERE a.id=?',(aid,)); ttk.Label(frm,text='Occupant: '+person['name']).pack(anchor='w',pady=5)
        ttk.Label(frm,text='New Hostel').pack(anchor='w');host=ttk.Combobox(frm,values=['Boys','Girls'],state='readonly');host.pack(fill='x');ttk.Label(frm,text='New Room').pack(anchor='w');room=ttk.Combobox(frm,state='readonly');room.pack(fill='x');ttk.Label(frm,text='New Bed').pack(anchor='w');bed=ttk.Combobox(frm,state='readonly');bed.pack(fill='x');ttk.Label(frm,text='Exchange Date').pack(anchor='w');dt=ttk.Entry(frm);dt.insert(0,date.today().isoformat());dt.pack(fill='x');ttk.Label(frm,text='Reason').pack(anchor='w');reason=ttk.Entry(frm);reason.pack(fill='x')
        def rh(e=None):
            rows=self.db.q('SELECT r.id,r.room_number FROM rooms r JOIN hostels h ON h.id=r.hostel_id WHERE h.name=? AND r.status NOT IN (\'Closed\',\'Maintenance\')',(host.get(),));room['values']=[f"{r['id']} | {r['room_number']}" for r in rows]
        def bh(e=None):
            rid=room.get().split('|')[0].strip();rows=self.db.q("SELECT id,bed_label FROM beds WHERE room_id=? AND status='Vacant'",(rid,));bed['values']=[f"{r['id']} | Bed {r['bed_label']}" for r in rows]
        host.bind('<<ComboboxSelected>>',rh);room.bind('<<ComboboxSelected>>',bh)
        def save():
            if not dvalid(dt.get()) or not bed.get():return messagebox.showerror('Invalid','Choose a vacant bed and valid date.',parent=w)
            nb=int(bed.get().split('|')[0]); oldbed=a['bed_id'];self.db.exec('UPDATE allotments SET bed_id=? WHERE id=?',(nb,aid));self.db.exec("UPDATE beds SET status='Vacant' WHERE id=?",(oldbed,));self.db.exec("UPDATE beds SET status='Occupied' WHERE id=?",(nb,));eid=self.db.x('INSERT INTO room_exchanges(student_id,staff_id,from_bed_id,to_bed_id,exchange_date,reason) VALUES(?,?,?,?,?,?)',(a['student_id'],a['staff_id'],oldbed,nb,dt.get(),reason.get()));self.db.log(self.user['username'],'EXCHANGE','Allotment',aid);w.destroy();self.refresh_all();self.print_exchange(eid)
        ttk.Button(frm,text='Exchange & Print Receipt',command=save).pack(pady=15)

    # Payments
    def payments(self,f):
        bar=ttk.Frame(f);bar.pack(fill='x');self.btn(bar,'Add Payment',self.pay_form,self.allow_edit()).pack(side='left');self.btn(bar,'Edit Selected',self.edit_payment,self.allow_edit()).pack(side='left',padx=4);self.btn(bar,'Delete Selected',self.delete_payment,self.allow_admin()).pack(side='left');self.ptree=self.tree(f,['Receipt','Payer','Fee Type','Month','Amount','Discount','Balance','Status','Payment Date'],18)
    def refresh_payments(self):
        if not hasattr(self,'ptree'):return
        self.ptree.delete(*self.ptree.get_children())
        for r in self.db.q('''SELECT p.id,p.receipt_no,COALESCE(s.name,st.name),p.fee_type,p.fee_month,p.amount,p.discount,p.balance,p.status,p.payment_date FROM payments p LEFT JOIN students s ON s.id=p.student_id LEFT JOIN staff_occupants st ON st.id=p.staff_id ORDER BY p.payment_date DESC,p.id DESC'''):self.ptree.insert('', 'end',iid=r['id'],values=tuple(r)[1:])
        self.ptree.bind('<Double-1>',lambda e:self.print_payment_selected())
    def pay_form(self):
        w=tk.Toplevel(self.root);w.title('Fee / Payment Record');w.geometry('610x520');frm=ttk.Frame(w,padding=18);frm.pack(fill='both',expand=True);v={};fields=['Student ID / Staff Name','Fee Month (YYYY-MM)','Fee Type','Amount','Discount','Balance','Status','Payment Date','Remarks']
        for i,k in enumerate(fields):
            ttk.Label(frm,text=k).grid(row=i,column=0,sticky='w',pady=5); vals=['Hostel Fee','Mess Fee','Security','Other'] if k=='Fee Type' else ['Paid','Pending','Partial','Not Applicable'] if k=='Status' else None;e=ttk.Combobox(frm,values=vals,state='readonly',width=34) if vals else ttk.Entry(frm,width=36);e.grid(row=i,column=1);v[k]=e
        v['Payment Date'].insert(0,date.today().isoformat());v['Status'].set('Paid')
        def save():
            ident=v['Student ID / Staff Name'].get().strip();s=self.db.one('SELECT id FROM students WHERE admission_no=?',(ident,));st=None if s else self.db.one('SELECT id FROM staff_occupants WHERE name LIKE ?',(ident,));
            if not s and not st:return messagebox.showerror('Payer','Enter an existing student ID or exact staff name.',parent=w)
            if not re.fullmatch(r'\d{4}-\d{2}',v['Fee Month (YYYY-MM)'].get()):return messagebox.showerror('Month','Use YYYY-MM.',parent=w)
            try:amt=float(v['Amount'].get() or 0);disc=float(v['Discount'].get() or 0);bal=float(v['Balance'].get() or 0)
            except:return messagebox.showerror('Amount','Amount/discount/balance must be numbers.',parent=w)
            rec='RCPT-'+datetime.now().strftime('%Y%m%d%H%M%S')
            pid=self.db.x('INSERT INTO payments(student_id,staff_id,payment_date,fee_month,fee_type,amount,discount,balance,status,receipt_no,remarks) VALUES(?,?,?,?,?,?,?,?,?,?,?)',((s['id'] if s else None),(st['id'] if st else None),v['Payment Date'].get(),v['Fee Month (YYYY-MM)'].get(),v['Fee Type'].get(),amt,disc,bal,v['Status'].get(),rec,v['Remarks'].get()));
            if s:self.db.exec("UPDATE students SET payment_status=?,fee_balance=? WHERE id=?",(v['Status'].get(),bal,s['id']))
            self.db.log(self.user['username'],'ADD','Payment',pid);w.destroy();self.refresh_all();self.print_payment(pid)
        ttk.Button(frm,text='Save & Print Receipt',command=save).grid(row=len(fields),column=0,columnspan=2,pady=15)
    def print_payment_selected(self):
        s=self.ptree.selection();
        if s:self.print_payment(int(s[0]))
    def edit_payment(self):
        s=self.ptree.selection();
        if not s:return
        pid=int(s[0]);a=self.db.one('SELECT * FROM payments WHERE id=?',(pid,));
        if not a:return
        w=tk.Toplevel(self.root);w.title('Edit Payment');frm=ttk.Frame(w,padding=18);frm.pack();v={};fields=[('Payment Date','payment_date'),('Fee Month','fee_month'),('Fee Type','fee_type'),('Amount','amount'),('Discount','discount'),('Balance','balance'),('Status','status'),('Remarks','remarks')]
        for i,(lab,k) in enumerate(fields):
            ttk.Label(frm,text=lab).grid(row=i,column=0,sticky='w',pady=5); vals=['Hostel Fee','Mess Fee','Security','Other'] if k=='fee_type' else ['Paid','Pending','Partial','Not Applicable'] if k=='status' else None;e=ttk.Combobox(frm,values=vals,state='readonly',width=30) if vals else ttk.Entry(frm,width=32);e.grid(row=i,column=1);v[k]=e;e.insert(0,clean(a[k]))
        def save():
            if not dvalid(v['payment_date'].get()) or not re.fullmatch(r'\d{4}-\d{2}',v['fee_month'].get()):return messagebox.showerror('Date','Use YYYY-MM-DD and YYYY-MM.',parent=w)
            self.db.exec('UPDATE payments SET payment_date=?,fee_month=?,fee_type=?,amount=?,discount=?,balance=?,status=?,remarks=? WHERE id=?',(v['payment_date'].get(),v['fee_month'].get(),v['fee_type'].get(),float(v['amount'].get() or 0),float(v['discount'].get() or 0),float(v['balance'].get() or 0),v['status'].get(),v['remarks'].get(),pid));
            if a['student_id']:
                agg=self.db.one("SELECT COALESCE(SUM(balance),0) bal, CASE WHEN SUM(CASE WHEN status='Pending' THEN 1 ELSE 0 END)>0 THEN 'Pending' WHEN SUM(CASE WHEN status='Partial' THEN 1 ELSE 0 END)>0 THEN 'Partial' ELSE 'Paid' END ps FROM payments WHERE student_id=?",(a['student_id'],))
                self.db.exec("UPDATE students SET payment_status=?,fee_balance=? WHERE id=?",(agg['ps'],agg['bal'],a['student_id']))
            w.destroy();self.refresh_all()
        ttk.Button(frm,text='Save Changes',command=save).grid(row=len(fields),column=0,columnspan=2,pady=12)
    def delete_payment(self):
        s=self.ptree.selection();
        if s and messagebox.askyesno('Delete','Delete selected payment/receipt?'):
            pid=int(s[0]); pay=self.db.one('SELECT student_id FROM payments WHERE id=?',(pid,)); self.db.exec('DELETE FROM payments WHERE id=?',(pid,))
            if pay and pay['student_id']:
                agg=self.db.one("SELECT COALESCE(SUM(balance),0) bal, CASE WHEN SUM(CASE WHEN status='Pending' THEN 1 ELSE 0 END)>0 THEN 'Pending' WHEN SUM(CASE WHEN status='Partial' THEN 1 ELSE 0 END)>0 THEN 'Partial' ELSE 'Paid' END ps FROM payments WHERE student_id=?",(pay['student_id'],))
                self.db.exec("UPDATE students SET payment_status=?,fee_balance=? WHERE id=?",(agg['ps'],agg['bal'],pay['student_id']))
            self.refresh_all()
    def print_payment(self,pid):
        r=self.db.one('''SELECT p.*,COALESCE(s.name,st.name) payer,COALESCE(s.admission_no,st.category) ident FROM payments p LEFT JOIN students s ON s.id=p.student_id LEFT JOIN staff_occupants st ON st.id=p.staff_id WHERE p.id=?''',(pid,));self.pdf_receipt('HOSTEL FEE PAYMENT RECEIPT', [('Receipt No',r['receipt_no']),('Payer',r['payer']),('ID / Category',r['ident']),('Fee Type',r['fee_type']),('Month',r['fee_month']),('Amount',f"{r['amount']:.2f}"),('Discount',f"{r['discount']:.2f}"),('Balance',f"{r['balance']:.2f}"),('Status',r['status']),('Payment Date',r['payment_date'])], signature=True)

    # Hostel left / vacation
    def left_tab(self,f):
        bar=ttk.Frame(f);bar.pack(fill='x');ttk.Button(bar,text='Vacate Selected / Gate Pass',command=self.vacate_selected,state=('normal' if self.allow_edit() else 'disabled')).pack(side='left');self.btn(bar,'Print Selected Gate Pass',self.print_left_selected,True).pack(side='left',padx=4);self.btn(bar,'Delete Selected History',self.delete_left,self.allow_admin()).pack(side='left',padx=4);self.ltree=self.tree(f,['Student ID','Name','University','Group','Semester','Hostel','Room','Bed','Occupation Date','Vacation Date','Payment','Reason'],18)
    def refresh_left(self):
        if not hasattr(self,'ltree'):return
        self.ltree.delete(*self.ltree.get_children())
        for r in self.db.q("SELECT lh.id,lh.admission_no,lh.student_name,lh.university,COALESCE(s.group_no,''),COALESCE(s.semester,''),lh.hostel,lh.room,lh.bed,lh.occupation_date,lh.vacation_date,lh.payment_status,lh.reason FROM left_hostel lh LEFT JOIN students s ON s.id=lh.student_id ORDER BY lh.vacation_date DESC"):self.ltree.insert('', 'end',iid=r['id'],values=tuple(r)[1:])
    def print_left_selected(self):
        sel=self.ltree.selection() if hasattr(self,'ltree') else []
        if not sel:return messagebox.showwarning('Select','Select a hostel-left record first.')
        r=self.db.one("SELECT lh.*,COALESCE(s.group_no,'') group_no,COALESCE(s.semester,'') semester FROM left_hostel lh LEFT JOIN students s ON s.id=lh.student_id WHERE lh.id=?",(int(sel[0]),))
        if r:self.pdf_receipt('HOSTEL LEFT / VACATION RECEIPT',[('Student',r['student_name']),('Student ID',r['admission_no'] or '-'),('University',r['university'] or '-'),('Group',r['group_no'] or '-'),('Semester',r['semester'] or '-'),('Hostel',r['hostel'] or '-'),('Room',r['room'] or '-'),('Bed',r['bed'] or '-'),('Occupation Date',r['occupation_date'] or '-'),('Vacation Date',r['vacation_date'] or '-'),('Payment Status',r['payment_status'] or '-'),('Reason',r['reason'] or '-')],True)

    def delete_left(self):
        s=self.ltree.selection();
        if s and messagebox.askyesno('Delete','Delete selected hostel-left history?'):self.db.exec('DELETE FROM left_hostel WHERE id=?',(int(s[0]),));self.refresh_all()
    def vacate_selected(self):
        sid=None
        if hasattr(self,'str') and self.str.selection():sid=int(self.str.selection()[0])
        elif hasattr(self,'global_tree') and self.global_tree.selection():
            v=self.global_tree.item(self.global_tree.selection()[0],'values');
            if v and v[0]=='Student': rr=self.db.one('SELECT id FROM students WHERE admission_no=?',(v[1],));sid=rr['id'] if rr else None
        if not sid:return messagebox.showwarning('Select','Select an active student first.')
        self.vacate(sid)
    def vacate(self,sid,reason=None,confirm=True):
        a=self.db.one('''SELECT a.*,s.admission_no,s.name,s.university,s.group_no,s.semester,h.name hostel,r.room_number,b.bed_label FROM allotments a JOIN students s ON s.id=a.student_id JOIN beds b ON b.id=a.bed_id JOIN rooms r ON r.id=b.room_id JOIN hostels h ON h.id=r.hostel_id WHERE a.student_id=? AND a.status='Active' ''',(sid,))
        if not a:return messagebox.showinfo('Not allotted','Student has no active hostel allotment.')
        w=None
        if confirm:
            w=tk.Toplevel(self.root);w.title('Vacation / Gate Pass');w.geometry('500x300');frm=ttk.Frame(w,padding=20);frm.pack(fill='both',expand=True);ttk.Label(frm,text=f"{a['name']} • Room {a['room_number']} • Bed {a['bed_label']}",font=('Segoe UI',13,'bold')).pack(anchor='w');ttk.Label(frm,text='Vacation Date (YYYY-MM-DD)').pack(anchor='w',pady=(15,3));dt=ttk.Entry(frm);dt.insert(0,date.today().isoformat());dt.pack(fill='x');ttk.Label(frm,text='Reason').pack(anchor='w',pady=(10,3));rs=ttk.Entry(frm);rs.pack(fill='x')
            def go():
                if not dvalid(dt.get()):return messagebox.showerror('Date','Invalid date.',parent=w)
                self._complete_vacate(a,dt.get(),rs.get());w.destroy();self.refresh_all();self.print_gatepass(a,dt.get(),rs.get())
            ttk.Button(frm,text='Vacate & Print Gate Pass',command=go).pack(pady=18)
        else:self._complete_vacate(a,date.today().isoformat(),reason or 'System deletion')
    def _complete_vacate(self,a,dt,reason):
        self.db.exec("UPDATE allotments SET vacation_date=?,status='Vacated' WHERE id=?",(dt,a['id']));self.db.exec("UPDATE beds SET status='Vacant' WHERE id=?",(a['bed_id'],));self.db.exec("UPDATE students SET status='Left',updated_at=CURRENT_TIMESTAMP WHERE id=?",(a['student_id'],));self.db.x('INSERT INTO left_hostel(student_id,student_name,admission_no,university,hostel,room,bed,occupation_date,vacation_date,reason,payment_status,remarks) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)',(a['student_id'],a['name'],a['admission_no'],a['university'],a['hostel'],a['room_number'],a['bed_label'],a['occupation_date'],dt,reason,a['payment_status'],a['remarks']));self.db.log(self.user['username'],'VACATE','Student',a['student_id'])

    # ID cards
    def idcards(self,f):
        bar=ttk.Frame(f);bar.pack(fill='x');ttk.Button(bar,text='Generate / Print Selected ID Card',command=self.id_from_student).pack(side='left');self.btn(bar,'Delete Selected Card',self.delete_card,self.allow_admin()).pack(side='left',padx=5);self.itree=self.tree(f,['Card No','Student ID','Name','Hostel','Room','Bed','Issue Date','Expiry'],18)
    def refresh_idcards(self):
        if not hasattr(self,'itree'):return
        self.itree.delete(*self.itree.get_children())
        for r in self.db.q('''SELECT c.id,c.card_no,s.admission_no,s.name,h.name,r.room_number,b.bed_label,c.issue_date,c.expiry_date FROM id_cards c JOIN students s ON s.id=c.student_id LEFT JOIN allotments a ON a.student_id=s.id AND a.status='Active' LEFT JOIN beds b ON b.id=a.bed_id LEFT JOIN rooms r ON r.id=b.room_id LEFT JOIN hostels h ON h.id=r.hostel_id ORDER BY c.id DESC'''):self.itree.insert('', 'end',iid=r['id'],values=tuple(r)[1:])
    def delete_card(self):
        s=self.itree.selection();
        if s and messagebox.askyesno('Delete','Delete selected ID card?'):self.db.exec('DELETE FROM id_cards WHERE id=?',(int(s[0]),))
    def id_from_student(self):
        sid=None
        if hasattr(self,'str') and self.str.selection():sid=int(self.str.selection()[0])
        elif hasattr(self,'itree') and self.itree.selection():
            rr=self.db.one('SELECT student_id FROM id_cards WHERE id=?',(int(self.itree.selection()[0]),));sid=rr['student_id'] if rr else None
        if not sid:return messagebox.showwarning('Select','Select a student first.')
        c=self.db.one('SELECT * FROM id_cards WHERE student_id=?',(sid,));
        if not c:
            no='HOSTEL-'+str(self.db.one('SELECT admission_no FROM students WHERE id=?',(sid,))['admission_no']);self.db.x('INSERT INTO id_cards(student_id,card_no,issue_date) VALUES(?,?,?)',(sid,no,date.today().isoformat()))
        self.card_pdf(sid)
    def card_pdf(self,sid):
        r=self.db.one('''SELECT s.*,c.card_no,c.issue_date,c.expiry_date,h.name hostel,rm.room_number,b.bed_label FROM students s JOIN id_cards c ON c.student_id=s.id LEFT JOIN allotments a ON a.student_id=s.id AND a.status='Active' LEFT JOIN beds b ON b.id=a.bed_id LEFT JOIN rooms rm ON rm.id=b.room_id LEFT JOIN hostels h ON h.id=rm.hostel_id WHERE s.id=?''',(sid,));
        self.pdf_receipt('HOSTEL STUDENT ID CARD',[('Institution',self.setting('institution_name')),('Card No',r['card_no']),('Student',r['name']),('Student ID',r['admission_no']),('University',r['university'] or '-'),('Hostel',r['hostel'] or '-'),('Room',r['room_number'] or '-'),('Bed',r['bed_label'] or '-'),('Issue Date',r['issue_date'])],signature=True)

    # Reports
    def reports(self,f):
        ttk.Label(f,text='Reports & Printing',style='Title.TLabel').pack(anchor='w'); top=ttk.Frame(f);top.pack(fill='x',pady=8)
        ttk.Label(top,text='Period').pack(side='left');self.rperiod=ttk.Combobox(top,values=['Weekly','Monthly','Yearly'],state='readonly');self.rperiod.set('Monthly');self.rperiod.pack(side='left',padx=5);ttk.Label(top,text='Date (YYYY-MM-DD)').pack(side='left',padx=5);self.rdate=ttk.Entry(top,width=14);self.rdate.insert(0,date.today().isoformat());self.rdate.pack(side='left');ttk.Button(top,text='Generate',command=self.refresh_reports).pack(side='left',padx=5);ttk.Button(top,text='Print / Save PDF',command=self.print_report).pack(side='left');ttk.Button(top,text='Export Excel',command=self.export_report).pack(side='left',padx=5)
        quick=ttk.Frame(f);quick.pack(fill='x',pady=5)
        for txt,cmd in [('Student List',lambda:self.report_custom('student')),('Room Occupancy',lambda:self.report_custom('rooms')),('Allotment Register',lambda:self.report_custom('allot')),('Fee Register',lambda:self.report_custom('fees')),('Hostel Left Register',lambda:self.report_custom('left'))]:ttk.Button(quick,text=txt,command=cmd).pack(side='left',padx=3)
        self.reptree=self.tree(f,['Report','Value'],18)
    def report_rows(self):
        p=self.rperiod.get(); d=self.rdate.get();
        if not dvalid(d): raise ValueError('Date must be YYYY-MM-DD')
        if p=='Weekly': start=(datetime.strptime(d,'%Y-%m-%d').date()-timedelta(days=6)).isoformat(); end=d
        elif p=='Monthly': start=d[:7]+'-01'; end=d
        else:start=d[:4]+'-01-01';end=d
        return start,end
    def refresh_reports(self):
        if not hasattr(self,'reptree'):return
        try:start,end=self.report_rows()
        except Exception as e:return messagebox.showerror('Report',str(e))
        self.reptree.delete(*self.reptree.get_children()); data=[('Period',f'{start} to {end}'),('Students Living',self.db.one("SELECT count(*) n FROM allotments WHERE status='Active' AND student_id IS NOT NULL")['n']),('Beds Occupied',self.db.one("SELECT count(*) n FROM beds WHERE status='Occupied'")['n']),('Beds Vacant',self.db.one("SELECT count(*) n FROM beds WHERE status='Vacant'")['n']),('New Allotments',self.db.one("SELECT count(*) n FROM allotments WHERE occupation_date BETWEEN ? AND ?",(start,end))['n']),('Hostel Vacations',self.db.one("SELECT count(*) n FROM left_hostel WHERE vacation_date BETWEEN ? AND ?",(start,end))['n']),('Fees Charged',self.db.one("SELECT coalesce(sum(hostel_fee+mess_fee+security_amount),0) n FROM students")['n']),('Fees Paid',self.db.one("SELECT coalesce(sum((hostel_fee+mess_fee+security_amount)-fee_balance),0) n FROM students")['n']),('Pending Fee Balance',self.db.one("SELECT coalesce(sum(fee_balance),0) n FROM students")['n'])]
        for x in data:self.reptree.insert('', 'end',values=x)
    def report_custom(self,kind):
        self.reptree.delete(*self.reptree.get_children())
        if kind=='student':
            headers=['Student ID','Name','University','Group','Semester','Hostel','Room','Bed','Occupation Date','Hostel Fee','Mess Fee','Paid','Balance','Fee Status','Status']
            rows=self.db.q("""SELECT s.admission_no,s.name,COALESCE(s.university,''),COALESCE(s.group_no,''),COALESCE(s.semester,''),COALESCE(h.name,''),COALESCE(r.room_number,''),COALESCE(b.bed_label,''),COALESCE(a.occupation_date,''),COALESCE(a.monthly_fee,s.hostel_fee,0),COALESCE(a.mess_fee,s.mess_fee,0),MAX(0,(s.hostel_fee+s.mess_fee+s.security_amount)-s.fee_balance),s.fee_balance,s.payment_status,CASE WHEN a.id IS NOT NULL THEN 'Living' ELSE s.status END FROM students s LEFT JOIN allotments a ON a.student_id=s.id AND a.status='Active' LEFT JOIN beds b ON b.id=a.bed_id LEFT JOIN rooms r ON r.id=b.room_id LEFT JOIN hostels h ON h.id=r.hostel_id ORDER BY s.name""")
        elif kind=='allot':
            headers=['Type','ID / Category','Name','University','Group','Semester','Hostel','Room','Bed','Occupation Date','Hostel Fee','Mess Fee','Security','Payment','Balance','Status']
            rows=self.db.q("""SELECT CASE WHEN a.student_id IS NOT NULL THEN 'Student' ELSE a.occupant_category END,COALESCE(s.admission_no,st.category),COALESCE(s.name,st.name),COALESCE(s.university,''),COALESCE(s.group_no,''),COALESCE(s.semester,''),h.name,r.room_number,b.bed_label,a.occupation_date,a.monthly_fee,a.mess_fee,a.security_amount,a.payment_status,COALESCE(s.fee_balance,0),a.status FROM allotments a LEFT JOIN students s ON s.id=a.student_id LEFT JOIN staff_occupants st ON st.id=a.staff_id JOIN beds b ON b.id=a.bed_id JOIN rooms r ON r.id=b.room_id JOIN hostels h ON h.id=r.hostel_id ORDER BY a.occupation_date DESC""")
        elif kind=='staff':
            headers=['Category','Name','Gender','Contact','Address','Hostel','Room','Bed','Occupation Date','Status']
            rows=self.db.q("""SELECT st.category,st.name,COALESCE(st.gender,''),COALESCE(st.contact_number,''),COALESCE(st.address,''),COALESCE(h.name,''),COALESCE(r.room_number,''),COALESCE(b.bed_label,''),COALESCE(a.occupation_date,''),CASE WHEN a.id IS NOT NULL THEN 'Living' ELSE CASE WHEN st.active=1 THEN 'Available' ELSE 'Inactive' END END FROM staff_occupants st LEFT JOIN allotments a ON a.staff_id=st.id AND a.status='Active' LEFT JOIN beds b ON b.id=a.bed_id LEFT JOIN rooms r ON r.id=b.room_id LEFT JOIN hostels h ON h.id=r.hostel_id ORDER BY st.category,st.name""")
        elif kind=='left':
            headers=['Student ID','Name','University','Group','Semester','Hostel','Room','Bed','Occupation Date','Vacation Date','Hostel Fee','Mess Fee','Balance','Payment','Reason']
            rows=self.db.q("""SELECT lh.admission_no,lh.student_name,COALESCE(lh.university,''),COALESCE(s.group_no,''),COALESCE(s.semester,''),COALESCE(lh.hostel,''),COALESCE(lh.room,''),COALESCE(lh.bed,''),COALESCE(lh.occupation_date,''),COALESCE(lh.vacation_date,''),COALESCE(s.hostel_fee,0),COALESCE(s.mess_fee,0),COALESCE(s.fee_balance,0),COALESCE(lh.payment_status,''),COALESCE(lh.reason,'') FROM left_hostel lh LEFT JOIN students s ON s.id=lh.student_id ORDER BY lh.vacation_date DESC""")
        elif kind=='fees':
            headers=['Student ID','Name','University','Group','Semester','Hostel Fee','Mess Fee','Security','Paid','Balance','Payment Status']
            rows=self.db.q("""SELECT admission_no,name,COALESCE(university,''),COALESCE(group_no,''),COALESCE(semester,''),hostel_fee,mess_fee,security_amount,MAX(0,(hostel_fee+mess_fee+security_amount)-fee_balance),fee_balance,payment_status FROM students ORDER BY name""")
        elif kind=='rooms':
            headers=['Hostel','Room','Beds','Occupied','Vacant','Student Occupied','Staff Occupied','Room Status']
            rows=self.db.q("""SELECT h.name,r.room_number,count(b.id),sum(CASE WHEN b.status='Occupied' THEN 1 ELSE 0 END),sum(CASE WHEN b.status='Vacant' THEN 1 ELSE 0 END),sum(CASE WHEN b.status='Occupied' AND EXISTS(SELECT 1 FROM allotments aa WHERE aa.bed_id=b.id AND aa.status='Active' AND aa.student_id IS NOT NULL) THEN 1 ELSE 0 END),sum(CASE WHEN b.status='Occupied' AND EXISTS(SELECT 1 FROM allotments aa WHERE aa.bed_id=b.id AND aa.status='Active' AND aa.staff_id IS NOT NULL) THEN 1 ELSE 0 END),r.status FROM rooms r JOIN hostels h ON h.id=r.hostel_id LEFT JOIN beds b ON b.room_id=r.id GROUP BY r.id ORDER BY h.name,r.room_number""")
        else:
            headers=['Student ID','Name','University','Group','Semester','Hostel','Room','Bed','Hostel Fee','Mess Fee','Security','Paid','Balance','Payment Status']
            rows=self.db.q("""SELECT s.admission_no,s.name,COALESCE(s.university,''),COALESCE(s.group_no,''),COALESCE(s.semester,''),COALESCE(h.name,''),COALESCE(r.room_number,''),COALESCE(b.bed_label,''),s.hostel_fee,s.mess_fee,s.security_amount,MAX(0,(s.hostel_fee+s.mess_fee+s.security_amount)-s.fee_balance),s.fee_balance,s.payment_status FROM students s LEFT JOIN allotments a ON a.student_id=s.id AND a.status='Active' LEFT JOIN beds b ON b.id=a.bed_id LEFT JOIN rooms r ON r.id=b.room_id LEFT JOIN hostels h ON h.id=r.hostel_id GROUP BY s.id ORDER BY s.name""")
        self.report_headers=headers
        for row in rows:self.reptree.insert('', 'end',values=tuple(row))
    def print_report(self):
        if not self.reptree.get_children():self.refresh_reports()
        data=[getattr(self,'report_headers',['Report','Value'])]+[self.reptree.item(i,'values') for i in self.reptree.get_children()];self.pdf_table('HOSTEL REPORT',data)
    def export_report(self):
        if not openpyxl:return messagebox.showerror('Excel','openpyxl is not installed.')
        p=filedialog.asksaveasfilename(defaultextension='.xlsx',filetypes=[('Excel','*.xlsx')]);
        if not p:return
        wb=openpyxl.Workbook();ws=wb.active;rows=[self.reptree.item(i,'values') for i in self.reptree.get_children()];
        for r in rows:ws.append(list(r))
        wb.save(p);messagebox.showinfo('Saved',p)

    # IO
    def io(self,f):
        ttk.Label(f,text='Import / Export / Backup',style='Title.TLabel').pack(anchor='w');
        for txt,cmd in [('Import Source Excel',self.import_source),('Export Students Excel',lambda:self.export_table('students')),('Export Rooms/Beds Excel',lambda:self.export_table('rooms')),('Export Staff Data Excel',lambda:self.export_table('staff')),('Export Hostel Left Excel',lambda:self.export_table('left')),('Backup Database',self.backup),('Restore Database',self.restore),('Reset Accommodation — Delete All Rooms & Beds',self.reset_accommodation),('Start New Session — Delete ALL Data Except Users & Settings',self.new_session)]: self.btn(f,txt,cmd,self.allow_edit() if 'Export' not in txt and 'Backup' not in txt and 'Restore' not in txt else True).pack(anchor='w',pady=6)
        ttk.Label(f,text='Original workbook is retained in data/source_reference.xlsx. Imports log duplicate/invalid information instead of silently discarding it.',wraplength=900).pack(anchor='w',pady=15)
    def import_source(self):
        if not openpyxl:return messagebox.showerror('Excel','Install openpyxl first.')
        p=filedialog.askopenfilename(filetypes=[('Excel Workbook','*.xlsx')]);
        if p:
            try:self.import_excel(p);self.refresh_all();messagebox.showinfo('Import','Import completed. Review Data_Quality/Import issues in the database if needed.')
            except Exception as e:messagebox.showerror('Import error',str(e))
    def import_excel(self,p,silent=False):
        wb=openpyxl.load_workbook(p,data_only=True);batch=datetime.now().strftime('%Y%m%d%H%M%S')
        if 'Students' in wb.sheetnames:
            ws=wb['Students']
            for i,row in enumerate(ws.iter_rows(min_row=2,values_only=True),2):
                adm=clean(row[1]); name=clean(row[2]); typ=clean(row[10]);
                if not adm or not name:
                    self.db.x('INSERT INTO import_issues(batch,source_sheet,source_row,field,value,problem,recommended_action) VALUES(?,?,?,?,?,?,?)',(batch,'Students',i,'Admission_No',adm,'Missing student ID or name','Verify source record'));continue
                if typ.lower().startswith('non-student'):
                    # Preserve non-student source rows as staff/facility occupants.
                    cat='Facility';
                    self.db.x('INSERT OR IGNORE INTO staff_occupants(name,category,gender,notes) VALUES(?,?,?,?)',(name,cat,clean(row[9]),'Imported from Students sheet'))
                    continue
                if self.db.one('SELECT 1 FROM students WHERE admission_no=?',(adm,)): continue
                sid=self.db.x('INSERT INTO students(admission_no,name,university,semester,group_no,gender,status,source_sheet,source_row,raw_json) VALUES(?,?,?,?,?,?,?,?,?,?)',(adm,name,clean(row[3]),clean(row[4]),clean(row[6]),clean(row[9]),'Active','Students',i,json.dumps(list(row),default=str)))
                hostel=clean(row[7]); room=clean(row[8]);
                if hostel and room:
                    h=self.db.one('SELECT id FROM hostels WHERE name=?',(hostel,));
                    if h:
                        rr=self.db.one('SELECT id FROM rooms WHERE hostel_id=? AND room_number=?',(h['id'],room));
                        if not rr: rr={'id':self.db.x('INSERT INTO rooms(hostel_id,room_number,capacity,status) VALUES(?,?,?,?)',(h['id'],room,1,'Occupied'))}
                        bed=self.db.one("SELECT b.id FROM beds b WHERE b.room_id=? AND b.status='Vacant' ORDER BY b.id LIMIT 1",(rr['id'],));
                        if not bed:
                            bid=self.db.x('INSERT INTO beds(room_id,bed_label,status) VALUES(?,?,?)',(rr['id'],str(self.db.one('SELECT count(*) n FROM beds WHERE room_id=?',(rr['id'],))['n']+1),'Occupied'))
                        else:bid=bed['id']
                        if not self.db.one("SELECT 1 FROM allotments WHERE student_id=? AND status='Active'",(sid,)): self.db.x('INSERT INTO allotments(student_id,occupant_category,bed_id,occupation_date,payment_status,status,remarks) VALUES(?,?,?,?,?,?,?)',(sid,'Student',bid,date.today().isoformat(),'Pending','Active','Imported from source Excel'))
                        self.db.exec("UPDATE beds SET status='Occupied' WHERE id=?",(bid,))
        if 'Rooms' in wb.sheetnames:
            ws=wb['Rooms']
            for i,row in enumerate(ws.iter_rows(min_row=2,values_only=True),2):
                hostel=clean(row[1]);room=clean(row[2]);
                if not hostel or not room:continue
                h=self.db.one('SELECT id FROM hostels WHERE name=?',(hostel,));
                if not h:continue
                cap=int(row[4] or 0)
                rr=self.db.one('SELECT id FROM rooms WHERE hostel_id=? AND room_number=?',(h['id'],room));
                if rr:self.db.exec('UPDATE rooms SET floor=?,capacity=? WHERE id=?',(clean(row[3]),cap,rr['id']));self.sync_beds(rr['id'],cap)
                else:rid=self.db.x('INSERT INTO rooms(hostel_id,room_number,floor,capacity,status) VALUES(?,?,?,?,?)',(h['id'],room,clean(row[3]),cap,'Vacant'));self.sync_beds(rid,cap)
        if 'Student_Leave' in wb.sheetnames:
            ws=wb['Student_Leave']
            for i,row in enumerate(ws.iter_rows(min_row=2,values_only=True),2):
                adm=clean(row[1]);dt=iso(row[7]);
                if not dt:continue
                s=self.db.one('SELECT id,name FROM students WHERE admission_no=?',(adm,)) if adm else None
                self.db.x('INSERT INTO left_hostel(student_id,student_name,admission_no,university,room,vacation_date,payment_status,remarks) VALUES(?,?,?,?,?,?,?,?)',((s['id'] if s else None),(s['name'] if s else clean(row[2])),adm,clean(row[3]),clean(row[6]),dt,clean(row[8]),clean(row[9])))
        if p != SRC:
            try: shutil.copy2(p,SRC)
            except Exception: pass
        self.db.commit()
        self.db.exec("INSERT INTO settings(key,value) VALUES('initial_import_done','1') ON CONFLICT(key) DO UPDATE SET value='1'")
    def export_table(self,kind):
        if not openpyxl:return
        p=filedialog.asksaveasfilename(defaultextension='.xlsx',filetypes=[('Excel','*.xlsx')]);
        if not p:return
        wb=openpyxl.Workbook();ws=wb.active
        if kind=='students':
            rows=self.db.q('SELECT * FROM students ORDER BY name')
        elif kind=='fees':
            rows=self.db.q("""SELECT admission_no,name,COALESCE(university,'') university,COALESCE(group_no,'') group_no,
                COALESCE(semester,'') semester,hostel_fee,mess_fee,security_amount,
                MAX(0,(hostel_fee+mess_fee+security_amount)-fee_balance) paid,
                fee_balance,payment_status FROM students ORDER BY name""")
        elif kind=='rooms':
            rows=self.db.q('SELECT h.name hostel,r.room_number,r.floor,r.capacity,r.status,b.bed_label,b.status bed_status FROM rooms r JOIN hostels h ON h.id=r.hostel_id LEFT JOIN beds b ON b.room_id=r.id ORDER BY h.name,r.room_number,b.id')
        elif kind=='staff':
            rows=self.db.q("""SELECT st.id,st.category,st.name,COALESCE(st.gender,'') gender,
                COALESCE(st.contact_number,'') contact_number,COALESCE(st.address,'') address,
                COALESCE(st.notes,'') notes,st.active,
                COALESCE(h.name,'') hostel,COALESCE(r.room_number,'') room,
                COALESCE(b.bed_label,'') bed,
                COALESCE(a.occupation_date,'') occupation_date,
                COALESCE(a.vacation_date,'') vacation_date,
                CASE WHEN a.id IS NULL THEN 'Not Allotted' ELSE 'Living' END status
                FROM staff_occupants st
                LEFT JOIN allotments a ON a.staff_id=st.id AND a.status='Active'
                LEFT JOIN beds b ON b.id=a.bed_id
                LEFT JOIN rooms r ON r.id=b.room_id
                LEFT JOIN hostels h ON h.id=r.hostel_id
                ORDER BY st.category,st.name""")
        elif kind=='payments':
            rows=self.db.q('SELECT * FROM payments')
        elif kind=='left':
            rows=self.db.q('SELECT * FROM left_hostel ORDER BY vacation_date DESC,id DESC')
        else:
            rows=[]
        if rows:
            headers=list(rows[0].keys())
            ws.append(headers)
            for r in rows:
                ws.append([r[c] for c in headers])
        else:
            ws.append(['No records found'])
        ws.freeze_panes='A2'
        ws.auto_filter.ref=ws.dimensions
        for col in ws.columns:
            max_len=max(len(str(cell.value or '')) for cell in col)
            ws.column_dimensions[col[0].column_letter].width=min(max(max_len+2,12),40)
        wb.save(p)
        messagebox.showinfo('Saved',f'Excel export completed successfully.\\n\\nFile: {p}\\nRecords: {len(rows)}')
    def backup(self):
        p=filedialog.asksaveasfilename(defaultextension='.db',filetypes=[('SQLite Database','*.db')]);
        if p: self.db.c.commit(); shutil.copy2(DB,p); messagebox.showinfo('Backup','Database backup created.')
    def restore(self):
        p=filedialog.askopenfilename(filetypes=[('SQLite Database','*.db')]);
        if p and messagebox.askyesno('Restore','Restore selected database? Current database will be backed up first.'):
            shutil.copy2(DB,DB+'.before_restore.bak');self.db.c.close();shutil.copy2(p,DB);self.db=DBX();self.refresh_all();messagebox.showinfo('Restore','Database restored.')
    def new_session(self):
        if not self.allow_admin():return
        if not messagebox.askyesno('DANGER — FULL NEW SESSION','Start a completely blank hostel session? This removes students, staff, rooms, beds, allotments, payments, hostel-left history, leave history, ID cards, exchanges and import issues. Users and settings will remain. A backup will be created first.'):
            return
        if not messagebox.askyesno('FINAL CONFIRMATION','The live dashboard will become ZERO students, ZERO rooms, ZERO beds and ZERO occupied beds. You will need to add hostels/rooms/beds/occupants again. Continue?'):
            return
        bp=os.path.join(DATA,'backup_before_full_new_session_'+datetime.now().strftime('%Y%m%d_%H%M%S')+'.db'); self.db.c.commit(); shutil.copy2(DB,bp)
        try:
            self.db.c.execute('BEGIN')
            for t in ['room_exchanges','payments','id_cards','leave_records','left_hostel','allotments','beds','rooms','students','staff_occupants','import_issues']:
                self.db.c.execute('DELETE FROM '+t)
            self.db.c.execute('DELETE FROM hostels')
            self.db.c.execute("DELETE FROM sqlite_sequence WHERE name IN ('students','staff_occupants','rooms','beds','allotments','room_exchanges','payments','leave_records','left_hostel','id_cards','import_issues','hostels')")
            self.db.c.execute("INSERT INTO hostels(name,gender,active) VALUES('Boys','Male',1)")
            self.db.c.execute("INSERT INTO hostels(name,gender,active) VALUES('Girls','Female',1)")
            self.db.c.commit()
        except Exception as e:
            self.db.c.rollback();return messagebox.showerror('Reset failed',f'No reset was committed.\n\n{e}')
        self.db.log(self.user['username'],'FULL_NEW_SESSION','ALL','ALL',f'backup={bp}')
        self.refresh_all(); self.root.update_idletasks(); messagebox.showinfo('New session','Full new hostel session started.\n\nDashboard: 0 students / 0 rooms / 0 beds / 0 occupied.\n\nBackup: '+bp)

    # Users/settings
    def settings(self,f):
        ttk.Label(f,text='Users & Settings',style='Title.TLabel').pack(anchor='w');bar=ttk.Frame(f);bar.pack(fill='x');ttk.Button(bar,text='Add User',command=self.user_form).pack(side='left');ttk.Button(bar,text='Edit Selected',command=self.user_edit).pack(side='left',padx=4);ttk.Button(bar,text='Delete Selected',command=self.user_delete).pack(side='left')
        self.utree=self.tree(f,['Username','Role','Full Name','Active'],10)
        box=ttk.LabelFrame(f,text='Institution / Printing Settings',padding=10);box.pack(fill='x',pady=10);self.inst=tk.StringVar(value=self.setting('institution_name'));self.ward=tk.StringVar(value=self.setting('warden_name'));ttk.Label(box,text='Institution').grid(row=0,column=0,padx=5,pady=5);ttk.Entry(box,textvariable=self.inst,width=45).grid(row=0,column=1);ttk.Label(box,text='Warden / Signature').grid(row=1,column=0,padx=5,pady=5);ttk.Entry(box,textvariable=self.ward,width=45).grid(row=1,column=1);self.logo=tk.StringVar(value=self.setting('logo_path'));ttk.Label(box,text='Hostel Logo').grid(row=2,column=0,padx=5,pady=5);ttk.Entry(box,textvariable=self.logo,width=45).grid(row=2,column=1);ttk.Button(box,text='Choose Logo',command=self.choose_logo).grid(row=2,column=2,padx=5);ttk.Button(box,text='Save Settings',command=self.save_settings).grid(row=3,column=0,columnspan=3,pady=5)
    def refresh_users(self):
        if not hasattr(self,'utree'):return
        self.utree.delete(*self.utree.get_children());
        for r in self.db.q('SELECT id,username,role,full_name,active FROM users ORDER BY username'):self.utree.insert('', 'end',iid=r['id'],values=tuple(r)[1:])
    def user_form(self,uid=None):
        w=tk.Toplevel(self.root);w.title('User');frm=ttk.Frame(w,padding=18);frm.pack();old=self.db.one('SELECT * FROM users WHERE id=?',(uid,)) if uid else None;v={}
        for i,(lab,k) in enumerate([('Username','username'),('Full Name','full_name'),('Role','role'),('Password','password')]):
            ttk.Label(frm,text=lab).grid(row=i,column=0,pady=6,sticky='w');e=ttk.Combobox(frm,values=['Administrator','Data Entry User','Viewer'],state='readonly',width=28) if k=='role' else ttk.Entry(frm,width=30,show='*' if k=='password' else '');e.grid(row=i,column=1);v[k]=e
            if old and k!='password':e.insert(0,clean(old[k]))
        def save():
            if not v['username'].get().strip():return
            if uid:
                self.db.x('UPDATE users SET username=?,full_name=?,role=? WHERE id=?',(v['username'].get().strip(),v['full_name'].get().strip(),v['role'].get(),uid));
                if v['password'].get():h,s=phash(v['password'].get());self.db.x('UPDATE users SET password_hash=?,salt=? WHERE id=?',(h,s,uid))
            else:
                if not v['password'].get():return messagebox.showerror('Password','Enter a password.',parent=w)
                h,s=phash(v['password'].get());self.db.x('INSERT INTO users(username,password_hash,salt,role,full_name) VALUES(?,?,?,?,?)',(v['username'].get().strip(),h,s,v['role'].get(),v['full_name'].get().strip()))
            w.destroy();self.refresh_users()
        ttk.Button(frm,text='Save',command=save).grid(row=4,column=0,columnspan=2,pady=10)
    def user_edit(self):
        s=self.utree.selection();
        if s:self.user_form(int(s[0]))
    def user_delete(self):
        s=self.utree.selection();
        if s and messagebox.askyesno('Delete','Delete selected user?'):self.db.exec('DELETE FROM users WHERE id=?',(int(s[0]),));self.refresh_users()
    def choose_logo(self):
        p=filedialog.askopenfilename(filetypes=[('Image Files','*.png *.jpg *.jpeg *.gif')])
        if p:
            try:
                from PIL import Image as PILImage
                target=os.path.join(DATA,'hostel_logo.png')
                PILImage.open(p).convert('RGB').save(target,'PNG')
                self.logo.set(os.path.abspath(target))
            except ImportError:
                messagebox.showerror('Logo','Pillow is required for reliable logo printing. Run: py -m pip install Pillow')
            except Exception as e: messagebox.showerror('Logo','Could not save logo.\n'+str(e))
    def save_settings(self):
        for k,v in [('institution_name',self.inst.get().strip()),('warden_name',self.ward.get().strip()),('logo_path',self.logo.get().strip())]: self.db.exec('INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value',(k,v))
        messagebox.showinfo('Saved','Settings saved. The logo will be used on printable documents.')


    # Printing helpers
    def pdf_table(self,title,data):
        if not SimpleDocTemplate:return messagebox.showerror('PDF','Install reportlab first.')
        p=filedialog.asksaveasfilename(defaultextension='.pdf',filetypes=[('PDF','*.pdf')])
        if not p:return
        doc=SimpleDocTemplate(p,pagesize=A4,rightMargin=25,leftMargin=25,topMargin=25,bottomMargin=25);styles=getSampleStyleSheet();elems=[]
        logo=self.setting('logo_path')
        if logo and os.path.exists(logo):
            try:
                im=Image(logo,width=75,height=75); elems.append(im)
            except Exception: pass
        elems += [Paragraph(self.setting('institution_name'),styles['Title']),Paragraph(title,styles['Heading2']),Spacer(1,10)]
        t=Table([[str(x) for x in r] for r in data],repeatRows=1);t.setStyle(TableStyle([('GRID',(0,0),(-1,-1),.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.lightgrey),('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('VALIGN',(0,0),(-1,-1),'TOP'),('PADDING',(0,0),(-1,-1),5)]));elems.append(t);doc.build(elems)
        if messagebox.askyesno('Print','PDF created successfully. Send it to the default printer now?'):
            try: os.startfile(p,'print')
            except Exception: messagebox.showinfo('Print','PDF saved. Open it and choose Print from your PDF viewer.')
        else: messagebox.showinfo('Saved',p)
    def pdf_receipt(self,title,pairs,signature=False):
        data=[['Field','Details']]+[[a,str(b)] for a,b in pairs];
        if signature:data += [['',''],['Warden / Hostel In-charge',self.setting('warden_name')],['Signature','____________________________']]
        self.pdf_table(title,data)
    def print_gatepass(self,a,dt,reason): self.pdf_receipt('HOSTEL VACATION / GATE PASS',[('Student',a['name']),('Student ID',a['admission_no']),('University',a['university'] or '-'),('Group',a['group_no'] or '-'),('Semester',a['semester'] or '-'),('Hostel',a['hostel']),('Room',a['room_number']),('Bed',a['bed_label']),('Occupation Date',a['occupation_date']),('Vacation Date',dt),('Reason',reason or '-')],True)
    def print_exchange(self,eid):
        r=self.db.one('''SELECT e.exchange_date,e.reason,COALESCE(s.name,st.name) name,COALESCE(s.admission_no,st.category) ident,h1.name from_host,r1.room_number from_room,b1.bed_label from_bed,h2.name to_host,r2.room_number to_room,b2.bed_label to_bed FROM room_exchanges e LEFT JOIN students s ON s.id=e.student_id LEFT JOIN staff_occupants st ON st.id=e.staff_id JOIN beds b1 ON b1.id=e.from_bed_id JOIN rooms r1 ON r1.id=b1.room_id JOIN hostels h1 ON h1.id=r1.hostel_id JOIN beds b2 ON b2.id=e.to_bed_id JOIN rooms r2 ON r2.id=b2.room_id JOIN hostels h2 ON h2.id=r2.hostel_id WHERE e.id=?''',(eid,));self.pdf_receipt('ROOM / BED EXCHANGE RECEIPT',[('Occupant',r['name']),('ID / Category',r['ident']),('From',f"{r['from_host']} / Room {r['from_room']} / Bed {r['from_bed']}"),('To',f"{r['to_host']} / Room {r['to_room']} / Bed {r['to_bed']}"),('Exchange Date',r['exchange_date']),('Reason',r['reason'] or '-')],True)
    def refresh_all(self):
        self.refresh_dashboard();self.refresh_students();self.refresh_rooms();self.refresh_allotments();self.refresh_staff();self.refresh_left();self.refresh_reports();self.refresh_users();self.global_search_run()

if __name__=='__main__':
    App().root.mainloop()

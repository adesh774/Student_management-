import os, io, sqlite3, traceback, re
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, flash, session, send_file, send_from_directory
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
import pandas as pd
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DB_PATH = os.path.join(BASE_DIR, "portal.db")
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
TEMPLATE_DIR = os.path.join(BASE_DIR, "sample_data")
os.makedirs(UPLOAD_DIR, exist_ok=True)

app = Flask(__name__)
app.config["SECRET_KEY"] = "change-this-in-prod"

# ---------------- DB helpers ----------------
def get_db():
    con = sqlite3.connect(DB_PATH, timeout=30)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA busy_timeout=5000")
    con.execute("PRAGMA journal_mode=WAL")
    con.execute("PRAGMA synchronous=NORMAL")
    con.execute("PRAGMA foreign_keys=ON")
    return con

def run_schema():
    with get_db() as con:
        con.executescript("""
        PRAGMA foreign_keys=ON;
        CREATE TABLE IF NOT EXISTS users(
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          username TEXT UNIQUE NOT NULL,
          password TEXT NOT NULL,
          role TEXT NOT NULL CHECK(role IN ('admin'))
        );

        CREATE TABLE IF NOT EXISTS uploads(
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          filename TEXT NOT NULL,
          upload_type TEXT NOT NULL,
          row_count INTEGER DEFAULT 0,
          created_at TEXT NOT NULL,
          uploader_username TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS students(
          roll_no TEXT PRIMARY KEY,
          name TEXT NOT NULL,
          email TEXT,
          phone TEXT,
          father_name TEXT,
          father_phone TEXT,
          course TEXT,
          address TEXT,
          session TEXT,
          regn_no TEXT
        );

        CREATE TABLE IF NOT EXISTS upload_students_map(
          upload_id INTEGER NOT NULL REFERENCES uploads(id) ON DELETE CASCADE,
          roll_no TEXT NOT NULL,
          created_new INTEGER NOT NULL DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS attendance(
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          roll_no TEXT NOT NULL,
          subject TEXT NOT NULL,
          attended INTEGER DEFAULT 0,
          total INTEGER DEFAULT 0,
          course TEXT,
          source_upload_id INTEGER REFERENCES uploads(id) ON DELETE SET NULL,
          UNIQUE(roll_no, subject),
          FOREIGN KEY (roll_no) REFERENCES students(roll_no) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS marks(
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          roll_no TEXT NOT NULL,
          exam TEXT NOT NULL,
          subject TEXT NOT NULL,
          max_marks INTEGER NOT NULL,
          marks_obtained INTEGER NOT NULL,
          credits INTEGER,
          course TEXT,
          source_upload_id INTEGER REFERENCES uploads(id) ON DELETE SET NULL,
          UNIQUE(roll_no, exam, subject),
          FOREIGN KEY (roll_no) REFERENCES students(roll_no) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS remarks(
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          roll_no TEXT NOT NULL,
          remark_text TEXT NOT NULL,
          author_username TEXT NOT NULL,
          created_at TEXT NOT NULL,
          FOREIGN KEY (roll_no) REFERENCES students(roll_no) ON DELETE CASCADE
        );
        """)
        # migrations for old DBs
        cur = con.cursor()
        cur.execute("PRAGMA table_info(students)")
        scols = [r[1] for r in cur.fetchall()]
        if "course" not in scols:   con.execute("ALTER TABLE students ADD COLUMN course TEXT")
        if "address" not in scols:  con.execute("ALTER TABLE students ADD COLUMN address TEXT")
        if "session" not in scols:  con.execute("ALTER TABLE students ADD COLUMN session TEXT")
        if "regn_no" not in scols:  con.execute("ALTER TABLE students ADD COLUMN regn_no TEXT")

        cur.execute("PRAGMA table_info(attendance)"); acols=[r[1] for r in cur.fetchall()]
        if "course" not in acols:   con.execute("ALTER TABLE attendance ADD COLUMN course TEXT")

        cur.execute("PRAGMA table_info(marks)"); mcols=[r[1] for r in cur.fetchall()]
        if "course" not in mcols:   con.execute("ALTER TABLE marks ADD COLUMN course TEXT")
        con.commit()

def ensure_admin():
    with get_db() as con:
        r = con.execute("SELECT 1 FROM users WHERE role='admin' LIMIT 1").fetchone()
        if not r:
            con.execute("INSERT INTO users(username,password,role) VALUES(?,?,?)",
                        ("admin", generate_password_hash("admin123"), "admin"))
            con.commit()

# --------------- Utils ---------------
def load_table(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xlsx":
        return pd.read_excel(path, sheet_name=0)
    if ext == ".csv":
        return pd.read_csv(path)
    raise ValueError("Only .xlsx or .csv supported.")

def save_upload(fs):
    ext = os.path.splitext(fs.filename)[1].lower()
    if ext not in {".xlsx", ".csv"}:
        raise ValueError("Only .xlsx or .csv allowed.")
    name = f"{int(pd.Timestamp.now().timestamp()*1000)}_{secure_filename(fs.filename)}"
    full = os.path.join(UPLOAD_DIR, name)
    fs.save(full)
    return full, name

def create_upload_record(filename: str, upload_type: str) -> int:
    with get_db() as con:
        cur = con.cursor()
        cur.execute(
            "INSERT INTO uploads(filename, upload_type, row_count, created_at, uploader_username) VALUES(?,?,?,?,?)",
            (filename, upload_type, 0, datetime.now().isoformat(timespec='seconds'), session.get("username", "admin")),
        )
        con.commit()
        return cur.lastrowid

def bump_rowcount(upload_id: int, n: int):
    with get_db() as con:
        con.execute("UPDATE uploads SET row_count=row_count+? WHERE id=?", (n, upload_id))
        con.commit()

def _parse_course(value):
    if value is None: return None
    s = str(value).strip()
    if not s or s.lower()=='nan': return None
    return s

def _clean_id(val):
    s = str(val).strip()
    if not s or s.lower() == 'nan':
        return ''
    s2 = s.replace(',', '').strip()
    if re.fullmatch(r'\d+(\.0+)?', s2):
        return s2.split('.')[0]
    return s

def _none_if_blank_or_zero(v):
    if v is None:
        return None
    s = str(v).strip()
    if s == "" or s.lower() == "nan" or s == "0":
        return None
    return s

def _norm(s):
    # collapse underscores etc.
    return (str(s).strip().lower()
            .replace("'", "")
            .replace('&', 'and')
            .replace('.', '')
            .replace('-', ' ')
            .replace('/', ' ')
            .replace('\\', ' ')
            .replace('_', ' ')
            .replace('  ',' '))

def _get_ci(d, keys):
    keymap = {_norm(k): k for k in d.keys()}
    for k in keys:
        kk = _norm(k)
        if kk in keymap:
            return d[keymap[kk]]
    return None

# --- smart resolver: roll → canonical roll_no using roll/regn/name
def _resolve_roll_smart(cur, roll_in: str, *, name_in: str|None=None, regn_in: str|None=None) -> str:
    roll_in = _clean_id(roll_in)
    regn_in = _clean_id(regn_in) if regn_in else None

    if roll_in:
        r = cur.execute("SELECT roll_no FROM students WHERE roll_no=? OR regn_no=?", (roll_in, roll_in)).fetchone()
        if r: return r["roll_no"]
    if regn_in:
        r = cur.execute("SELECT roll_no FROM students WHERE roll_no=? OR regn_no=?", (regn_in, regn_in)).fetchone()
        if r: return r["roll_no"]

    if name_in:
        key = re.sub(r'\s+', '', str(name_in).strip().lower())
        row = cur.execute("SELECT roll_no, name FROM students").fetchall()
        for rr in row:
            if re.sub(r'\s+','', rr["name"].strip().lower()) == key:
                return rr["roll_no"]

    return roll_in or (regn_in or "")

# ========== STUDENTS IMPORT ==========
def import_students(df: pd.DataFrame, upload_id: int):
    cols = list(df.columns)

    def pick(opt_list):
        for o in opt_list:
            if o in cols:
                return o
        return None

    ROLL_ALIASES = [
        "Roll_number", "Roll_Number",
        "University Roll", "UNIVERSITY ROLL",
        "University Roll No", "UNIVERSITY ROLL NO"
    ]
    REGN_ALIASES = [
        "Regn No", "REGN NO", "Regn. No", "REGN. NO", "Regn. no",
        "Registration No", "REGISTRATION NO", "Reg No", "REG NO", "Regd No", "REGD NO"
    ]

    ACCEPTED = {
        "name":         ["Name of student", "NAME OF STUDENT"],
        "father_name":  ["Father's name", "FATHER'S NAME"],
        "phone":        ["student contact", "STUDENT Contact", "STUDENT Conact"],
        "father_phone": ["PARENT Contact"],
        "address":      ["address", "Address"],
        "email":        ["email", "Email"],
        "session":      ["session", "Session"],
    }

    roll_col = pick(ROLL_ALIASES)
    regn_col = pick(REGN_ALIASES)
    name_col = pick(ACCEPTED["name"])

    if not name_col:
        raise ValueError("Students: missing required case-sensitive column Name of student / NAME OF STUDENT.")
    if not roll_col and not regn_col:
        raise ValueError("Students: need either a Roll column (Roll_number/Roll_Number/University Roll) or a Regn/Registration No column.")

    n = 0
    with get_db() as con:
        cur = con.cursor()
        for _, r in df.iterrows():
            roll_raw = _clean_id(r.get(roll_col, "")) if roll_col else ""
            regn_raw = _clean_id(r.get(regn_col, "")) if regn_col else ""
            roll_no  = roll_raw or regn_raw

            name = r.get(name_col, "")

            def get(opt):
                col = pick(opt); return r.get(col, None) if col else None

            father_name  = get(ACCEPTED["father_name"])
            phone        = get(ACCEPTED["phone"])
            father_phone = get(ACCEPTED["father_phone"])
            address      = get(ACCEPTED["address"])
            email        = get(ACCEPTED["email"])
            session_txt  = get(ACCEPTED["session"])

            update_or_create_student_from_row(cur, {
                "roll_no": roll_no,
                "regn_no": regn_raw or None,
                "name": name,
                "father_name": father_name,
                "father_phone": father_phone,
                "phone": phone,
                "email": email,
                "address": address,
                "session": session_txt,
            })

            exists = cur.execute("SELECT 1 FROM students WHERE roll_no=? OR regn_no=?", (roll_no, roll_no)).fetchone()
            cur.execute("INSERT INTO upload_students_map(upload_id, roll_no, created_new) VALUES(?,?,?)",
                        (upload_id, roll_no, 0 if exists else 1))
            n += 1
        con.commit()
    bump_rowcount(upload_id, n)

# -------- Attendance/Marks helpers --------
def ensure_marks_long(df: pd.DataFrame) -> pd.DataFrame:
    cols_norm = [_norm(c) for c in df.columns]

    # LONG format
    if {'roll_no','subject'}.issubset(set(cols_norm)) and ('marks_obtained' in cols_norm or 'marks' in cols_norm):
        out = df.copy()
        rename = {}
        for c in out.columns:
            n = _norm(c)
            if n == 'marks' and 'marks_obtained' not in [_norm(x) for x in out.columns]:
                rename[c] = 'marks_obtained'
            if n == 'roll no':
                rename[c] = 'roll_no'
            if n in ('regn no','registration no','reg no','regd no','reg','regn'):
                rename[c] = 'regn_no'
        if rename: out = out.rename(columns=rename)
        if 'max_marks' not in [_norm(x) for x in out.columns]: out['max_marks'] = 30
        if 'exam' not in [_norm(x) for x in out.columns]: out['exam'] = 'MTE'
        out['roll_no'] = out['roll_no'].apply(_clean_id)
        if 'regn_no' in out.columns:
            out['regn_no'] = out['regn_no'].apply(_clean_id)
        if 'course' not in out.columns:
            for c in df.columns:
                if _norm(c) in ('course','cource','semester','sem'): out['course'] = df[c]; break
        return out

    # WIDE format
    id_candidates = {
        'roll_no','roll no','roll number','rollnumber','roll',
        'university roll','university roll no','university roll number',
        'regn no','registration no','reg no','regd no','reg','regn',
        'name','student name','students name',"student's name",'name of student','student s name',
        'email','e mail','phone','mobile','contact','student contact','students contact','student conact',
        "father's name",'father name','fathers name','father phone','parent contact','parent phone',
        'course','cource','semester','sem','address','session','sr no','sr'
    }
    subject_cols = [c for c in df.columns if _norm(c) not in id_candidates]
    if not subject_cols:
        raise ValueError('Marks: could not detect any subject columns.')

    rows = []
    for _, r in df.iterrows():
        rdict = {k: r[k] for k in df.columns}
        roll  = _get_ci(rdict, ['roll_no','roll no','roll number','rollnumber','roll','university roll','university roll no','university roll number'])
        regn  = _get_ci(rdict, ['regn no','registration no','reg no','regd no','reg','regn'])
        name  = _get_ci(rdict, ['name of student','student name','students name',"student's name",'name'])
        email = _get_ci(rdict, ['email','e mail'])
        phone = _get_ci(rdict, ['student contact','students contact','student conact','phone','mobile','contact'])
        father_name  = _get_ci(rdict, ["father's name",'father name','fathers name'])
        father_phone = _get_ci(rdict, ['parent contact','father phone','parent phone'])
        course = _parse_course(_get_ci(rdict, ['course','cource','semester','sem']))

        for subj_col in subject_cols:
            val = rdict.get(subj_col, '')
            s = str(val).strip().lower()
            if s in ('', 'ne', 'na', 'absent', 'n/a'):
                continue
            try:
                got = int(float(val))
            except Exception:
                continue
            rows.append({
                'roll_no': _clean_id(roll) or (str(name).strip() or 'UNKNOWN'),
                'regn_no': _clean_id(regn) if regn else None,
                'name': name,
                'exam': 'MTE',
                'subject': str(subj_col).strip(),
                'max_marks': 30,
                'marks_obtained': got,
                'credits': None,
                'course': course,
                'email': email,
                'phone': phone,
                'father_name': father_name,
                'father_phone': father_phone
            })
    if not rows:
        raise ValueError('Marks: could not detect subjects/values in wide sheet. Provide long format or numeric subject columns.')
    return pd.DataFrame(rows)

def ensure_attendance_long(df: pd.DataFrame) -> pd.DataFrame:
    cols_norm = [_norm(c) for c in df.columns]

    # LONG format
    if {'roll_no','subject','attended','total'}.issubset(set(cols_norm)):
        out = df.copy()
        ren = {}
        for c in out.columns:
            n = _norm(c)
            if n == 'roll no': ren[c] = 'roll_no'
            if n in ('regn no','registration no','reg no','regd no','reg','regn'):
                ren[c] = 'regn_no'
        if ren: out = out.rename(columns=ren)
        out['roll_no'] = out['roll_no'].apply(_clean_id)
        if 'regn_no' in out.columns:
            out['regn_no'] = out['regn_no'].apply(_clean_id)
        if 'course' not in out.columns:
            for c in df.columns:
                if _norm(c) in ('course','cource','semester','sem'): out['course'] = df[c]; break
        return out

    # WIDE format
    pair_rx = re.compile(r'^(?P<sub>.+?)\s*[-_/]?\s*(?P<kind>attended|total|att|tot)\s*$', re.I)
    pairs = {}
    for c in df.columns:
        m = pair_rx.match(str(c))
        if not m: continue
        sub = m.group('sub').strip()
        kind = m.group('kind').lower()
        kind = 'attended' if kind in ('attended','att') else 'total'
        pairs.setdefault(sub, {'attended': None, 'total': None})
        pairs[sub][kind] = c
    if not pairs:
        raise ValueError('Attendance: could not find wide columns. Expected "Subject - Attended" and "Subject - Total" (Att/Tot allowed).')

    rows = []
    for _, r in df.iterrows():
        rdict = {k: r[k] for k in df.columns}
        roll  = _get_ci(rdict, ['roll_no','roll no','roll number','rollnumber','roll','university roll','university roll no','university roll number'])
        regn  = _get_ci(rdict, ['regn no','registration no','reg no','regd no','reg','regn'])
        name  = _get_ci(rdict, ['name of student','student name','students name',"student's name",'name'])
        email = _get_ci(rdict, ['email','e mail'])
        phone = _get_ci(rdict, ['student contact','students contact','student conact','phone','mobile','contact'])
        father_name  = _get_ci(rdict, ["father's name",'father name','fathers name'])
        father_phone = _get_ci(rdict, ['parent contact','father phone','parent phone'])
        course       = _parse_course(_get_ci(rdict, ['course','cource','semester','sem']))

        for sub, d in pairs.items():
            att_col, tot_col = d['attended'], d['total']
            if not att_col or not tot_col: continue
            att = rdict.get(att_col, 0)
            tot = rdict.get(tot_col, 0)
            try:
                att_i = int(float(att)); tot_i = int(float(tot))
            except Exception:
                continue
            rows.append({
                'roll_no': _clean_id(roll) or (str(name).strip() or 'UNKNOWN'),
                'regn_no': _clean_id(regn) if regn else None,
                'name': name,
                'subject': str(sub).strip(),
                'attended': att_i,
                'total': tot_i,
                'course': course,
                'email': email,
                'phone': phone,
                'father_name': father_name,
                'father_phone': father_phone
            })
    if not rows:
        raise ValueError('Attendance: detected pair columns but values are empty/non-numeric.')
    return pd.DataFrame(rows)

def do_import_attendance(df: pd.DataFrame, upload_id: int):
    df = ensure_attendance_long(df)
    cols = [c.strip().lower() for c in df.columns]
    for col in ["roll_no","subject","attended","total"]:
        if col not in cols:
            raise ValueError(f"Attendance: missing column '{col}'")

    n = 0
    subjects = set()
    with get_db() as con:
        cur = con.cursor()
        for _, r in df.iterrows():
            r = {str(k).strip(): r[k] for k in df.keys()}
            roll_raw = _clean_id(r.get("roll_no",""))
            name_raw = r.get("name","")
            regn_raw = _clean_id(r.get("regn_no","")) if "regn_no" in r else None
            subj = str(r.get("subject","")).strip()
            if not (roll_raw or regn_raw or name_raw) or not subj:
                continue

            roll = _resolve_roll_smart(cur, roll_raw, name_in=name_raw, regn_in=regn_raw)

            payload = {k.lower(): r[k] for k in r}
            payload["roll_no"] = roll
            if regn_raw: payload["regn_no"] = regn_raw
            update_or_create_student_from_row(cur, payload)

            attended = int(float(r.get("attended",0))) if str(r.get("attended","")).strip().lower() not in ("", "nan") else 0
            total    = int(float(r.get("total",0)))    if str(r.get("total","")).strip().lower() not in ("", "nan") else 0
            course   = _parse_course(r.get("course", r.get("cource", r.get("semester"))))
            cur.execute(
                """INSERT INTO attendance(roll_no,subject,attended,total,course,source_upload_id)
                   VALUES(?,?,?,?,?,?)
                   ON CONFLICT(roll_no,subject) DO UPDATE SET
                     attended=excluded.attended,
                     total=excluded.total,
                     course=COALESCE(excluded.course, attendance.course),
                     source_upload_id=excluded.source_upload_id""",
                (roll, subj, attended, total, course, upload_id),
            )
            subjects.add(subj)
            n += 1
        con.commit()
    bump_rowcount(upload_id, n)
    return n, sorted(subjects)

def do_import_marks(df: pd.DataFrame, upload_id: int):
    df = ensure_marks_long(df)
    df.columns = [str(c).strip() for c in df.columns]
    lower = [c.lower() for c in df.columns]
    for col in ["roll_no", "subject", "marks_obtained"]:
        if col not in lower:
            raise ValueError(f"Marks: missing column '{col}'")
    if "max_marks" not in lower:
        df["max_marks"] = 30
    if "exam" not in lower:
        df["exam"] = "MTE"

    n = 0
    subjects, exams = set(), set()
    with get_db() as con:
        cur = con.cursor()
        for _, r in df.iterrows():
            r = dict(r)
            roll_raw = _clean_id(r.get("roll_no",""))
            name_raw = r.get("name","")
            regn_raw = _clean_id(r.get("regn_no","")) if "regn_no" in r else None
            exam = str(r.get("exam","")).strip() or "MTE"
            subj = str(r.get("subject","")).strip()
            if not (roll_raw or regn_raw or name_raw) or not subj:
                continue

            roll = _resolve_roll_smart(cur, roll_raw, name_in=name_raw, regn_in=regn_raw)

            payload = {k.lower(): r[k] for k in r}
            payload["roll_no"] = roll
            if regn_raw: payload["regn_no"] = regn_raw
            update_or_create_student_from_row(cur, payload)

            try: maxm = int(float(r.get("max_marks", 30)))
            except: maxm = 30
            try: got = int(float(r.get("marks_obtained", 0)))
            except: got = 0
            credits = None
            if "credits" in lower:
                cr = str(r.get("credits","")).strip()
                if cr and cr.lower() != "nan":
                    try: credits = int(float(cr))
                    except: credits = None
            course = _parse_course(r.get("course", r.get("cource", r.get("semester"))))
            cur.execute(
                """INSERT INTO marks(roll_no,exam,subject,max_marks,marks_obtained,credits,course,source_upload_id)
                   VALUES(?,?,?,?,?,?,?,?)
                   ON CONFLICT(roll_no,exam,subject) DO UPDATE SET
                     max_marks=excluded.max_marks,
                     marks_obtained=excluded.marks_obtained,
                     credits=COALESCE(excluded.credits, marks.credits),
                     course=COALESCE(excluded.course, marks.course),
                     source_upload_id=excluded.source_upload_id""",
                (roll, exam, subj, maxm, got, credits, course, upload_id),
            )
            subjects.add(subj); exams.add(exam)
            n += 1
        con.commit()
    bump_rowcount(upload_id, n)
    return n, sorted(subjects), sorted(exams)

# --------------- Student record helpers ---------------
def update_or_create_student_from_row(cur, rowdict):
    def get_course(d):
        for k in ['course','cource','semester','sem']:
            if k in d and str(d[k]).strip():
                return _parse_course(d[k])
        return _parse_course(rowdict.get('course'))

    roll = _clean_id(rowdict.get("roll_no",""))
    regn = _none_if_blank_or_zero(rowdict.get('regn_no',''))
    name_raw = _none_if_blank_or_zero(rowdict.get('name',''))
    name = name_raw if name_raw else (f"Student {roll}" if roll else None)
    email = _none_if_blank_or_zero(rowdict.get('email',''))
    phone = _none_if_blank_or_zero(rowdict.get('phone',''))
    father_name = _none_if_blank_or_zero(rowdict.get('father_name',''))
    father_phone = _none_if_blank_or_zero(rowdict.get('father_phone',''))
    course = get_course(rowdict)
    address = _none_if_blank_or_zero(rowdict.get('address',''))
    session_val = _none_if_blank_or_zero(rowdict.get('session',''))

    existing = None
    if roll:
        existing = cur.execute("SELECT * FROM students WHERE roll_no=? OR regn_no=?", (roll, roll)).fetchone()
    if not existing and regn:
        existing = cur.execute("SELECT * FROM students WHERE roll_no=? OR regn_no=?", (regn, regn)).fetchone()

    if not existing:
        if not roll and regn:
            roll = regn
        if not roll:
            return
        cur.execute(
            "INSERT INTO students(roll_no,name,email,phone,father_name,father_phone,course,address,session,regn_no) VALUES(?,?,?,?,?,?,?,?,?,?)",
            (roll, name or f"Student {roll}", email, phone, father_name, father_phone, course, address, session_val, regn)
        )
    else:
        anchor = existing["roll_no"]
        cur.execute(
            """UPDATE students
                   SET name = COALESCE(?, name),
                       email = COALESCE(?, email),
                       phone = COALESCE(?, phone),
                       father_name = COALESCE(?, father_name),
                       father_phone = COALESCE(?, father_phone),
                       course = COALESCE(?, course),
                       address = COALESCE(?, address),
                       session = COALESCE(?, session),
                       regn_no = COALESCE(?, regn_no)
                 WHERE roll_no = ?""",
            (name or None, email, phone, father_name, father_phone, course, address, session_val, regn, anchor)
        )

# --------------- Delete helpers ---------------
def delete_upload(upload_id: int) -> bool:
    with get_db() as con:
        cur = con.cursor()
        up = cur.execute("SELECT * FROM uploads WHERE id=?", (upload_id,)).fetchone()
        if not up:
            return False
        cur.execute("DELETE FROM attendance WHERE source_upload_id=?", (upload_id,))
        cur.execute("DELETE FROM marks WHERE source_upload_id=?", (upload_id,))
        created = cur.execute(
            "SELECT roll_no FROM upload_students_map WHERE upload_id=? AND created_new=1", (upload_id,)
        ).fetchall()
        if created:
            rolls = [r["roll_no"] for r in created]
            q = f"DELETE FROM students WHERE roll_no IN ({','.join('?'*len(rolls))})"
            cur.execute(q, rolls)
        cur.execute("DELETE FROM upload_students_map WHERE upload_id=?", (upload_id,))
        cur.execute("DELETE FROM uploads WHERE id=?", (upload_id,))
        con.commit()
    try:
        fpath = os.path.join(UPLOAD_DIR, up["filename"])
        if os.path.exists(fpath): os.remove(fpath)
    except Exception:
        traceback.print_exc()
    return True

def _delete_student_and_related(roll: str):
    with get_db() as con:
        con.execute('DELETE FROM attendance WHERE roll_no=?', (roll,))
        con.execute('DELETE FROM marks WHERE roll_no=?', (roll,))
        con.execute('DELETE FROM remarks WHERE roll_no=?', (roll,))
        con.execute('DELETE FROM students WHERE roll_no=?', (roll,))
        con.commit()

# --------------- Wide views ---------------
def letter_grade(pct: float) -> str:
    if pct >= 85: return "A+"
    if pct >= 75: return "A"
    if pct >= 65: return "B+"
    if pct >= 55: return "B"
    if pct >= 45: return "C"
    if pct >= 35: return "D"
    return "F"

def _fetch_marks_for_rolls(rolls):
    q = ",".join("?"*len(rolls))
    with get_db() as con:
        return con.execute(
            f"SELECT exam, subject, max_marks, marks_obtained FROM marks WHERE roll_no IN ({q}) ORDER BY exam, subject", tuple(rolls)
        ).fetchall()

def _fetch_att_for_rolls(rolls):
    q = ",".join("?"*len(rolls))
    with get_db() as con:
        return con.execute(
            f"SELECT subject, attended, total FROM attendance WHERE roll_no IN ({q}) ORDER BY subject", tuple(rolls)
        ).fetchall()

def build_marks_grid(rolls):
    rows = _fetch_marks_for_rolls(rolls)
    if not rows: return [], []
    subjects = sorted({r["subject"] for r in rows})
    exams = []
    for r in rows:
        if r["exam"] not in exams:
            exams.append(r["exam"])
    grid_rows = []
    for ex in exams:
        row = {"Exam": ex}
        total_max = 0; total_got = 0
        for sub in subjects:
            match = [r for r in rows if r["exam"]==ex and r["subject"]==sub]
            if match:
                rmatch = match[0]
                pct = (100.0 * rmatch["marks_obtained"] / rmatch["max_marks"]) if rmatch["max_marks"] else 0.0
                row[sub] = f"{letter_grade(pct)} ({rmatch['marks_obtained']}/{rmatch['max_marks']})"
                total_max += rmatch["max_marks"]; total_got += rmatch["marks_obtained"]
            else:
                row[sub] = "—"
        row["Total"] = f"{total_got}/{total_max} ({(100.0*total_got/total_max):.0f}%)" if total_max>0 else "—"
        grid_rows.append(row)
    headers = ["Exam"] + subjects + ["Total"]
    return headers, grid_rows

def build_attendance_grid(rolls):
    rows = _fetch_att_for_rolls(rolls)
    if not rows: return [], []
    subjects = [r["subject"] for r in rows]
    row = {"Row": "Attendance"}
    t_att = 0; t_tot = 0
    for r in rows:
        pct = (100.0 * r["attended"] / r["total"]) if r["total"] else 0.0
        row[r["subject"]] = f"{r['attended']}/{r['total']} ({pct:.0f}%)"
        t_att += r["attended"]; t_tot += r["total"]
    row["Total"] = f"{t_att}/{t_tot} ({(100.0*t_att/t_tot):.0f}%)" if t_tot>0 else "—"
    headers = ["Row"] + subjects + ["Total"]
    return headers, [row]

# --------------- Routes ---------------
@app.route('/')
def home():
    return render_template('index.html')

@app.route('/admin')
def admin_root():
    if session.get('role')=='admin':
        return redirect(url_for('admin_imports'))
    return redirect(url_for('admin_login'))

@app.route('/account')
def account_shortcut():
    return redirect(url_for('admin_account'))

@app.route('/templates/<kind>')
def download_template(kind):
    filename = kind
    path = os.path.join(TEMPLATE_DIR, filename)
    if not os.path.exists(path):
        flash('Template not found','error')
        return redirect(url_for('home'))
    return send_from_directory(TEMPLATE_DIR, filename, as_attachment=True)

# download an uploaded file
@app.route('/admin/uploads/<int:upload_id>/download', endpoint='admin_download_upload')
def admin_download_upload(upload_id):
    if session.get('role') != 'admin':
        return redirect(url_for('admin_login'))
    with get_db() as con:
        row = con.execute('SELECT filename FROM uploads WHERE id=?', (upload_id,)).fetchone()
    if not row:
        flash('Upload not found.','error'); return redirect(url_for('admin_imports'))
    fpath = os.path.join(UPLOAD_DIR, row['filename'])
    if not os.path.exists(fpath):
        flash('File missing on disk.','error'); return redirect(url_for('admin_imports'))
    return send_from_directory(UPLOAD_DIR, row['filename'], as_attachment=True)

@app.route('/search')
def search():
    roll=(request.args.get('roll') or '').strip()
    if not roll:
        flash('Enter a roll number','error'); return redirect(url_for('home'))
    # accept roll or regn in search box
    return redirect(url_for('student_view', roll_no=roll))

@app.route('/student/<roll_no>')
def student_view(roll_no):
    key = _clean_id(roll_no)
    with get_db() as con:
        cur=con.cursor()
        s=cur.execute('SELECT * FROM students WHERE roll_no=? OR regn_no=?',(key, key)).fetchone()
        if not s:
            flash('Student not found.','error'); return redirect(url_for('home'))

        # collect possible roll keys that might hold this student's data
        alt_rolls = {s['roll_no']}
        if s['regn_no']: alt_rolls.add(s['regn_no'])
        for rr in cur.execute("SELECT roll_no FROM students WHERE regn_no=? OR roll_no=?", (s['roll_no'], s['regn_no'] or s['roll_no'])).fetchall():
            alt_rolls.add(rr['roll_no'])
        rolls_list = sorted(alt_rolls)

        remarks=cur.execute('SELECT * FROM remarks WHERE roll_no=? ORDER BY created_at DESC',(s['roll_no'],)).fetchall()

    marks_headers, marks_rows = build_marks_grid(rolls_list)
    att_headers, att_rows = build_attendance_grid(rolls_list)
    return render_template('student.html', student=s, remarks=remarks,
                           marks_headers=marks_headers, marks_rows=marks_rows,
                           att_headers=att_headers, att_rows=att_rows)

@app.route('/student/<roll_no>/remark', methods=['POST'])
def add_remark(roll_no):
    text=(request.form.get('remark') or '').strip()
    if not text:
        flash('Remark cannot be empty.','error'); return redirect(url_for('student_view', roll_no=roll_no))
    author_label = 'HOD' if session.get('role')=='admin' else (session.get('username','User'))
    key = _clean_id(roll_no)
    with get_db() as con:
        con.execute('INSERT INTO remarks(roll_no,remark_text,author_username,created_at) VALUES(?,?,?,?)',
                    (key, text, author_label, datetime.now().isoformat(timespec='seconds')))
        con.commit()
    flash('Remark added.','success'); return redirect(url_for('student_view', roll_no=key))

# --- Admin auth ---
@app.route('/admin/login', methods=['GET','POST'])
def admin_login():
    if request.method=='POST':
        u=(request.form.get('username') or '').strip()
        p=(request.form.get('password') or '')
        with get_db() as con:
            row=con.execute('SELECT * FROM users WHERE username=?',(u,)).fetchone()
        if row and check_password_hash(row['password'], p) and row['role']=='admin':
            session['username']=u; session['role']='admin'
            flash('Welcome, Admin!','success'); return redirect(url_for('admin_imports'))
        flash('Invalid credentials','error')
    return render_template('admin_login.html')

@app.route('/admin/logout')
def admin_logout():
    session.clear(); flash('Logged out.','success'); return redirect(url_for('admin_login'))

# --- Admin account management ---
@app.route('/admin/account', methods=['GET','POST'])
def admin_account():
    if session.get('role') != 'admin':
        return redirect(url_for('admin_login'))

    if request.method == 'POST':
        current = request.form.get('current_password') or ''
        new_user = (request.form.get('new_username') or '').strip()
        new_pass = request.form.get('new_password') or ''

        with get_db() as con:
            row = con.execute("SELECT * FROM users WHERE username=?", (session.get('username'),)).fetchone()
            if not row or not check_password_hash(row['password'], current):
                flash('Current password is incorrect.','error')
                return redirect(url_for('admin_account'))

            if new_user and new_user != row['username']:
                exists = con.execute("SELECT 1 FROM users WHERE username=?", (new_user,)).fetchone()
                if exists:
                    flash('That username is already taken.','error')
                    return redirect(url_for('admin_account'))

            username_to_set = new_user if new_user else row['username']
            password_to_set = generate_password_hash(new_pass) if new_pass else row['password']

            con.execute("UPDATE users SET username=?, password=? WHERE id=?", (username_to_set, password_to_set, row['id']))
            con.commit()

        session['username'] = username_to_set
        flash('Account updated.','success')
        return redirect(url_for('admin_account'))

    return render_template('admin_account.html', username=session.get('username'))

# --- Admin: Single uploads + recent uploads list ---
@app.route('/admin/imports')
def admin_imports():
    if session.get('role')!='admin': return redirect(url_for('admin_login'))
    with get_db() as con:
        stats = {
            'students': con.execute('SELECT COUNT(*) FROM students').fetchone()[0],
            'attendance': con.execute('SELECT COUNT(*) FROM attendance').fetchone()[0],
            'marks': con.execute('SELECT COUNT(*) FROM marks').fetchone()[0],
        }
        recent = con.execute('SELECT * FROM uploads ORDER BY id DESC LIMIT 50').fetchall()
    student_cols = ["Roll_number","Name of student","Father's name","student contact","address","email","session","(optional) Regn No / Registration No"]
    student_cols_alt = [
        "Roll_Number","NAME OF STUDENT","FATHER'S NAME","STUDENT Contact","PARENT Contact",
        "Address","Email","session","(optional) REGN NO / REGISTRATION NO"
    ]
    attendance_cols = ["roll_no","subject","attended","total"]
    marks_cols = ["roll_no","exam","subject","max_marks","marks_obtained","credits"]
    return render_template('admin_imports.html', stats=stats, uploads=recent,
                           student_cols=student_cols, student_cols_alt=student_cols_alt,
                           attendance_cols=attendance_cols, marks_cols=marks_cols)

@app.route('/admin/upload/students', methods=['POST'])
def upload_students():
    if session.get('role')!='admin': return redirect(url_for('admin_login'))
    f = request.files.get('students_file')
    if not f or not f.filename:
        flash('Choose a Students Excel/CSV file.','error'); return redirect(url_for('admin_imports'))
    try:
        path, name = save_upload(f)
        up_id = create_upload_record(name, 'Students')
        df = load_table(path)
        import_students(df, up_id)
        flash('Students imported/updated successfully.','success')
    except Exception as e:
        traceback.print_exc(); flash(f'Import failed (Students): {e}','error')
    return redirect(url_for('admin_imports'))

@app.route('/admin/upload/attendance', methods=['POST'])
def upload_attendance():
    if session.get('role')!='admin': return redirect(url_for('admin_login'))
    f = request.files.get('attendance_file')
    if not f or not f.filename:
        flash('Choose an Attendance Excel/CSV file.','error'); return redirect(url_for('admin_imports'))
    try:
        path, name = save_upload(f)
        up_id = create_upload_record(name, 'Attendance')
        df_raw = load_table(path)
        n, subs = do_import_attendance(df_raw, up_id)
        preview = ", ".join(subs[:8]) + (" …" if len(subs)>8 else "")
        flash(f'Attendance imported: {n} rows across {len(subs)} subjects{(" — " + preview) if subs else ""}.','success')
    except Exception as e:
        traceback.print_exc(); flash(f'Import failed (Attendance): {e}','error')
    return redirect(url_for('admin_imports'))

@app.route('/admin/upload/marks', methods=['POST'])
def upload_marks():
    if session.get('role')!='admin': return redirect(url_for('admin_login'))
    f = request.files.get('marks_file')
    if not f or not f.filename:
        flash('Choose a Marks Excel/CSV file.','error'); return redirect(url_for('admin_imports'))
    try:
        path, name = save_upload(f)
        up_id = create_upload_record(name, 'Marks')
        df_raw = load_table(path)
        n, subs, exams = do_import_marks(df_raw, up_id)
        preview = ", ".join(subs[:8]) + (" …" if len(subs)>8 else "")
        flash(f'Marks imported: {n} rows across {len(subs)} subjects, exams={", ".join(exams)}{(" — " + preview) if subs else ""}.','success')
    except Exception as e:
        traceback.print_exc(); flash(f'Import failed (Marks): {e}','error')
    return redirect(url_for('admin_imports'))

@app.route('/admin/uploads/<int:upload_id>/delete', methods=['POST'])
def delete_upload_route(upload_id):
    if session.get('role')!='admin': return redirect(url_for('admin_login'))
    ok = delete_upload(upload_id)
    flash('Upload deleted (file + data rolled back).' if ok else 'Upload not found.',
         'success' if ok else 'error')
    return redirect(url_for('admin_imports'))

# --- Admin: delete students tools (MISSING endpoints fixed here) ---
@app.route('/admin/students/clear', methods=['POST'], endpoint='admin_clear_students')
def admin_clear_students():
    if session.get('role') != 'admin':
        return redirect(url_for('admin_login'))
    try:
        with get_db() as con:
            con.execute('DELETE FROM attendance')
            con.execute('DELETE FROM marks')
            con.execute('DELETE FROM remarks')
            con.execute('DELETE FROM students')
            con.commit()
        flash('All students and related data removed.','success')
    except Exception as e:
        traceback.print_exc()
        flash(f'Failed to clear students: {e}','error')
    return redirect(url_for('admin_imports'))

@app.route('/admin/students/delete_one', methods=['POST'], endpoint='admin_delete_one_student')
def admin_delete_one_student():
    if session.get('role') != 'admin':
        return redirect(url_for('admin_login'))
    key = (request.form.get('roll_no') or '').strip()
    if not key:
        flash('Enter a roll/regn number.','error')
        return redirect(url_for('admin_imports'))
    try:
        key = _clean_id(key)
        with get_db() as con:
            cur = con.cursor()
            s = cur.execute('SELECT * FROM students WHERE roll_no=? OR regn_no=?',(key, key)).fetchone()
            if not s:
                flash(f'Student not found for "{key}".','error')
                return redirect(url_for('admin_imports'))
            alt = {s['roll_no']}
            if s['regn_no']: alt.add(s['regn_no'])
            siblings = cur.execute(
                "SELECT roll_no FROM students WHERE regn_no=? OR roll_no=?",
                (s['roll_no'], s['regn_no'] or s['roll_no'])
            ).fetchall()
            for rr in siblings: alt.add(rr['roll_no'])
            q = ",".join("?" * len(alt))
            cur.execute(f'DELETE FROM attendance WHERE roll_no IN ({q})', tuple(alt))
            cur.execute(f'DELETE FROM marks WHERE roll_no IN ({q})', tuple(alt))
            cur.execute(f'DELETE FROM remarks WHERE roll_no IN ({q})', tuple(alt))
            cur.execute(f'DELETE FROM students WHERE roll_no IN ({q})', tuple(alt))
            con.commit()
        flash(f'Removed student {s["roll_no"]} and linked IDs ({", ".join(sorted(alt))}).','success')
    except Exception as e:
        traceback.print_exc()
        flash(f'Failed to remove {key}: {e}','error')
    return redirect(url_for('admin_imports'))

# PDF (detailed report)
@app.route('/student/<roll>/pdf')
def student_pdf(roll):
    key = _clean_id(roll)
    with get_db() as con:
        cur = con.cursor()
        student = cur.execute('SELECT * FROM students WHERE roll_no=? OR regn_no=?', (key, key)).fetchone()
        if not student:
            buf = io.BytesIO()
            c = canvas.Canvas(buf, pagesize=A4)
            c.drawString(100, 800, f"Student not found: {roll}")
            c.showPage(); c.save(); buf.seek(0)
            return send_file(buf, as_attachment=True, download_name=f'{roll}_report.pdf')

    alt_rolls = {student['roll_no']}
    if student['regn_no']: alt_rolls.add(student['regn_no'])
    with get_db() as con:
        for rr in con.execute(
            "SELECT roll_no FROM students WHERE regn_no=? OR roll_no=?",
            (student['roll_no'], student['regn_no'] or student['roll_no'])
        ).fetchall():
            alt_rolls.add(rr['roll_no'])
    rolls_list = sorted(alt_rolls)

    marks_headers, marks_rows = build_marks_grid(rolls_list)
    att_headers, att_rows   = build_attendance_grid(rolls_list)

    cols_marks = len(marks_headers) if marks_headers else 0
    cols_att   = len(att_headers) if att_headers else 0
    use_landscape = max(cols_marks, cols_att) > 7
    pagesize = landscape(A4) if use_landscape else A4
    W, H = pagesize
    x_margin = 36
    y = H - 54
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=pagesize)

    c.setFont("Helvetica-Bold", 16); c.drawString(x_margin, y, "MVN UNIVERSITY"); y -= 18
    c.setFont("Helvetica", 11); c.drawString(x_margin, y, "Department of Computer Science & Engineering"); y -= 22
    c.setFont("Helvetica-Bold", 13); c.drawString(x_margin, y, "Student Report"); y -= 14
    c.setFont("Helvetica", 10); c.drawString(x_margin, y, f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"); y -= 20

    c.setFont("Helvetica-Bold", 11); c.drawString(x_margin, y, "Student Details"); y -= 10
    c.setLineWidth(0.3); c.line(x_margin, y, W - x_margin, y); y -= 8
    c.setFont("Helvetica", 10)
    details = [
        f"Name: {student['name']}", f"Roll No: {student['roll_no']}",
        f"Regn No: {student['regn_no'] or '-'}", f"Phone: {student['phone'] or '-'}",
        f"Email: {student['email'] or '-'}", f"Father: {student['father_name'] or '-'}",
        f"Father Phone: {student['father_phone'] or '-'}", f"Address: {student['address'] or '-'}",
        f"Session: {student['session'] or '-'}"
    ]
    for i in range(0, len(details), 2):
        left = details[i]; right = details[i+1] if i+1 < len(details) else ""
        c.drawString(x_margin, y, left)
        if right: c.drawString(W/2, y, right)
        y -= 14

    def draw_table(title, data):
        nonlocal y
        if not data or not data[0]: return
        c.setFont("Helvetica-Bold", 11); c.drawString(x_margin, y, title); y -= 8
        c.line(x_margin, y, W - x_margin, y); y -= 8
        hdr_font = 9 if not use_landscape else 8
        body_font = 8 if not use_landscape else 7
        avail = W - 2 * x_margin
        col_widths = [avail / len(data[0])] * len(data[0])
        table = Table(data, colWidths=col_widths, repeatRows=1)
        table.setStyle(TableStyle([
            ('FONT', (0,0), (-1,0), 'Helvetica-Bold', hdr_font),
            ('FONT', (0,1), (-1,-1), 'Helvetica', body_font),
            ('BACKGROUND', (0,0), (-1,0), colors.whitesmoke),
            ('GRID', (0,0), (-1,-1), 0.3, colors.lightgrey),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.Color(0.97,0.99,1.0)]),
        ]))
        tw, th = table.wrapOn(c, avail, y)
        if y - th < 72:
            c.showPage(); y = H - 54
        table.drawOn(c, x_margin, y - th)
        y -= (th + 16)

    def chunk_and_draw(title, headers, rows, fixed_first_col=True, max_cols=8):
        fixed = 1 if fixed_first_col else 0
        total = len(headers)
        if total <= fixed + max_cols:
            data = [headers] + [[r.get(h, "") for h in headers] for r in rows]
            draw_table(title, data); return
        part = 1
        for i in range(fixed, total, max_cols):
            subset = headers[:fixed] + headers[i:i+max_cols]
            data = [subset] + [[r.get(h, "") for h in subset] for r in rows]
            draw_table(f"{title} (Part {part})", data)
            part += 1

    if marks_headers and marks_rows:
        chunk_and_draw("Marks (By Exam)", marks_headers, marks_rows, fixed_first_col=True, max_cols=8)
    if att_headers and att_rows:
        chunk_and_draw("Attendance (By Subject)", att_headers, att_rows, fixed_first_col=True, max_cols=8)

    c.showPage(); c.save(); buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=f'{student["roll_no"]}_report.pdf')

# --------------- Entry ---------------
if __name__ == '__main__':
    run_schema()
    ensure_admin()
    app.run(debug=True, use_reloader=False, threaded=False)

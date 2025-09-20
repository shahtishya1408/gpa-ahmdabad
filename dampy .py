"""
advanced_attendance_singlefile.py
Single-file Face Recognition Attendance System (advanced features)

Features:
- SQLite backed students + attendance tables
- Auto-restore daily attendance (pre-fill "Absent" rows for the day)
- Late marking by threshold time (configurable)
- Departments / class filter
- Tkinter Treeview Attendance dashboard
- Export to Excel (via pandas if present) or CSV fallback
- Export to PDF (via reportlab if present)
- Modern UI with ttkbootstrap (optional; falls back to ttk)
- Sidebar navigation & theme toggle
- Basic email alerts for absentees (sends at scheduled time via tkinter.after)
- Simple face registration + recognition hooks (uses OpenCV LBPH if available)
All in one file.
"""

import os
import sqlite3
import csv
from datetime import datetime, date, time as dtime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import smtplib
from email.message import EmailMessage

# Optional imports
try:
    import ttkbootstrap as tb
    from ttkbootstrap.constants import *
    TB_AVAILABLE = True
except Exception:
    TB_AVAILABLE = False

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except Exception:
    PANDAS_AVAILABLE = False

try:
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfgen import canvas
    REPORTLAB_AVAILABLE = True
except Exception:
    REPORTLAB_AVAILABLE = False

# Face recognition optional
try:
    import cv2
    import numpy as np
    CV2_AVAILABLE = True
except Exception:
    CV2_AVAILABLE = False

APP_DB = "attendance.db"
DATASET_DIR = "face_dataset"
MODEL_FILE = "trainer.yml"

os.makedirs(DATASET_DIR, exist_ok=True)

# ---------- Database helpers ----------
def get_db_connection():
    conn = sqlite3.connect(APP_DB)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    cur = conn.cursor()
    # students: id (auto), name, roll, department, email
    cur.execute("""
        CREATE TABLE IF NOT EXISTS students(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            roll TEXT,
            department TEXT,
            email TEXT
        )
    """)
    # attendance: id, student_id, date (YYYY-MM-DD), time (HH:MM:SS), status (Present/Late/Absent)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS attendance(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER,
            date TEXT,
            time TEXT,
            status TEXT,
            FOREIGN KEY(student_id) REFERENCES students(id)
        )
    """)
    # config table for settings
    cur.execute("""
        CREATE TABLE IF NOT EXISTS config(
            key TEXT PRIMARY KEY,
            value TEXT
        )
    """)
    # Default config values
    defaults = {
        "late_threshold": "11:30:00",   # HH:MM:SS -> after this considered Late
        "absent_notify_time": "18:00:00",  # time to send absent emails
        "email_sender": "vipan49634@kwifa.com ", # hear we have tamp emali 
        "email_password": "",
        "smtp_server": "smtp.gmail.com",
        "smtp_port": "587",
        "auto_send_email": "0",
        "last_run_date": "2025-09-15"  # to track daily auto-restore
    }
    for k, v in defaults.items():
        cur.execute("INSERT OR IGNORE INTO config(key, value) VALUES (?, ?)", (k, v))
    conn.commit()
    conn.close()

def get_config(key):
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT value FROM config WHERE key=?", (key,))
    row = cur.fetchone()
    conn.close()
    return row["value"] if row else None

def set_config(key, value):
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("INSERT OR REPLACE INTO config(key, value) VALUES (?, ?)", (key, str(value)))
    conn.commit()
    conn.close()

# ---------- Auto-restore daily attendance ----------
def auto_restore_today():
    """For each student, ensure an attendance row exists for today;
       default to Absent with empty time."""
    today = date.today().isoformat()
    conn = get_db_connection()
    cur = conn.cursor()
    # find students without attendance today
    cur.execute("""
        SELECT id FROM students WHERE id NOT IN (
            SELECT student_id FROM attendance WHERE date=?
        )
    """, (today,))
    missing = cur.fetchall()
    for r in missing:
        cur.execute("INSERT INTO attendance(student_id, date, time, status) VALUES (?, ?, ?, ?)",
                    (r["id"], today, "", "Absent"))
    conn.commit()
    conn.close()

# ---------- Student management helpers ----------
def add_student(name, roll, department, email):
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("INSERT INTO students(name, roll, department, email) VALUES (?, ?, ?, ?)",
                (name, roll, department, email))
    conn.commit()
    student_id = cur.lastrowid
    conn.close()
    # ensure today's attendance exists for this new student
    auto_restore_today()
    return student_id

def get_students(dept_filter=None):
    conn = get_db_connection()
    cur = conn.cursor()
    if dept_filter and dept_filter != "All":
        cur.execute("SELECT * FROM students WHERE department=?", (dept_filter,))
    else:
        cur.execute("SELECT * FROM students")
    rows = cur.fetchall()
    conn.close()
    return rows

def get_departments():
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT DISTINCT department FROM students")
    rows = cur.fetchall()
    conn.close()
    depts = [r["department"] for r in rows if r["department"]]
    depts = sorted(list(set(depts)))
    return ["All"] + depts

# ---------- Attendance helpers ----------
def mark_attendance_db(student_id, mark_time=None):
    """Mark Present or Late in DB for today for student_id. If already present, update if needed."""
    if mark_time is None:
        mark_time = datetime.now().time()
    today = date.today().isoformat()
    cur_time = mark_time.strftime("%H:%M:%S")
    late_thresh = get_config("late_threshold") or "09:15:00"
    is_late = cur_time > late_thresh
    status = "Late" if is_late else "Present"
    conn = get_db_connection()
    cur = conn.cursor()
    # If an attendance row exists for today, update; else insert
    cur.execute("SELECT id, status FROM attendance WHERE student_id=? AND date=?", (student_id, today))
    row = cur.fetchone()
    if row:
        cur.execute("UPDATE attendance SET time=?, status=? WHERE id=?", (cur_time, status, row["id"]))
    else:
        cur.execute("INSERT INTO attendance(student_id, date, time, status) VALUES (?, ?, ?, ?)",
                    (student_id, today, cur_time, status))
    conn.commit()
    conn.close()
    return status

def get_attendance_rows(date_filter=None, dept_filter=None):
    conn = get_db_connection()
    cur = conn.cursor()
    q = """
        SELECT a.id, a.student_id, s.name, s.roll, s.department, a.date, a.time, a.status, s.email
        FROM attendance a
        JOIN students s ON a.student_id = s.id
    """
    params = []
    conds = []
    if date_filter:
        conds.append("a.date=?")
        params.append(date_filter)
    if dept_filter and dept_filter != "All":
        conds.append("s.department=?")
        params.append(dept_filter)
    if conds:
        q += " WHERE " + " AND ".join(conds)
        q += " ORDER BY s.department, s.name"
        cur.execute(q, tuple(params))
    rows = cur.fetchall()
    conn.close()
    return rows
def upgrade_db():
    conn = get_db_connection()
    cur = conn.cursor()
    # check columns
    cur.execute("PRAGMA table_info(attendance)")
    cols = [r[1] for r in cur.fetchall()]
    if "status" not in cols:
        cur.execute("ALTER TABLE attendance ADD COLUMN status TEXT DEFAULT 'Absent'")
    cur.execute("PRAGMA table_info(students)")
    cols = [r[1] for r in cur.fetchall()]
    if "email" not in cols:
        cur.execute("ALTER TABLE students ADD COLUMN email TEXT")
    conn.commit()
    conn.close()

# Call after init_db()
init_db()
upgrade_db()

# ---------- Exports ----------
def export_to_csv(filename, rows):
    header = ["id", "student_id", "name", "roll", "department", "date", "time", "status", "email"]
    with open(filename, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(header)
        for r in rows:
            w.writerow([r["id"], r["student_id"], r["name"], r["roll"], r["department"], r["date"], r["time"], r["status"], r["email"]])

def export_to_excel(filename, rows):
    if not PANDAS_AVAILABLE:
        # fallback to CSV but use .xlsx extension won't be a real Excel - warn user
        export_to_csv(filename.replace(".xlsx", ".csv"), rows)
        return False
    df = pd.DataFrame([dict(r) for r in rows])
    df.to_excel(filename, index=False)
    return True

def export_to_pdf(filename, rows, title="Attendance Report"):
    if not REPORTLAB_AVAILABLE:
        return False
    c = canvas.Canvas(filename, pagesize=letter)
    width, height = letter
    c.setFont("Helvetica-Bold", 14)
    c.drawString(40, height - 40, title)
    c.setFont("Helvetica", 10)
    y = height - 70
    header = ["Name", "Roll", "Dept", "Date", "Time", "Status"]
    x_positions = [40, 200, 260, 350, 430, 500]
    for i, h in enumerate(header):
        c.drawString(x_positions[i], y, h)
    y -= 18
    for r in rows:
        if y < 60:
            c.showPage()
            y = height - 40
        c.drawString(x_positions[0], y, str(r["name"]))
        c.drawString(x_positions[1], y, str(r["roll"] or ""))
        c.drawString(x_positions[2], y, str(r["department"] or ""))
        c.drawString(x_positions[3], y, str(r["date"]))
        c.drawString(x_positions[4], y, str(r["time"] or ""))
        c.drawString(x_positions[5], y, str(r["status"]))
        y -= 16
    c.save()
    return True

# ---------- Email Alerts ----------
def send_absent_emails(send_time_str=None):
    """Send email alerts for students still Absent for today."""
    today = date.today().isoformat()
    rows = get_attendance_rows(date_filter=today, dept_filter=None)
    absents = [r for r in rows if r["status"] == "Absent" and r["email"]]
    if not absents:
        return {"sent": 0, "skipped": 0}

    smtp_server = get_config("smtp_server") or "smtp.gmail.com"
    smtp_port = int(get_config("smtp_port") or 587)
    sender = get_config("email_sender") or ""
    password = get_config("email_password") or ""
    if not sender or not password:
        return {"error": "Email sender/password not configured."}

    sent = 0
    skipped = 0
    try:
        server = smtplib.SMTP(smtp_server, smtp_port, timeout=10)
        server.starttls()
        server.login(sender, password)
    except Exception as e:
        return {"error": f"SMTP login failed: {e}"}

    for r in absents:
        try:
            msg = EmailMessage()
            msg["Subject"] = f"Absent Notice - {r['name']}"
            msg["From"] = sender
            msg["To"] = r["email"]
            body = f"Dear {r['name']},\n\nOur records show you are marked Absent today ({today}). If this is incorrect, please contact administration.\n\nRegards,\nAttendance System"
            msg.set_content(body)
            server.send_message(msg)
            sent += 1
        except Exception:
            skipped += 1
    server.quit()
    return {"sent": sent, "skipped": skipped}

# ---------- Face Registration & Recognition (minimal LBPH hooks) ----------
if CV2_AVAILABLE:
    face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + "haarcascade_frontalface_default.xml")
    def capture_face_samples_for(name, samples=5):
        user_dir = os.path.join(DATASET_DIR, name)
        os.makedirs(user_dir, exist_ok=True)
        cam = cv2.VideoCapture(0)
        if not cam.isOpened():
            raise RuntimeError("Camera not available")
        count = 0
        while True:
            ret, frame = cam.read()
            if not ret:
                break
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            faces = face_cascade.detectMultiScale(gray, 1.3, 5)
            for (x,y,w,h) in faces:
                count += 1
                face_img = gray[y:y+h, x:x+w]
                face_img = cv2.resize(face_img, (200, 200))
                cv2.imwrite(os.path.join(user_dir, f"{count}.jpg"), face_img)
                cv2.rectangle(frame, (x,y), (x+w,y+h), (255,0,0), 2)
            cv2.imshow("Capture Faces", frame)
            if cv2.waitKey(1) & 0xFF == ord('q'):
                break
            if count >= samples:
                break
        cam.release()
        cv2.destroyAllWindows()
        return count

    def train_lbph_model():
        recognizer = cv2.face.LBPHFaceRecognizer_create()
        faces = []
        labels = []
        label_map = {}
        cur_label = 0
        for person_name in sorted(os.listdir(DATASET_DIR)):
            person_dir = os.path.join(DATASET_DIR, person_name)
            if not os.path.isdir(person_dir):
                continue
            label_map[cur_label] = person_name
            for f in os.listdir(person_dir):
                img = cv2.imread(os.path.join(person_dir, f), cv2.IMREAD_GRAYSCALE)
                if img is None:
                    continue
                faces.append(img)
                labels.append(cur_label)
            cur_label += 1
        if not faces:
            raise RuntimeError("No faces found to train")
        recognizer.train(faces, np.array(labels, dtype=np.int32))
        recognizer.write(MODEL_FILE)
        # save mapping as simple text
        with open("label_map.csv", "w", newline="", encoding="utf-8") as fh:
            w = csv.writer(fh)
            for k, v in label_map.items():
                w.writerow([k, v])
        return len(label_map)

    def load_label_map():
        mapping = {}
        if os.path.exists("label_map.csv"):
            with open("label_map.csv", newline="", encoding="utf-8") as fh:
                r = csv.reader(fh)
                for row in r:
                    if len(row) >= 2:
                        mapping[int(row[0])] = row[1]
        return mapping

    def recognize_and_mark(conf_threshold=70):
        """Runs camera until recognizes a student and marks attendance; returns (name, status) or None."""
        if not os.path.exists(MODEL_FILE):
            raise RuntimeError("Model file not trained yet.")
        recognizer = cv2.face.LBPHFaceRecognizer_create()
        recognizer.read(MODEL_FILE)
        label_map = load_label_map()
        cam = cv2.VideoCapture(0)
        if not cam.isOpened():
            raise RuntimeError("Camera not available")
        start = datetime.now().timestamp()
        recognized = None
        while True:
            ret, frame = cam.read()
            if not ret:
                break
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            faces = face_cascade.detectMultiScale(gray, 1.3, 5)
            for (x,y,w,h) in faces:
                roi = cv2.resize(gray[y:y+h, x:x+w], (200, 200))
                label, conf = recognizer.predict(roi)
                if conf < conf_threshold:
                    name = label_map.get(label, "Unknown")
                    # find student ID by name
                    conn = get_db_connection()
                    cur = conn.cursor()
                    cur.execute("SELECT id FROM students WHERE name=?", (name,))
                    r = cur.fetchone()
                    conn.close()
                    if r:
                        status = mark_attendance_db(r["id"], datetime.now().time())
                        cam.release()
                        cv2.destroyAllWindows()
                        return name, status
            if datetime.now().timestamp() - start > 20:  # timeout 20s
                break
        cam.release()
        cv2.destroyAllWindows()
        return None

else:
    # stubs if OpenCV not installed
    def capture_face_samples_for(name, samples=5):
        raise RuntimeError("OpenCV not available. Install opencv-python to use face features.")
    def train_lbph_model():
        raise RuntimeError("OpenCV not available.")
    def recognize_and_mark(conf_threshold=70):
        raise RuntimeError("OpenCV not available.")

# ---------- GUI ----------
class AttendanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Advanced Attendance System (Single File)")
        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        root.geometry(f"{int(sw*0.9)}x{int(sh*0.85)}+20+20")

        # Use ttkbootstrap if available
        if TB_AVAILABLE:
            style_theme = "darkly" if get_config("theme") == "dark" else "flatly"
            self.style = tb.Style(theme=style_theme)
            self.base = tb.Frame(root)
            self.base.pack(fill="both", expand=True)
        else:
            self.style = ttk.Style()
            self.base = ttk.Frame(root)
            self.base.pack(fill="both", expand=True)

        # Left sidebar
        self.sidebar = ttk.Frame(self.base, width=220)
        self.sidebar.pack(side="left", fill="y")
        # Content frame
        self.content = ttk.Frame(self.base)
        self.content.pack(side="right", fill="both", expand=True)

        # Sidebar buttons
        btns = [
            ("Dashboard", self.show_dashboard),
            ("Students", self.show_students),
            ("Attendance", self.show_attendance),
            ("Reports", self.show_reports),
            ("Settings", self.show_settings),
        ]
        for i, (t, cmd) in enumerate(btns):
            b = ttk.Button(self.sidebar, text=t, command=cmd)
            b.pack(fill="x", padx=8, pady=(8 if i==0 else 4))

        # Theme toggle
        self.theme_var = tk.StringVar(value=get_config("theme") or "light")
        tbtn = ttk.Button(self.sidebar, text="Toggle Theme", command=self.toggle_theme)
        tbtn.pack(fill="x", padx=8, pady=8)

        # Quick Face buttons
        ttk.Label(self.sidebar, text="Face Actions:", font=("Arial", 10, "bold")).pack(pady=(16,4))
        ttk.Button(self.sidebar, text="Capture Face (selected student)", command=self.capture_face_for_selected).pack(fill="x", padx=8, pady=4)
        ttk.Button(self.sidebar, text="Train Model", command=self.train_model).pack(fill="x", padx=8, pady=4)
        ttk.Button(self.sidebar, text="Start Recognition (mark)", command=self.start_recognition_thread).pack(fill="x", padx=8, pady=4)

        # Initialize frames for pages
        self.frames = {}
        for name in ("dashboard", "students", "attendance", "reports", "settings"):
            f = ttk.Frame(self.content)
            self.frames[name] = f

        # Start with dashboard
        self.show_dashboard()

        # Do daily auto-restore
        auto_restore_today()

        # Schedule email auto-send if enabled
        try:
            if get_config("auto_send_email") == "1":
                self.schedule_email_at_config_time()
        except Exception:
            pass

    # ---------- Page: Dashboard ----------
    def show_dashboard(self):
        self._show_frame("dashboard")
        f = self.frames["dashboard"]
        for widget in f.winfo_children(): widget.destroy()
        ttk.Label(f, text="Dashboard", font=("Arial", 16, "bold")).pack(pady=8)
        # Summary stats for today
        today = date.today().isoformat()
        rows = get_attendance_rows(date_filter=today)
        total = len(rows)
        present = len([r for r in rows if r["status"] in ("Present", "Late")])
        absent = len([r for r in rows if r["status"] == "Absent"])
        percent = (present/total*100) if total else 0
        ttk.Label(f, text=f"Date: {today}").pack()
        ttk.Label(f, text=f"Total students: {total}").pack()
        ttk.Label(f, text=f"Present: {present}").pack()
        ttk.Label(f, text=f"Absent: {absent}").pack()
        ttk.Label(f, text=f"Attendance %: {percent:.2f}%").pack()
        # buttons quick
        bframe = ttk.Frame(f)
        bframe.pack(pady=10)
        ttk.Button(bframe, text="View Attendance Page", command=self.show_attendance).pack(side="left", padx=6)
        ttk.Button(bframe, text="Send Absent Emails Now", command=lambda: self.send_absent_emails_and_notify_ui()).pack(side="left", padx=6)

    # ---------- Page: Students ----------
    def show_students(self):
        self._show_frame("students")
        f = self.frames["students"]
        for widget in f.winfo_children(): widget.destroy()
        ttk.Label(f, text="Students", font=("Arial", 14, "bold")).pack(pady=8)

        form = ttk.Frame(f)
        form.pack(pady=6, padx=10, fill="x")
        ttk.Label(form, text="Name").grid(row=0, column=0, sticky="w")
        name_e = ttk.Entry(form); name_e.grid(row=0, column=1, sticky="ew")
        ttk.Label(form, text="Roll").grid(row=1, column=0, sticky="w")
        roll_e = ttk.Entry(form); roll_e.grid(row=1, column=1, sticky="ew")
        ttk.Label(form, text="Department").grid(row=2, column=0, sticky="w")
        dept_e = ttk.Entry(form); dept_e.grid(row=2, column=1, sticky="ew")
        ttk.Label(form, text="Email").grid(row=3, column=0, sticky="w")
        email_e = ttk.Entry(form); email_e.grid(row=3, column=1, sticky="ew")
        form.columnconfigure(1, weight=1)
        def add():
            name = name_e.get().strip()
            if not name:
                messagebox.showerror("Error", "Name required")
                return
            add_student(name, roll_e.get().strip(), dept_e.get().strip(), email_e.get().strip())
            messagebox.showinfo("Added", f"Student {name} added")
            name_e.delete(0, "end"); roll_e.delete(0,"end"); dept_e.delete(0,"end"); email_e.delete(0,"end")
            self.show_students()
        ttk.Button(form, text="Add Student", command=add).grid(row=4, column=0, columnspan=2, pady=8)

        # list
        cols = ("id","name","roll","department","email")
        tree = ttk.Treeview(f, columns=cols, show="headings", height=12)
        for c in cols:
            tree.heading(c, text=c.title())
            tree.column(c, width=120)
        tree.pack(fill="both", expand=True, padx=10, pady=6)
        for s in get_students():
            tree.insert("", "end", values=(s["id"], s["name"], s["roll"], s["department"], s["email"]))
        self.student_tree = tree

    # ---------- Page: Attendance ----------
    def show_attendance(self):
        self._show_frame("attendance")
        f = self.frames["attendance"]
        for widget in f.winfo_children(): widget.destroy()
        ttk.Label(f, text="Attendance", font=("Arial", 14, "bold")).pack(pady=8)
        ctrl = ttk.Frame(f); ctrl.pack(fill="x", padx=10)
        ttk.Label(ctrl, text="Date (YYYY-MM-DD)").pack(side="left")
        date_e = ttk.Entry(ctrl); date_e.pack(side="left", padx=6); date_e.insert(0, date.today().isoformat())
        ttk.Label(ctrl, text="Department").pack(side="left", padx=(10,0))
        dept_cb = ttk.Combobox(ctrl, values=get_departments(), state="readonly"); dept_cb.pack(side="left", padx=6); dept_cb.set("All")

        def refresh():
            dt = date_e.get().strip()
            dept = dept_cb.get()
            try:
                rows = get_attendance_rows(date_filter=dt, dept_filter=dept)
            except Exception as e:
                messagebox.showerror("Error", str(e)); return
            for r in tree.get_children():
                tree.delete(r)
            for r in rows:
                tree.insert("", "end", values=(r["id"], r["student_id"], r["name"], r["roll"], r["department"], r["date"], r["time"], r["status"]))
            # stats
            total = len(rows)
            present = len([r for r in rows if r["status"] in ("Present","Late")])
            absent = len([r for r in rows if r["status"]=="Absent"])
            stats_var.set(f"Total {total}  Present {present}  Absent {absent}  % {(present/total*100) if total else 0:.2f}")
        ttk.Button(ctrl, text="Refresh", command=refresh).pack(side="left", padx=6)
        stats_var = tk.StringVar(value="")
        ttk.Label(f, textvariable=stats_var).pack()

        cols = ("id","student_id","name","roll","department","date","time","status")
        tree = ttk.Treeview(f, columns=cols, show="headings", height=14)
        for c in cols:
            tree.heading(c, text=c.title())
            tree.column(c, width=110)
        tree.pack(fill="both", expand=True, padx=10, pady=8)

        # right-click / actions
        def on_mark_present():
            sel = tree.selection()
            if not sel:
                messagebox.showerror("Error", "Select a row")
                return
            row = tree.item(sel[0])["values"]
            student_id = row[1]
            status = mark_attendance_db(student_id)
            messagebox.showinfo("Marked", f"Marked {status} for student id {student_id}")
            refresh()
        def on_export():
            dt = date_e.get().strip()
            dept = dept_cb.get()
            rows = get_attendance_rows(date_filter=dt, dept_filter=dept)
            if not rows:
                messagebox.showinfo("No data", "No attendance rows to export")
                return
            dest = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV","*.csv"),("Excel","*.xlsx"),("PDF","*.pdf")])
            if not dest:
                return
            if dest.endswith(".csv"):
                export_to_csv(dest, rows)
                messagebox.showinfo("Exported", f"Exported to {dest}")
            elif dest.endswith(".xlsx"):
                ok = export_to_excel(dest, rows)
                if ok:
                    messagebox.showinfo("Exported", f"Exported to {dest}")
                else:
                    messagebox.showwarning("Partial", "pandas not installed; exported CSV fallback")
            elif dest.endswith(".pdf"):
                ok = export_to_pdf(dest, rows)
                if ok:
                    messagebox.showinfo("Exported", f"Exported to {dest}")
                else:
                    messagebox.showwarning("PDF missing", "reportlab not installed; cannot export PDF")
        action_frame = ttk.Frame(f)
        action_frame.pack(pady=6)
        ttk.Button(action_frame, text="Mark Present (selected)", command=on_mark_present).pack(side="left", padx=6)
        ttk.Button(action_frame, text="Export", command=on_export).pack(side="left", padx=6)

        # initial refresh
        refresh()

    # ---------- Page: Reports ----------
    def show_reports(self):
        self._show_frame("reports")
        f = self.frames["reports"]
        for widget in f.winfo_children(): widget.destroy()
        ttk.Label(f, text="Reports", font=("Arial", 14, "bold")).pack(pady=8)
        ttk.Label(f, text="Select Date").pack()
        date_e = ttk.Entry(f); date_e.pack(); date_e.insert(0, date.today().isoformat())
        dept_cb = ttk.Combobox(f, values=get_departments(), state="readonly"); dept_cb.pack(); dept_cb.set("All")
        def gen_stats():
            dt = date_e.get().strip()
            dept = dept_cb.get()
            rows = get_attendance_rows(date_filter=dt, dept_filter=dept)
            total = len(rows)
            present = len([r for r in rows if r["status"] in ("Present","Late")])
            absent = len([r for r in rows if r["status"]=="Absent"])
            per = (present/total*100) if total else 0
            msg = f"Date {dt}\nTotal: {total}\nPresent: {present}\nAbsent: {absent}\nAttendance %: {per:.2f}"
            messagebox.showinfo("Report", msg)
        ttk.Button(f, text="Generate Summary", command=gen_stats).pack(pady=6)
        ttk.Button(f, text="Export Summary CSV", command=lambda: self._export_summary(date_e.get().strip(), dept_cb.get())).pack(pady=6)

    def _export_summary(self, dt, dept):
        rows = get_attendance_rows(date_filter=dt, dept_filter=dept)
        if not rows:
            messagebox.showinfo("No data", "No rows to export")
            return
        dest = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV","*.csv")])
        if not dest:
            return
        export_to_csv(dest, rows)
        messagebox.showinfo("Exported", dest)

    # ---------- Page: Settings ----------
    def show_settings(self):
        self._show_frame("settings")
        f = self.frames["settings"]
        for widget in f.winfo_children(): widget.destroy()
        ttk.Label(f, text="Settings", font=("Arial", 14, "bold")).pack(pady=8)
        # late threshold
        ttk.Label(f, text="Late threshold (HH:MM:SS)").pack()
        late_e = ttk.Entry(f); late_e.pack(); late_e.insert(0, get_config("late_threshold") or "09:15:00")
        ttk.Label(f, text="Absent email send time (HH:MM:SS)").pack()
        send_e = ttk.Entry(f); send_e.pack(); send_e.insert(0, get_config("absent_notify_time") or "18:00:00")
        # email settings
        ttk.Label(f, text="SMTP Server").pack()
        smtp_e = ttk.Entry(f); smtp_e.pack(); smtp_e.insert(0, get_config("smtp_server") or "smtp.gmail.com")
        ttk.Label(f, text="SMTP Port").pack()
        port_e = ttk.Entry(f); port_e.pack(); port_e.insert(0, get_config("smtp_port") or "587")
        ttk.Label(f, text="Sender Email").pack()
        sender_e = ttk.Entry(f); sender_e.pack(); sender_e.insert(0, get_config("email_sender") or "")
        ttk.Label(f, text="Sender Password (app password recommended)").pack()
        pw_e = ttk.Entry(f, show="*"); pw_e.pack(); pw_e.insert(0, get_config("email_password") or "")
        auto_var = tk.IntVar(value=1 if get_config("auto_send_email") == "1" else 0)
        ttk.Checkbutton(f, text="Auto send absent emails daily at configured time", variable=auto_var).pack(pady=6)

        def save_settings():
            set_config("late_threshold", late_e.get().strip())
            set_config("absent_notify_time", send_e.get().strip())
            set_config("smtp_server", smtp_e.get().strip())
            set_config("smtp_port", port_e.get().strip())
            set_config("email_sender", sender_e.get().strip())
            set_config("email_password", pw_e.get().strip())
            set_config("auto_send_email", "1" if auto_var.get() else "0")
            messagebox.showinfo("Saved", "Settings saved")
            if auto_var.get():
                self.schedule_email_at_config_time()
        ttk.Button(f, text="Save Settings", command=save_settings).pack(pady=8)
        ttk.Button(f, text="Send Absent Emails Now", command=lambda: self.send_absent_emails_and_notify_ui()).pack(pady=4)

    # ---------- Utility and actions ----------
    def _show_frame(self, name):
        for key, fr in self.frames.items():
            fr.pack_forget()
        self.frames[name].pack(fill="both", expand=True)

    def toggle_theme(self):
        if TB_AVAILABLE:
            cur = self.style.theme.name
            new = "flatly" if "dark" in cur else "darkly"
            self.style.theme_use(new)
            set_config("theme", "dark" if "dark" in new else "light")
        else:
            messagebox.showinfo("Theme", "ttkbootstrap not installed; can't toggle advanced themes.")

    def capture_face_for_selected(self):
        # capture for the currently selected student in students page tree
        sel_id = None
        try:
            sel = self.student_tree.selection()
            if not sel:
                messagebox.showerror("Select", "Select a student in the Students page list first")
                return
            sel_id = self.student_tree.item(sel[0])["values"][0]
            name = self.student_tree.item(sel[0])["values"][1]
        except Exception:
            messagebox.showerror("Error", "Open Students page and select a student first"); return
        if not CV2_AVAILABLE:
            messagebox.showerror("OpenCV missing", "Install opencv-python to capture faces")
            return
        try:
            cnt = capture_face_samples_for(name)
            messagebox.showinfo("Captured", f"Captured {cnt} images for {name}")
        except Exception as e:
            messagebox.showerror("Capture error", str(e))

    def train_model(self):
        if not CV2_AVAILABLE:
            messagebox.showerror("Missing", "OpenCV not installed; cannot train")
            return
        try:
            n = train_lbph_model()
            messagebox.showinfo("Trained", f"Trained model with {n} people")
        except Exception as e:
            messagebox.showerror("Training error", str(e))

    def start_recognition_thread(self):
        # run recognition in a thread to avoid blocking UI
        def target():
            try:
                res = recognize_and_mark()
                if res:
                    name, status = res
                    messagebox.showinfo("Recognized", f"{name} marked as {status}")
                else:
                    messagebox.showinfo("Result", "No recognition or timed out")
            except Exception as e:
                messagebox.showerror("Recognition error", str(e))
        t = threading.Thread(target=target, daemon=True)
        t.start()

    def send_absent_emails_and_notify_ui(self):
        res = send_absent_emails()
        if "error" in res:
            messagebox.showerror("Email error", res["error"])
        else:
            messagebox.showinfo("Email report", f"Sent: {res.get('sent',0)}  Skipped: {res.get('skipped',0)}")

    def schedule_email_at_config_time(self):
        # Schedule daily call at configured time (only while app runs)
        time_str = get_config("absent_notify_time") or "18:00:00"
        try:
            h, m, s = [int(x) for x in time_str.split(":")]
        except Exception:
            h, m, s = (18, 0, 0)
        now = datetime.now()
        target = datetime(now.year, now.month, now.day, h, m, s)
        if target < now:
            # schedule for next day
            target = datetime(now.year, now.month, now.day, h, m, s)
            # if in past, add one day
            from datetime import timedelta
            target = target + timedelta(days=1)
        delay_ms = int((target - now).total_seconds() * 1000)
        # Cancel existing (no references stored) - simple approach: just call after()
        def scheduled_send():
            try:
                self.send_absent_emails_and_notify_ui()
            finally:
                # re-schedule for next day
                self.root.after(24*3600*1000, scheduled_send)
        self.root.after(delay_ms, scheduled_send)

# ---------- App init ----------


if __name__ == "__main__":
    init_db()
    # create default TK root; if ttkbootstrap available, use its theme root approach
    if TB_AVAILABLE:
        root = tb.Window(themename="flatly")
    else:
        root = tk.Tk()
    app = AttendanceApp(root)
    root.mainloop()
 
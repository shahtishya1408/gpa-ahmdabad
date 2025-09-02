import os
import csv
import cv2
import time
import numpy as np
from datetime import datetime
from tkinter import *
from tkinter import ttk, messagebox

# ------------- Paths and Model Setup -------------
DATASET_DIR = "face_dataset"
MODEL_PATH = "trainer.yml"
ATTENDANCE_CSV = "attendance.csv"

os.makedirs(DATASET_DIR, exist_ok=True)

# Load Haar Cascade for face detection
FACE_CASCADE = cv2.CascadeClassifier(cv2.data.haarcascades + "haarcascade_frontalface_default.xml")

# Copyright text
COPYRIGHT_TEXT = "\u00A9 2025 My Institute"  # e.g., © 2025 My Institute

# ------------- Attendance Helpers -------------
def ensure_attendance_csv():
    if not os.path.exists(ATTENDANCE_CSV):
        with open(ATTENDANCE_CSV, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["name", "timestamp", "status" ,"date"])

def mark_attendance(name, status="Present"):
    ensure_attendance_csv()
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    datetoday = datetime.now().strftime("%Y-%m-%d")
    with open(ATTENDANCE_CSV, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow([name, ts, status , datetoday])

# ------------- Student Form + Photo Samples -------------
class StudentApp(Frame):
    def __init__(self, master):
        super().__init__(master, bd=2, relief=RIDGE, padx=10, pady=10)
        self.pack(fill=BOTH, expand=True)
        self.master.title("Student Management + Face Samples")
        self.master.geometry("900x600")
        
        # Variables
        self.var_student_id = StringVar()
        self.var_name = StringVar()
        self.var_department = StringVar()
        self.var_roll = StringVar()
        self.var_gender = StringVar()

        title = Label(self, text="STUDENT MANAGEMENT SYSTEM", font=("Arial", 18, "bold"), bg="green", fg="white")
        title.pack(side=TOP, fill=X, pady=(0,10))

        form = Frame(self)
        form.pack(pady=10)

        def add_row(r, text, var=None, widget="entry", values=None):
            Label(form, text=text, font=("Arial", 12)).grid(row=r, column=0, padx=5, pady=6, sticky=W)
            if widget == "entry":
                e = Entry(form, textvariable=var, font=("Arial", 12))
                e.grid(row=r, column=1, padx=5, pady=6)
                return e
            elif widget == "combo":
                cb = ttk.Combobox(form, textvariable=var, font=("Arial", 12), state="readonly")
                cb["values"] = values or []
                cb.grid(row=r, column=1, padx=5, pady=6)
                return cb

        add_row(0, "Student ID", self.var_student_id)
        add_row(1, "Name", self.var_name)
        add_row(2, "Department", self.var_department)
        add_row(3, "Roll", self.var_roll)
        add_row(4, "Gender", self.var_gender, widget="combo", values=("Male", "Female", "Other"))

        btns = Frame(form)
        btns.grid(row=5, column=0, columnspan=2, pady=12)
        Button(btns, text="Clear", width=12, command=self.clear).grid(row=0, column=0, padx=5)
        Button(btns, text="Add Photo Sample", width=18, command=self.capture_face_samples).grid(row=0, column=1, padx=5)

        note = Label(self, text="Note: Photo samples save to face_dataset/<Name>/", fg="gray")
        note.pack()

    def clear(self):
        self.var_student_id.set("")
        self.var_name.set("")
        self.var_department.set("")
        self.var_roll.set("")
        self.var_gender.set("")

    

    def capture_face_samples(self):
        name = self.var_name.get().strip()
        if not name:
            messagebox.showerror("Error", "Enter Name before capturing photos")
            return

        user_dir = os.path.join(DATASET_DIR, name)
        os.makedirs(user_dir, exist_ok=True)

        cam = cv2.VideoCapture(0)
        if not cam.isOpened():
            messagebox.showerror("Error", "Cannot access camera")
            return

        messagebox.showinfo("Info", "Capturing faces automatically. Window will close when enough samples are saved.")
        count = 0
        while True:
            ret, frame = cam.read()
            if not ret:
                break
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            faces = FACE_CASCADE.detectMultiScale(gray, 1.3, 5)
            for (x,y,w,h) in faces:
                count += 1
                face_img = gray[y:y+h, x:x+w]
                face_img = cv2.resize(face_img, (200, 200))
                cv2.imwrite(os.path.join(user_dir, f"{count}.jpg"), face_img)
                cv2.rectangle(frame, (x,y), (x+w,y+h), (0,255,0), 2)
                cv2.putText(frame, f"Saved: {count}", (x, y-10), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0,255,0), 2)

            self._overlay_copyright(frame)
            cv2.imshow("Capturing Face Samples", frame)

            # HighGUI needs a wait call to refresh; 1 ms is fine. No key handling.
            cv2.waitKey(1)  # keeps window responsive and renders frames [9]

            if count >= 5:  # Limit to 5 samples
                break

        cam.release()
        cv2.destroyAllWindows()
        messagebox.showinfo("Success", f"Saved {count} images for {name}")

# ------------- Face Functions (Register, Train, Login/Attendance) -------------
class FaceFunctions:
    def register_face(self, name: str):
        name = name.strip()
        if not name:
            messagebox.showerror("Error", "Please enter a name before registering.")
            return
        user_dir = os.path.join(DATASET_DIR, name)
        os.makedirs(user_dir, exist_ok=True)

        cam = cv2.VideoCapture(0)
        if not cam.isOpened():
            messagebox.showerror("Error", "Cannot access camera")
            return

        messagebox.showinfo("Info", "Registering automatically. Window will close when enough samples are saved.")
        count = 0
        
        while True:
            ret, frame = cam.read()
            if not ret:
                break
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            faces = FACE_CASCADE.detectMultiScale(gray, 1.3, 5)
            for (x,y,w,h) in faces:
                count += 1
                face_img = gray[y:y+h, x:x+w]
                face_img = cv2.resize(face_img, (200, 200))
                cv2.imwrite(os.path.join(user_dir, f"{count}.jpg"), face_img)
                cv2.rectangle(frame, (x,y), (x+w,y+h), (255,0,0), 2)

          
            cv2.imshow("Register Face", frame)
            cv2.waitKey(1)  # render and keep responsive [9]

            if count >= 5:
                break

        cam.release()
        cv2.destroyAllWindows()
        messagebox.showinfo("Success", f"Registered {count} images for {name}")

    def train_model(self):
        faces, labels = [], []
        label_map = {}
        current_label = 0

        # Build dataset
        for person_name in sorted(os.listdir(DATASET_DIR)):
            person_dir = os.path.join(DATASET_DIR, person_name)
            if not os.path.isdir(person_dir):
                continue
            label_map[current_label] = person_name
            for img_file in os.listdir(person_dir):
                img_path = os.path.join(person_dir, img_file)
                img = cv2.imread(img_path, cv2.IMREAD_GRAYSCALE)
                if img is None:
                    continue
                faces.append(img)
                labels.append(current_label)
            current_label += 1

        if len(faces) == 0:
            messagebox.showerror("Error", "No data found. Please register faces first.")
            return

        recognizer = cv2.face.LBPHFaceRecognizer_create()
        recognizer.train(faces, np.array(labels, dtype=np.int32))
        recognizer.write(MODEL_PATH)
        messagebox.showinfo("Success", f"Model trained. People: {len(label_map)}")

    def face_login_and_mark(self, confidence_threshold=70.0, max_run_seconds=20):
        if not os.path.exists(MODEL_PATH):
            messagebox.showerror("Error", "Model not found. Please train first.")
            return

        recognizer = cv2.face.LBPHFaceRecognizer_create()
        recognizer.read(MODEL_PATH)

        # Ensure label map matches training order
        label_map = {i: name for i, name in enumerate(sorted(os.listdir(DATASET_DIR)))}

        cam = cv2.VideoCapture(0)
        if not cam.isOpened():
            messagebox.showerror("Error", "Cannot access camera")
            return

        start_time = time.time()
        recognized_name = None
        last_mark_time = 0

        # No keypress loop; auto-exit upon success or timeout
        while True:
            ret, frame = cam.read()
            if not ret:
                break

            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            faces = FACE_CASCADE.detectMultiScale(gray, 1.3, 5)
            for (x,y,w,h) in faces:
                roi = cv2.resize(gray[y:y+h, x:x+w], (200, 200))
                label, confidence = recognizer.predict(roi)
                if confidence < confidence_threshold:
                    recognized_name = label_map.get(label, "Unknown")
                    color = (0, 255, 0)
                    text = f"{recognized_name} ({confidence:.1f})"
                    now = time.time()
                    if recognized_name != "Unknown" and (now - last_mark_time) > 2:
                        mark_attendance(recognized_name, "Present")
                        last_mark_time = now
                        # Auto-stop after first successful mark
                        cv2.rectangle(frame, (x,y), (x+w,y+h), color, 2)
                        cv2.putText(frame, text, (x, y-10), cv2.FONT_HERSHEY_SIMPLEX, 0.7, color, 2)
                        
                        cv2.imshow("Face Attendance", frame)
                        cv2.waitKey(500)  # brief show of result before closing [1][9]
                        cam.release()
                        cv2.destroyAllWindows()
                        messagebox.showinfo("Attendance", f"Marked Present: {recognized_name}")
                        return
                else:
                    recognized_name = "Unknown"
                    color = (0, 0, 255)
                    text = f"Unknown ({confidence:.1f})"

                cv2.rectangle(frame, (x,y), (x+w,y+h), color, 2)
                cv2.putText(frame, text, (x, y-10), cv2.FONT_HERSHEY_SIMPLEX, 0.7, color, 2)
                break  # show one face at a time for clarity

            cv2.imshow("Face Attendance", frame)
            cv2.waitKey(1)  # render frames; no key handling [9]

            # Auto-timeout
            if time.time() - start_time > max_run_seconds:
                break

        cam.release()
        cv2.destroyAllWindows()
        messagebox.showinfo("Attendance", "No recognized faces within time limit.")

# ------------- Main Dashboard -------------
class FaceRecognitionSystem(Tk, FaceFunctions):
    def __init__(self):
        Tk.__init__(self)
        FaceFunctions.__init__(self)
        self.title("Face Recognition Attendance System")
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{sw}x{sh}+0+0")
        self.resizable(False, False)

        # Background
        bg_frame = Frame(self, bg="white")
        bg_frame.place(x=0, y=60, width=sw, height=sh-60)

        # Header
        Label(self, text="FACE RECOGNITION ATTENDANCE SYSTEM",
              font=("tahoma", 18, "bold"), bg="white", fg="red").place(x=0, y=0, width=sw, height=60)

        # Controls panel
        panel = Frame(bg_frame, bg="white")
        panel.pack(pady=20)

        self.reg_name = StringVar()
        Label(panel, text="Name:", font=("Arial", 12), bg="white").grid(row=0, column=0, padx=8, pady=8, sticky=E)
        Entry(panel, textvariable=self.reg_name, font=("Arial", 12), width=26).grid(row=0, column=1, padx=8, pady=8)

        Button(panel, text="Register Face", font=("Arial", 12, "bold"), bg="blue", fg="white",
               command=lambda: self.register_face(self.reg_name.get())).grid(row=0, column=2, padx=8, pady=8)
        Button(panel, text="Train Model", font=("Arial", 12, "bold"), bg="orange", fg="black",
               command=self.train_model).grid(row=0, column=3, padx=8, pady=8)
        Button(panel, text="Face Attendance", font=("Arial", 12, "bold"), bg="green", fg="white",
               command=self.face_login_and_mark).grid(row=0, column=4, padx=8, pady=8)
        Button(panel, text="Open Student Form", font=("Arial", 12, "bold"),
               command=self.open_student_form).grid(row=0, column=5, padx=8, pady=8)
        Button(panel, text="Open Dataset Folder", font=("Arial", 12),
               command=lambda: os.startfile(os.path.abspath(DATASET_DIR)) if os.name == "nt" else None).grid(row=0, column=6, padx=8, pady=8)
        Button(panel, text="Exit", font=("Arial", 12, "bold"), bg="red", fg="white",
               command=self.destroy).grid(row=0, column=7, padx=8, pady=8)

        # Info label
        Label(bg_frame, text="Workflow: 1) Register Face  2) Train Model  3) Face Attendance (auto; no key press) -> attendance.csv",
              font=("Arial", 12), bg="white", fg="gray").pack(pady=10)

        # Footer bar with ©
        footer = Frame(self, bg="white")
        footer.place(x=0, y=sh-28, width=sw, height=28)
        Label(footer, text=COPYRIGHT_TEXT, font=("Arial", 10), bg="white", fg="gray").pack(side=RIGHT, padx=12)

    def open_student_form(self):
        top = Toplevel(self)
        StudentApp(top)

# ------------- Run -------------
if __name__ == "__main__":
    app = FaceRecognitionSystem()
    app.mainloop()

import io
import os
import re
from datetime import datetime, timedelta
from functools import wraps

from flask import Flask, flash, jsonify, redirect, render_template, request, send_file, session, url_for
from flask_sqlalchemy import SQLAlchemy
from openpyxl import Workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from sqlalchemy import func, or_
from werkzeug.security import check_password_hash, generate_password_hash

app = Flask(__name__)
app.secret_key = "attendance-calculator-secret-key"
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///" + os.path.join(os.path.dirname(__file__), "database.db")
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)


class Student(db.Model):
    __tablename__ = "students"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    roll_number = db.Column(db.String(50), unique=True, nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password = db.Column(db.String(255), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    default_minimum_attendance = db.Column(db.Integer, default=75)
    dark_mode = db.Column(db.Boolean, default=False)

    subjects = db.relationship("Subject", backref="student", cascade="all, delete-orphan")
    attendance_entries = db.relationship("AttendanceEntry", backref="student", cascade="all, delete-orphan")


class Subject(db.Model):
    __tablename__ = "subjects"
    id = db.Column(db.Integer, primary_key=True)
    student_id = db.Column(db.Integer, db.ForeignKey("students.id"), nullable=False)
    subject_name = db.Column(db.String(100), nullable=False)
    faculty_name = db.Column(db.String(100), nullable=False)
    minimum_attendance = db.Column(db.Integer, default=75)
    total_classes = db.Column(db.Integer, default=0)
    present_classes = db.Column(db.Integer, default=0)

    attendance_entries = db.relationship("AttendanceEntry", backref="subject", cascade="all, delete-orphan")


class AttendanceEntry(db.Model):
    __tablename__ = "attendance"
    id = db.Column(db.Integer, primary_key=True)
    student_id = db.Column(db.Integer, db.ForeignKey("students.id"), nullable=False)
    subject_id = db.Column(db.Integer, db.ForeignKey("subjects.id"), nullable=False)
    date = db.Column(db.String(20), nullable=False)
    time = db.Column(db.String(20), nullable=False)
    status = db.Column(db.String(20), nullable=False)


with app.app_context():
    db.create_all()


def login_required(view):
    @wraps(view)
    def wrapped_view(*args, **kwargs):
        if "student_id" not in session:
            flash("Please log in to continue.", "warning")
            return redirect(url_for("login"))
        return view(*args, **kwargs)

    return wrapped_view


def calculate_percentage(total, present):
    if total == 0:
        return 0
    return round((present / total) * 100, 1)


def get_status_label(percent):
    if percent >= 75:
        return "Safe", "safe"
    if percent >= 65:
        return "Warning", "warning"
    return "Critical", "danger"


def get_current_student():
    if "student_id" not in session:
        return None
    return Student.query.get(session["student_id"])


def validate_email(email):
    return re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email) is not None


def validate_password(password):
    return len(password) >= 8 and any(char.isdigit() for char in password) and any(char.isalpha() for char in password)


def get_dashboard_data(student):
    subjects = Subject.query.filter_by(student_id=student.id).all()
    overall_total = sum(subject.total_classes for subject in subjects)
    overall_present = sum(subject.present_classes for subject in subjects)
    overall_absent = overall_total - overall_present
    overall_percent = calculate_percentage(overall_total, overall_present)
    safe_subjects = sum(1 for subject in subjects if calculate_percentage(subject.total_classes, subject.present_classes) >= subject.minimum_attendance)
    low_subjects = sum(1 for subject in subjects if calculate_percentage(subject.total_classes, subject.present_classes) < subject.minimum_attendance)

    bar_labels = [subject.subject_name for subject in subjects]
    bar_values = [calculate_percentage(subject.total_classes, subject.present_classes) for subject in subjects]

    recent_entries = (
        AttendanceEntry.query.filter_by(student_id=student.id)
        .order_by(AttendanceEntry.id.desc())
        .limit(8)
        .all()
    )

    trend_labels = []
    trend_present = []
    trend_absent = []
    for i in range(6):
        day = (datetime.utcnow() - timedelta(days=5 - i)).date().isoformat()
        day_entries = [entry for entry in recent_entries if entry.date == day]
        present_count = sum(1 for entry in day_entries if entry.status == "present")
        absent_count = sum(1 for entry in day_entries if entry.status == "absent")
        trend_labels.append(day)
        trend_present.append(present_count)
        trend_absent.append(absent_count)

    warning_subjects = [subject for subject in subjects if calculate_percentage(subject.total_classes, subject.present_classes) < subject.minimum_attendance]
    warning_message = None
    if warning_subjects:
        warning_message = f"{len(warning_subjects)} subject(s) need attention."
    return {
        "subjects": subjects,
        "overall_total": overall_total,
        "overall_present": overall_present,
        "overall_absent": overall_absent,
        "overall_percent": overall_percent,
        "safe_subjects": safe_subjects,
        "low_subjects": low_subjects,
        "bar_labels": bar_labels,
        "bar_values": bar_values,
        "recent_entries": recent_entries,
        "trend_labels": trend_labels,
        "trend_present": trend_present,
        "trend_absent": trend_absent,
        "warning_subjects": warning_subjects,
        "warning_message": warning_message,
    }


@app.context_processor
def inject_student():
    return {
        "current_student": get_current_student(),
        "datetime": datetime,
        "calculate_percentage": calculate_percentage,
        "get_status_label": get_status_label,
    }


@app.route("/")
def index():
    if "student_id" in session:
        return redirect(url_for("dashboard"))
    return redirect(url_for("login"))


@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        roll_number = request.form.get("roll_number", "").strip()
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")
        confirm_password = request.form.get("confirm_password", "")

        if not all([name, roll_number, email, password, confirm_password]):
            flash("All fields are required.", "danger")
            return render_template("register.html")

        if Student.query.filter_by(email=email).first():
            flash("An account with this email already exists.", "danger")
            return render_template("register.html")

        if Student.query.filter_by(roll_number=roll_number).first():
            flash("This roll number is already registered.", "danger")
            return render_template("register.html")

        if not validate_email(email):
            flash("Please enter a valid email address.", "danger")
            return render_template("register.html")

        if password != confirm_password:
            flash("Passwords do not match.", "danger")
            return render_template("register.html")

        if not validate_password(password):
            flash("Password must be at least 8 characters and include letters and numbers.", "danger")
            return render_template("register.html")

        student = Student(
            name=name,
            roll_number=roll_number,
            email=email,
            password=generate_password_hash(password),
            default_minimum_attendance=75,
        )
        db.session.add(student)
        db.session.commit()
        flash("Registration successful! Please log in.", "success")
        return redirect(url_for("login"))

    return render_template("register.html")


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")
        student = Student.query.filter_by(email=email).first()
        if student and check_password_hash(student.password, password):
            session["student_id"] = student.id
            flash("Welcome back!", "success")
            return redirect(url_for("dashboard"))
        flash("Invalid email or password.", "danger")

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.pop("student_id", None)
    flash("You have been logged out.", "info")
    return redirect(url_for("login"))


@app.route("/dashboard")
@login_required
def dashboard():
    student = get_current_student()
    data = get_dashboard_data(student)
    return render_template("dashboard.html", student=student, data=data)


@app.route("/subjects/add", methods=["GET", "POST"])
@login_required
def add_subject():
    student = get_current_student()
    if request.method == "POST":
        subject_name = request.form.get("subject_name", "").strip()
        faculty_name = request.form.get("faculty_name", "").strip()
        minimum_attendance = request.form.get("minimum_attendance", "")
        total_classes = request.form.get("total_classes", "")
        present_classes = request.form.get("present_classes", "")

        if not all([subject_name, faculty_name, minimum_attendance, total_classes, present_classes]):
            flash("Please fill in all fields.", "danger")
            return render_template("add_subject.html", student=student)

        if Subject.query.filter_by(student_id=student.id).filter(func.lower(Subject.subject_name) == subject_name.lower()).first():
            flash("A subject with this name already exists.", "danger")
            return render_template("add_subject.html", student=student)

        try:
            minimum_attendance = int(minimum_attendance)
            total_classes = int(total_classes)
            present_classes = int(present_classes)
        except ValueError:
            flash("Attendance numbers must be valid integers.", "danger")
            return render_template("add_subject.html", student=student)

        if minimum_attendance < 0 or minimum_attendance > 100:
            flash("Minimum attendance must be between 0 and 100.", "danger")
            return render_template("add_subject.html", student=student)

        if total_classes < 0 or present_classes < 0:
            flash("Class counts cannot be negative.", "danger")
            return render_template("add_subject.html", student=student)

        if present_classes > total_classes:
            flash("Present classes cannot exceed total classes.", "danger")
            return render_template("add_subject.html", student=student)

        subject = Subject(
            student_id=student.id,
            subject_name=subject_name,
            faculty_name=faculty_name,
            minimum_attendance=minimum_attendance,
            total_classes=total_classes,
            present_classes=present_classes,
        )
        db.session.add(subject)
        db.session.commit()
        flash("Subject added successfully.", "success")
        return redirect(url_for("dashboard"))

    return render_template("add_subject.html", student=student)


@app.route("/subjects/<int:subject_id>/edit", methods=["GET", "POST"])
@login_required
def edit_subject(subject_id):
    student = get_current_student()
    subject = Subject.query.filter_by(id=subject_id, student_id=student.id).first_or_404()
    if request.method == "POST":
        subject_name = request.form.get("subject_name", "").strip()
        faculty_name = request.form.get("faculty_name", "").strip()
        minimum_attendance = request.form.get("minimum_attendance", "")
        total_classes = request.form.get("total_classes", "")
        present_classes = request.form.get("present_classes", "")

        if not all([subject_name, faculty_name, minimum_attendance, total_classes, present_classes]):
            flash("Please fill in all fields.", "danger")
            return render_template("edit_subject.html", student=student, subject=subject)

        existing = Subject.query.filter_by(student_id=student.id).filter(func.lower(Subject.subject_name) == subject_name.lower()).first()
        if existing and existing.id != subject.id:
            flash("A subject with this name already exists.", "danger")
            return render_template("edit_subject.html", student=student, subject=subject)

        try:
            minimum_attendance = int(minimum_attendance)
            total_classes = int(total_classes)
            present_classes = int(present_classes)
        except ValueError:
            flash("Attendance numbers must be valid integers.", "danger")
            return render_template("edit_subject.html", student=student, subject=subject)

        if minimum_attendance < 0 or minimum_attendance > 100:
            flash("Minimum attendance must be between 0 and 100.", "danger")
            return render_template("edit_subject.html", student=student, subject=subject)

        if total_classes < 0 or present_classes < 0:
            flash("Class counts cannot be negative.", "danger")
            return render_template("edit_subject.html", student=student, subject=subject)

        if present_classes > total_classes:
            flash("Present classes cannot exceed total classes.", "danger")
            return render_template("edit_subject.html", student=student, subject=subject)

        subject.subject_name = subject_name
        subject.faculty_name = faculty_name
        subject.minimum_attendance = minimum_attendance
        subject.total_classes = total_classes
        subject.present_classes = present_classes
        db.session.commit()
        flash("Subject updated successfully.", "success")
        return redirect(url_for("dashboard"))

    return render_template("edit_subject.html", student=student, subject=subject)


@app.route("/subjects/<int:subject_id>/delete", methods=["POST"])
@login_required
def delete_subject(subject_id):
    student = get_current_student()
    subject = Subject.query.filter_by(id=subject_id, student_id=student.id).first_or_404()
    db.session.delete(subject)
    db.session.commit()
    flash("Subject deleted successfully.", "success")
    return redirect(url_for("dashboard"))


@app.route("/attendance")
@login_required
def attendance():
    student = get_current_student()
    subjects = Subject.query.filter_by(student_id=student.id).all()
    recent_entries = (
        AttendanceEntry.query.filter_by(student_id=student.id)
        .order_by(AttendanceEntry.id.desc())
        .limit(10)
        .all()
    )
    return render_template("attendance.html", student=student, subjects=subjects, recent_entries=recent_entries)


@app.route("/api/subjects/<int:subject_id>/attendance", methods=["POST"])
@login_required
def record_attendance(subject_id):
    student = get_current_student()
    subject = Subject.query.filter_by(id=subject_id, student_id=student.id).first_or_404()
    status = request.form.get("status", "present")
    if status not in {"present", "absent"}:
        return jsonify({"success": False, "message": "Invalid attendance status."}), 400

    subject.total_classes += 1
    if status == "present":
        subject.present_classes += 1

    entry = AttendanceEntry(
        student_id=student.id,
        subject_id=subject.id,
        date=datetime.utcnow().date().isoformat(),
        time=datetime.utcnow().strftime("%H:%M"),
        status=status,
    )
    db.session.add(entry)
    db.session.commit()

    percent = calculate_percentage(subject.total_classes, subject.present_classes)
    label, badge_class = get_status_label(percent)
    return jsonify(
        {
            "success": True,
            "message": f"Attendance updated for {subject.subject_name} as {status}.",
            "percent": percent,
            "label": label,
            "badge_class": badge_class,
        }
    )


@app.route("/reports")
@login_required
def reports():
    student = get_current_student()
    query = request.args.get("q", "").strip()
    status_filter = request.args.get("status", "all")
    sort_by = request.args.get("sort", "name")

    subjects = Subject.query.filter_by(student_id=student.id)
    if query:
        subjects = subjects.filter(
            or_(
                Subject.subject_name.ilike(f"%{query}%"),
                Subject.faculty_name.ilike(f"%{query}%"),
            )
        )

    if status_filter != "all":
        subjects = subjects.all()
        subjects = [subject for subject in subjects if status_filter == get_status_label(calculate_percentage(subject.total_classes, subject.present_classes))[0].lower()]
    else:
        subjects = subjects.all()

    if sort_by == "attendance":
        subjects.sort(key=lambda subject: calculate_percentage(subject.total_classes, subject.present_classes), reverse=True)
    elif sort_by == "name":
        subjects.sort(key=lambda subject: subject.subject_name.lower())
    elif sort_by == "faculty":
        subjects.sort(key=lambda subject: subject.faculty_name.lower())
    else:
        subjects.sort(key=lambda subject: subject.subject_name.lower())

    return render_template("reports.html", student=student, subjects=subjects, query=query, status_filter=status_filter, sort_by=sort_by)


@app.route("/reports/export/pdf")
@login_required
def export_pdf():
    student = get_current_student()
    subjects = Subject.query.filter_by(student_id=student.id).all()
    buffer = io.BytesIO()
    pdf = SimpleDocTemplate(buffer, pagesize=letter)
    elements = []
    elements.append(Paragraph(f"Attendance Report - {student.name}", style="Heading1"))
    elements.append(Spacer(1, 12))
    table_data = [["Subject", "Faculty", "Present", "Absent", "Total", "Attendance %", "Status"]]
    for subject in subjects:
        percent = calculate_percentage(subject.total_classes, subject.present_classes)
        label, _ = get_status_label(percent)
        table_data.append([subject.subject_name, subject.faculty_name, subject.present_classes, subject.total_classes - subject.present_classes, subject.total_classes, f"{percent}%", label])
    table = Table(table_data, repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2563eb")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ]
        )
    )
    elements.append(table)
    pdf.build(elements)
    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name="attendance_report.pdf", mimetype="application/pdf")


@app.route("/reports/export/excel")
@login_required
def export_excel():
    student = get_current_student()
    subjects = Subject.query.filter_by(student_id=student.id).all()
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Attendance Report"
    headers = ["Subject", "Faculty", "Present", "Absent", "Total", "Attendance %", "Status"]
    sheet.append(headers)
    for subject in subjects:
        percent = calculate_percentage(subject.total_classes, subject.present_classes)
        label, _ = get_status_label(percent)
        sheet.append([subject.subject_name, subject.faculty_name, subject.present_classes, subject.total_classes - subject.present_classes, subject.total_classes, f"{percent}%", label])
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name="attendance_report.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/settings", methods=["GET", "POST"])
@login_required
def settings():
    student = get_current_student()
    if request.method == "POST":
        if "profile_submit" in request.form:
            name = request.form.get("name", "").strip()
            roll_number = request.form.get("roll_number", "").strip()
            email = request.form.get("email", "").strip().lower()
            if not all([name, roll_number, email]):
                flash("Profile fields cannot be blank.", "danger")
                return render_template("settings.html", student=student)
            existing_email = Student.query.filter(Student.email == email, Student.id != student.id).first()
            if existing_email:
                flash("This email is already in use.", "danger")
                return render_template("settings.html", student=student)
            existing_roll = Student.query.filter(Student.roll_number == roll_number, Student.id != student.id).first()
            if existing_roll:
                flash("This roll number is already registered.", "danger")
                return render_template("settings.html", student=student)
            student.name = name
            student.roll_number = roll_number
            student.email = email
            db.session.commit()
            flash("Profile updated successfully.", "success")
        elif "password_submit" in request.form:
            current_password = request.form.get("current_password", "")
            new_password = request.form.get("new_password", "")
            confirm_password = request.form.get("confirm_password", "")
            if not check_password_hash(student.password, current_password):
                flash("Current password is incorrect.", "danger")
                return render_template("settings.html", student=student)
            if new_password != confirm_password:
                flash("New passwords do not match.", "danger")
                return render_template("settings.html", student=student)
            if not validate_password(new_password):
                flash("New password must be at least 8 characters and include letters and numbers.", "danger")
                return render_template("settings.html", student=student)
            student.password = generate_password_hash(new_password)
            db.session.commit()
            flash("Password changed successfully.", "success")
        elif "settings_submit" in request.form:
            try:
                student.default_minimum_attendance = int(request.form.get("default_minimum_attendance", student.default_minimum_attendance))
            except ValueError:
                flash("Minimum attendance must be an integer.", "danger")
                return render_template("settings.html", student=student)
            student.dark_mode = bool(request.form.get("dark_mode"))
            db.session.commit()
            flash("Settings updated successfully.", "success")
        return redirect(url_for("settings"))

    return render_template("settings.html", student=student)


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)

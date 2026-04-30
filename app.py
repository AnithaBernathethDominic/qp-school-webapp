import os
import sqlite3
import uuid
from datetime import datetime

from flask import Flask, flash, redirect, render_template, request, send_from_directory, url_for
from flask_login import LoginManager, UserMixin, current_user, login_required, login_user, logout_user
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

from analyzer import analyze_files, make_docx_report, make_excel_report, summary_tables

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
REPORT_DIR = os.path.join(BASE_DIR, "reports")
DB_PATH = os.path.join(BASE_DIR, "instance", "app.db")
ALLOWED_EXTENSIONS = {"pdf"}

app = Flask(__name__)
app.secret_key = "change-this-secret-key"
app.config["UPLOAD_DIR"] = UPLOAD_DIR
app.config["REPORT_DIR"] = REPORT_DIR

login_manager = LoginManager(app)
login_manager.login_view = "login"

class User(UserMixin):
    def __init__(self, id, username):
        self.id = str(id)
        self.username = username


def db():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    return con


def init_db():
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    os.makedirs(REPORT_DIR, exist_ok=True)
    with db() as con:
        con.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL
            )
        """)
        con.execute("""
            CREATE TABLE IF NOT EXISTS analyses (
                id TEXT PRIMARY KEY,
                created_at TEXT NOT NULL,
                teacher TEXT NOT NULL,
                qp_filename TEXT NOT NULL,
                syllabus_filename TEXT NOT NULL,
                total_questions INTEGER NOT NULL,
                docx_file TEXT NOT NULL,
                xlsx_file TEXT NOT NULL
            )
        """)
        user = con.execute("SELECT id FROM users WHERE username=?", ("admin",)).fetchone()
        if not user:
            con.execute("INSERT INTO users(username, password_hash) VALUES (?, ?)", ("admin", generate_password_hash("admin123")))
        con.commit()


@login_manager.user_loader
def load_user(user_id):
    with db() as con:
        row = con.execute("SELECT id, username FROM users WHERE id=?", (user_id,)).fetchone()
    return User(row["id"], row["username"]) if row else None


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/", methods=["GET"])
def home():
    if current_user.is_authenticated:
        return redirect(url_for("dashboard"))
    return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        with db() as con:
            row = con.execute("SELECT * FROM users WHERE username=?", (username,)).fetchone()
        if row and check_password_hash(row["password_hash"], password):
            login_user(User(row["id"], row["username"]))
            return redirect(url_for("dashboard"))
        flash("Invalid username or password", "danger")
    return render_template("login.html")


@app.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("login"))


@app.route("/dashboard")
@login_required
def dashboard():
    with db() as con:
        rows = con.execute("SELECT * FROM analyses ORDER BY created_at DESC LIMIT 10").fetchall()
    return render_template("dashboard.html", analyses=rows)


@app.route("/analyze", methods=["POST"])
@login_required
def analyze():
    qp = request.files.get("question_paper")
    syllabus = request.files.get("syllabus")
    if not qp or not syllabus:
        flash("Please upload both the question paper PDF and syllabus PDF.", "danger")
        return redirect(url_for("dashboard"))
    if not allowed_file(qp.filename) or not allowed_file(syllabus.filename):
        flash("Only PDF files are allowed.", "danger")
        return redirect(url_for("dashboard"))

    analysis_id = uuid.uuid4().hex[:12]
    qp_name = secure_filename(qp.filename)
    sy_name = secure_filename(syllabus.filename)
    qp_path = os.path.join(UPLOAD_DIR, f"{analysis_id}_qp_{qp_name}")
    sy_path = os.path.join(UPLOAD_DIR, f"{analysis_id}_syllabus_{sy_name}")
    qp.save(qp_path)
    syllabus.save(sy_path)

    try:
        df, _ = analyze_files(qp_path, sy_path)
        if df.empty:
            flash("No questions were detected. Please check that the question paper PDF has selectable text.", "danger")
            return redirect(url_for("dashboard"))
        docx_file = f"{analysis_id}_topic_weightage_report.docx"
        xlsx_file = f"{analysis_id}_topic_weightage_report.xlsx"
        make_docx_report(df, os.path.join(REPORT_DIR, docx_file))
        make_excel_report(df, os.path.join(REPORT_DIR, xlsx_file))
        with db() as con:
            con.execute(
                "INSERT INTO analyses VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                (analysis_id, datetime.now().isoformat(timespec="seconds"), current_user.username, qp_name, sy_name, int(len(df)), docx_file, xlsx_file),
            )
            con.commit()
        return redirect(url_for("result", analysis_id=analysis_id))
    except Exception as e:
        flash(f"Analysis failed: {e}", "danger")
        return redirect(url_for("dashboard"))


@app.route("/result/<analysis_id>")
@login_required
def result(analysis_id):
    with db() as con:
        row = con.execute("SELECT * FROM analyses WHERE id=?", (analysis_id,)).fetchone()
    if not row:
        flash("Analysis not found.", "danger")
        return redirect(url_for("dashboard"))

    # Re-read Excel for display.
    xlsx_path = os.path.join(REPORT_DIR, row["xlsx_file"])
    import pandas as pd
    df = pd.read_excel(xlsx_path, sheet_name="Question Mapping")
    chapter = pd.read_excel(xlsx_path, sheet_name="Chapter Weightage")
    subtopic = pd.read_excel(xlsx_path, sheet_name="Subtopic Weightage")
    return render_template(
        "result.html",
        analysis=row,
        mapping=df.to_dict("records"),
        chapters=chapter.to_dict("records"),
        subtopics=subtopic.to_dict("records"),
        chart_labels=list(chapter["chapter"].astype(str)),
        chart_values=list(chapter["marks"].astype(int)),
    )


@app.route("/download/<filename>")
@login_required
def download(filename):
    return send_from_directory(REPORT_DIR, filename, as_attachment=True)


if __name__ == "__main__":
    init_db()
    app.run(debug=True)

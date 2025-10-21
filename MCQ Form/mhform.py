import sqlite3
import json
import tempfile
import pythoncom
from flask import send_file
from docxtpl import DocxTemplate
from docx2pdf import convert
import tempfile, os, pythoncom
from datetime import datetime
from flask import render_template, send_file
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.colors import HexColor
from reportlab.lib.utils import ImageReader
import os
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx2pdf import convert
from io import BytesIO
import io, os
from datetime import datetime, timedelta
import pandas as pd
from flask import (
    Flask, g, render_template, request, redirect, url_for,
    session, send_file, flash
)
import random
import string

# -------------------------
# Config
# -------------------------
APP_SECRET = 'change_this_secret_for_prod'
DB_PATH = os.path.join(os.path.dirname(__file__), 'data.db')
ADMIN_USERNAME = 'RamG'
ADMIN_PASSWORD = 'Ram.v@123'
USER_PASSWORD_VALIDITY_HOURS = 24  # 1 day
MAX_ATTEMPTS = 2

app = Flask(__name__)
app.secret_key = APP_SECRET

# -------------------------
# Database helpers
# -------------------------
def get_db():
    db = getattr(g, '_database', None)
    if db is None:
        db = g._database = sqlite3.connect(DB_PATH)
        db.row_factory = sqlite3.Row
    return db

@app.teardown_appcontext
def close_connection(exception):
    db = getattr(g, '_database', None)
    if db is not None:
        db.close()

def init_db():
    db = get_db()
    cur = db.cursor()
    # Sections and questions
    cur.execute('''
        CREATE TABLE IF NOT EXISTS sections (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            description TEXT
        )
    ''')
    cur.execute('''
        CREATE TABLE IF NOT EXISTS questions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            section_id INTEGER NOT NULL,
            text TEXT NOT NULL,
            options TEXT NOT NULL,
            correct_option INTEGER NOT NULL,
            FOREIGN KEY(section_id) REFERENCES sections(id)
        )
    ''')
    # Submissions and answers
    cur.execute('''
        CREATE TABLE IF NOT EXISTS submissions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            section_id INTEGER NOT NULL,
            responder_id INTEGER NOT NULL,
            attempt_no INTEGER NOT NULL,
            timestamp TEXT NOT NULL,
            FOREIGN KEY(section_id) REFERENCES sections(id),
            FOREIGN KEY(responder_id) REFERENCES responders(id)
        )
    ''')
    cur.execute('''
        CREATE TABLE IF NOT EXISTS answers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            submission_id INTEGER NOT NULL,
            question_id INTEGER NOT NULL,
            selected_option INTEGER NOT NULL,
            is_correct INTEGER NOT NULL,
            FOREIGN KEY(submission_id) REFERENCES submissions(id),
            FOREIGN KEY(question_id) REFERENCES questions(id)
        )
    ''')
    # Responders table with username/password
    cur.execute('''
        CREATE TABLE IF NOT EXISTS responders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            first_name TEXT NOT NULL,
            middle_name TEXT,
            last_name TEXT NOT NULL,
            email TEXT NOT NULL UNIQUE,
            mobile TEXT,
            dob TEXT,
            designation TEXT,
            username TEXT UNIQUE,
            password TEXT,
            assigned_section_id INTEGER,
            created_at TEXT NOT NULL,
            FOREIGN KEY(assigned_section_id) REFERENCES sections(id)
        )
    ''')
    db.commit()

with app.app_context():
    init_db()

# -------------------------
# Decorators
# -------------------------
def admin_required(fn):
    from functools import wraps
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not session.get('admin'):
            flash('Admin login required')
            return redirect(url_for('login'))
        return fn(*args, **kwargs)
    return wrapper

def responder_required(fn):
    from functools import wraps
    @wraps(fn)
    def wrapper(*args, **kwargs):
        username = session.get('responder_username')
        if not username:
            flash('Please login first.')
            return redirect(url_for('login'))
        db = get_db()
        responder = db.execute('SELECT * FROM responders WHERE username=?', (username,)).fetchone()
        if not responder:
            flash('User not found.')
            return redirect(url_for('login'))
        # Check 24h validity
        created = datetime.strptime(responder['created_at'], "%Y-%m-%d %H:%M:%S")
        if datetime.now() > created + timedelta(hours=USER_PASSWORD_VALIDITY_HOURS):
            flash('Your credentials expired. Contact admin.')
            return redirect(url_for('login'))
        return fn(*args, **kwargs)
    return wrapper

# -------------------------
# Public routes
# -------------------------
@app.route('/')
def home():
    return redirect(url_for('login'))

@app.route('/login', methods=['GET','POST'])
def login():
    if request.method=='POST':
        username = request.form['username']
        password = request.form['password']
        db = get_db()
        # Admin login
        if username==ADMIN_USERNAME and password==ADMIN_PASSWORD:
            session['admin'] = True
            flash("Logged in as admin")
            return redirect(url_for('admin_dashboard'))
        # Responder login
        user = db.execute('SELECT * FROM responders WHERE username=? AND password=?', (username, password)).fetchone()
        if user:
            created = datetime.strptime(user['created_at'], "%Y-%m-%d %H:%M:%S")
            if datetime.now() > created + timedelta(hours=USER_PASSWORD_VALIDITY_HOURS):
                flash('Your credentials expired. Contact admin.')
                return redirect(url_for('login'))
            session['responder_username'] = username
            flash('Logged in successfully')
            return redirect(url_for('section_page', section_id=user['assigned_section_id']))
        flash('Invalid credentials')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.pop('responder_username', None)
    session.pop('admin', None)
    flash('Logged out successfully')
    return redirect(url_for('login'))

# -------------------------
# Admin routes
# -------------------------
@app.route('/admin/dashboard')
@admin_required
def admin_dashboard():
    db = get_db()
    sections = db.execute('SELECT * FROM sections').fetchall()
    responders_count = db.execute('SELECT COUNT(*) as cnt FROM responders').fetchone()['cnt']
    return render_template('admin_dashboard.html', sections=sections, responders_count=responders_count)

@app.route('/admin/create_user', methods=['GET','POST'])
@admin_required
def create_user():
    db = get_db()
    sections = db.execute('SELECT * FROM sections').fetchall()
    if request.method=='POST':
        first = request.form['first_name']
        middle = request.form.get('middle_name')
        last = request.form['last_name']
        email = request.form['email']
        mobile = request.form.get('mobile')
        dob = request.form.get('dob')
        designation = request.form.get('designation')
        assigned_section_id = int(request.form['assigned_section_id'])
        # Generate random username/password
        username = 'user' + ''.join(random.choices(string.digits, k=4))
        password = ''.join(random.choices(string.ascii_letters + string.digits, k=6))
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            db.execute('''
                INSERT INTO responders
                (first_name,middle_name,last_name,email,mobile,dob,designation,username,password,assigned_section_id,created_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,?)
            ''', (first,middle,last,email,mobile,dob,designation,username,password,assigned_section_id,timestamp))
            db.commit()
            flash(f'User created successfully. Username: {username}, Password: {password}')
        except sqlite3.IntegrityError:
            flash('Email or username already exists')
            return redirect(url_for('create_user'))
    return render_template('create_user.html', sections=sections)

@app.route('/admin/section/create', methods=['GET','POST'])
@admin_required
def create_section():
    db = get_db()
    if request.method=='POST':
        name = request.form['name']
        desc = request.form.get('description')
        db.execute('INSERT INTO sections (name, description) VALUES (?,?)', (name, desc))
        db.commit()
        flash("Section created successfully")
        return redirect(url_for('admin_dashboard'))
    return render_template('create_section.html')

@app.route('/admin/section/<int:section_id>/submissions')
@admin_required
def view_submissions(section_id):
    db = get_db()
    section = db.execute('SELECT * FROM sections WHERE id=?', (section_id,)).fetchone()
    if not section:
        flash("Section not found")
        return redirect(url_for('admin_dashboard'))
    submissions = db.execute('''
        SELECT 
            s.id, s.timestamp, r.first_name || ' ' || r.last_name AS name,
            r.email, SUM(a.is_correct) AS total_correct, COUNT(a.id) AS total_questions,
            (COUNT(a.id) - SUM(a.is_correct)) AS total_wrong
        FROM submissions s
        JOIN responders r ON s.responder_id = r.id
        JOIN answers a ON a.submission_id = s.id
        WHERE s.section_id = ?
        GROUP BY s.id
        ORDER BY s.timestamp DESC
    ''', (section_id,)).fetchall()
    return render_template('view_submissions.html', section=section, subs=submissions)

@app.route('/admin/section/<int:section_id>/download')
@admin_required
def download_file(section_id):
    db = get_db()
    section = db.execute('SELECT * FROM sections WHERE id=?', (section_id,)).fetchone()
    if not section:
        return "Section not found", 404

    tq_row = db.execute('SELECT COUNT(*) AS cnt FROM questions WHERE section_id=?', (section_id,)).fetchone()
    total_questions = tq_row['cnt'] if tq_row is not None else 0

    # ========================= Summary Sheet =========================
    summary_sql = '''
        SELECT 
            s.id AS submission_id, 
            r.first_name || ' ' || r.last_name AS responder_name,
            r.email AS responder_email, 
            s.timestamp AS submitted_on,
            COUNT(a.id) AS answered_count, 
            COALESCE(SUM(a.is_correct),0) AS correct_count
        FROM submissions s
        JOIN responders r ON s.responder_id = r.id
        LEFT JOIN answers a ON a.submission_id = s.id
        WHERE s.section_id=?
        GROUP BY s.id
        ORDER BY s.timestamp DESC
    '''
    df_summary = pd.read_sql_query(summary_sql, db, params=(section_id,))

    if df_summary.empty:
        df_summary = pd.DataFrame(columns=[
            'submission_id','responder_name','responder_email','submitted_on',
            'answered_count','correct_count','wrong_count','total_questions','status'
        ])
    else:
        # 🕓 Format timestamp properly
        df_summary['submitted_on'] = pd.to_datetime(df_summary['submitted_on'], errors='coerce')
        df_summary['submitted_on'] = df_summary['submitted_on'].dt.strftime('%Y-%m-%d %H:%M:%S')

        df_summary['correct_count'] = df_summary['correct_count'].fillna(0).astype(int)
        df_summary['answered_count'] = df_summary['answered_count'].fillna(0).astype(int)
        df_summary['wrong_count'] = df_summary['answered_count'] - df_summary['correct_count']
        df_summary['total_questions'] = total_questions
        df_summary['status'] = df_summary['correct_count'].apply(
            lambda c: 'Pass' if (total_questions > 0 and int(c) == total_questions) else 'Fail'
        )

    df_summary = df_summary[
        ['submission_id','responder_name','responder_email','submitted_on',
         'total_questions','answered_count','correct_count','wrong_count','status']
    ]

    # ========================= Details Sheet =========================
    details_sql = '''
        SELECT 
            s.id AS submission_id,
            r.first_name || ' ' || r.last_name AS responder_name,
            r.email AS responder_email,
            s.timestamp AS submitted_on,
            q.id AS question_id,
            q.text AS question_text,
            CASE 
                WHEN a.selected_option >= 0 THEN json_extract(q.options, '$[' || a.selected_option || ']')
                ELSE 'No answer'
            END AS selected_answer,
            CASE 
                WHEN q.correct_option >= 0 THEN json_extract(q.options, '$[' || q.correct_option || ']')
                ELSE 'N/A'
            END AS correct_answer,
            a.is_correct AS is_correct
        FROM submissions s
        JOIN responders r ON s.responder_id = r.id
        JOIN answers a ON a.submission_id = s.id
        JOIN questions q ON q.id = a.question_id
        WHERE s.section_id=?
        ORDER BY s.timestamp DESC, s.id, q.id
    '''
    df_details = pd.read_sql_query(details_sql, db, params=(section_id,))

    if df_details.empty:
        df_details = pd.DataFrame(columns=[
            'submission_id','responder_name','responder_email','submitted_on',
            'question_id','question_text','selected_answer','correct_answer','is_correct'
        ])
    else:
        # 🕓 Format timestamp properly
        df_details['submitted_on'] = pd.to_datetime(df_details['submitted_on'], errors='coerce')
        df_details['submitted_on'] = df_details['submitted_on'].dt.strftime('%Y-%m-%d %H:%M:%S')

    df_details['is_correct'] = df_details['is_correct'].apply(
        lambda v: 'Correct' if int(v) == 1 else 'Wrong'
    )

    # ========================= Excel Export =========================
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl', datetime_format='YYYY-MM-DD HH:MM:SS') as writer:
        df_summary.to_excel(writer, index=False, sheet_name='Summary')
        df_details.to_excel(writer, index=False, sheet_name='Details')

    output.seek(0)
    filename = f"submissions_section_{section_id}.xlsx"
    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

@app.route('/admin/section/<int:section_id>/questions', methods=['GET','POST'])
@admin_required
def add_question(section_id):
    db = get_db()
    section = db.execute('SELECT * FROM sections WHERE id=?', (section_id,)).fetchone()
    if not section:
        flash("Section not found")
        return redirect(url_for('admin_dashboard'))

    if request.method=='POST':
        text = request.form['text']
        options = [request.form[f'opt_{i}'] for i in range(4)]
        correct_option = int(request.form['correct']) - 1
        db.execute('INSERT INTO questions (section_id,text,options,correct_option) VALUES (?,?,?,?)',
                   (section_id,text,json.dumps(options),correct_option))
        db.commit()
        flash("Question added successfully")
        return redirect(url_for('add_question', section_id=section_id))

    questions = db.execute('SELECT * FROM questions WHERE section_id=?', (section_id,)).fetchall()
    questions_list = []
    for q in questions:
        questions_list.append({
            'id': q['id'],
            'text': q['text'],
            'options': json.loads(q['options']),
            'correct_option': q['correct_option']
        })
    return render_template('add_question.html', section=section, questions=questions_list)

@app.route('/admin/question/<int:question_id>/edit', methods=['GET','POST'])
@admin_required
def edit_question(question_id):
    db = get_db()
    q = db.execute('SELECT * FROM questions WHERE id=?', (question_id,)).fetchone()
    if not q:
        flash("Question not found")
        return redirect(url_for('admin_dashboard'))
    if request.method=='POST':
        text = request.form['text']
        options = [request.form[f'opt_{i}'] for i in range(4)]
        correct_option = int(request.form['correct']) - 1
        db.execute('UPDATE questions SET text=?, options=?, correct_option=? WHERE id=?',
                   (text,json.dumps(options),correct_option,question_id))
        db.commit()
        flash("Question updated successfully")
        return redirect(url_for('add_question', section_id=q['section_id']))
    opts = json.loads(q['options'])
    return render_template('edit_question.html', question=q, options=opts)

@app.route('/admin/question/<int:question_id>/delete')
@admin_required
def delete_question(question_id):
    db = get_db()
    q = db.execute('SELECT * FROM questions WHERE id=?', (question_id,)).fetchone()
    if q:
        db.execute('DELETE FROM questions WHERE id=?', (question_id,))
        db.commit()
        flash("Question deleted")
        return redirect(url_for('add_question', section_id=q['section_id']))
    flash("Question not found")
    return redirect(url_for('admin_dashboard'))

# -------------------------
# Responder Exam routes
# -------------------------
@app.route('/section/<int:section_id>')
@responder_required
def section_page(section_id):
    db = get_db()
    section = db.execute('SELECT * FROM sections WHERE id=?', (section_id,)).fetchone()
    if not section:
        return 'Section not found',404
    user = db.execute('SELECT * FROM responders WHERE username=?', (session['responder_username'],)).fetchone()
    attempts = db.execute('SELECT COUNT(*) as cnt FROM submissions WHERE responder_id=? AND section_id=?',
                          (user['id'], section_id)).fetchone()['cnt']
    if attempts >= MAX_ATTEMPTS:
        return "You have reached maximum 2 attempts for this exam.", 403
    qrows = db.execute('SELECT * FROM questions WHERE section_id=?', (section_id,)).fetchall()
    questions = []
    for q in qrows:
        questions.append({
            'id': q['id'],
            'text': q['text'],
            'options': json.loads(q['options'])
        })
    return render_template('section.html', section=section, questions=questions)

@app.route('/submit/<int:section_id>', methods=['POST'])
@responder_required
def submit_section(section_id):
    db = get_db()
    responder = db.execute('SELECT * FROM responders WHERE username=?', (session['responder_username'],)).fetchone()
    if not responder:
        flash('Responder not found')
        return redirect(url_for('login'))
    attempts = db.execute('SELECT COUNT(*) as cnt FROM submissions WHERE responder_id=? AND section_id=?',
                          (responder['id'], section_id)).fetchone()['cnt']
    attempt_no = attempts + 1
    timestamp = datetime.now().strftime("%d-%m-%Y, %H:%M:%S")
    cur = db.cursor()
    cur.execute('INSERT INTO submissions (section_id,responder_id,attempt_no,timestamp) VALUES (?,?,?,?)',
                (section_id,responder['id'],attempt_no,timestamp))
    submission_id = cur.lastrowid

    questions = db.execute('SELECT * FROM questions WHERE section_id=?', (section_id,)).fetchall()
    total = 0
    correct_count = 0
    for q in questions:
        qid = q['id']
        selected = request.form.get(f'q_{qid}')
        try:
            selected_idx = int(selected) if selected is not None else -1
        except ValueError:
            selected_idx = -1
        is_correct = 1 if selected_idx == q['correct_option'] else 0
        cur.execute('INSERT INTO answers (submission_id,question_id,selected_option,is_correct) VALUES (?,?,?,?)',
                    (submission_id,qid,selected_idx,is_correct))
        total +=1
        correct_count += is_correct
    db.commit()

    rows = db.execute('''
        SELECT q.text, q.options, q.correct_option, a.selected_option, a.is_correct
        FROM questions q
        JOIN answers a ON q.id=a.question_id
        WHERE a.submission_id=?
    ''', (submission_id,)).fetchall()

    details = []
    for r in rows:
        opts = json.loads(r['options'])
        sel_idx = r['selected_option']
        correct_idx = r['correct_option']
        details.append({
            'question_text': r['text'],
            'selected_text': opts[sel_idx] if 0 <= sel_idx < len(opts) else 'No answer',
            'correct_text': opts[correct_idx] if 0 <= correct_idx < len(opts) else 'N/A',
            'is_correct': bool(r['is_correct'])
        })

    submission = db.execute('SELECT * FROM submissions WHERE id=?', (submission_id,)).fetchone()
    return render_template('result.html', submission=submission, total=total, correct=correct_count, wrong=total-correct_count, details=details)

# -------------------------
# Certificate route
# -------------------------

@app.route('/certificate_pdf/<int:submission_id>')
def download_certificate_pdf(submission_id):
    db = get_db()
    
    # Fetch submission and responder details
    submission = db.execute('''
        SELECT s.id, s.section_id, r.first_name, r.middle_name, r.last_name, r.designation
        FROM submissions s
        JOIN responders r ON s.responder_id = r.id
        WHERE s.id=?
    ''', (submission_id,)).fetchone()
    
    if not submission:
        return "Submission not found", 404

    # Check if passed
    answers = db.execute('SELECT * FROM answers WHERE submission_id=?', (submission_id,)).fetchall()
    total = len(answers)
    correct = sum(a['is_correct'] for a in answers)
    if correct != total:
        return "Certificate available only for passed submissions.", 403

    # Fetch section name
    section = db.execute('SELECT name FROM sections WHERE id=?', (submission['section_id'],)).fetchone()
    full_name = f"{submission['first_name']} {submission['middle_name'] or ''} {submission['last_name']}".strip()

    # Word template path
    template_path = r"C:\Users\ramla\OneDrive\Desktop\MCQ Form\static\certificate_template.docx"
    if not os.path.exists(template_path):
        return "Certificate template not found", 500

    # Load template with docxtpl
    doc = DocxTemplate(template_path)
    context = {
        'full_name': full_name,
        'designation': submission['designation'] or '',
        'section_name': section['name'],
        'issued_date': datetime.now().strftime("%d-%m-%Y")
    }
    doc.render(context)

    # Save temporary Word file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_word:
        doc.save(tmp_word.name)
        tmp_word_path = tmp_word.name

    # Convert to PDF
    pythoncom.CoInitialize()
    tmp_pdf_path = tmp_word_path.replace(".docx", ".pdf")
    convert(tmp_word_path, tmp_pdf_path)
    pythoncom.CoUninitialize()

    # Send PDF
    return send_file(
        tmp_pdf_path,
        as_attachment=True,
        download_name=f"Certificate_{full_name.replace(' ', '_')}.pdf"
    )

# -------------------------
# Run server
# -------------------------
if __name__=='__main__':
    app.run(debug=True)

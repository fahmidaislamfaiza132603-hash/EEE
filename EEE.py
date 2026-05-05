"""
EduTrack Pro 2026 - COMPLETE WORKING VERSION
=============================================
Fixed Parser for EEE Department Excel Format
Handles: Midterm Exam, Final Exam, Assignment, Analysis of CO, Analysis of PO
With all original features + production upgrades
"""

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import plotly.graph_objects as go
import plotly.express as px
import seaborn as sns
import hashlib
import json
import os
import re
import smtplib
import time
import secrets
import pickle
import warnings
import sqlite3
import threading
import queue
import logging
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path
import pytz
import bcrypt
from openpyxl import load_workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph,
    Spacer, Image, PageBreak
)
from sklearn.linear_model import LinearRegression
from sklearn.ensemble import RandomForestClassifier
from sklearn.preprocessing import StandardScaler, LabelEncoder
from sklearn.model_selection import train_test_split, cross_val_score
from sklearn.metrics import accuracy_score
import joblib

warnings.filterwarnings('ignore')

# ============================================================================
# CONFIGURATION
# ============================================================================
class Config:
    APP_NAME = "EduTrack Pro 2026"
    VERSION = "2.0.0"
    DATA_DIR = Path("data")
    MODEL_DIR = Path("models")
    
    SMTP_SERVER = "smtp.gmail.com"
    SMTP_PORT = 587
    SMTP_EMAIL = "your_email@gmail.com"
    SMTP_PASSWORD = "your_app_password"
    
    CAREER_DOMAINS = [
        "Power Systems & Energy",
        "Electronics & Embedded Systems",
        "Telecommunications",
        "Control & Automation",
        "Research & Academia",
        "Renewable Energy",
        "AI & Machine Learning in EEE"
    ]

for d in [Config.DATA_DIR, Config.MODEL_DIR]:
    d.mkdir(parents=True, exist_ok=True)

# ============================================================================
# DATABASE SETUP
# ============================================================================
def init_db():
    conn = sqlite3.connect(str(Config.DATA_DIR / "edutrack.db"))
    conn.row_factory = sqlite3.Row
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            email TEXT,
            full_name TEXT NOT NULL,
            user_type TEXT NOT NULL,
            is_active INTEGER DEFAULT 1,
            student_id TEXT,
            parent_email TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        
        CREATE TABLE IF NOT EXISTS courses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            course_code TEXT NOT NULL,
            course_name TEXT NOT NULL,
            semester TEXT NOT NULL,
            teacher_username TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        
        CREATE TABLE IF NOT EXISTS activity_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL,
            action TEXT NOT NULL,
            details TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    """)
    conn.commit()
    conn.close()

# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================
def hash_password(password):
    return bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()

def verify_password(password, hashed):
    try:
        return bcrypt.checkpw(password.encode(), hashed.encode())
    except:
        return False

def calculate_sgpa(total_marks):
    if total_marks >= 80: return 4.00
    elif total_marks >= 75: return 3.75
    elif total_marks >= 70: return 3.50
    elif total_marks >= 65: return 3.25
    elif total_marks >= 60: return 3.00
    elif total_marks >= 55: return 2.75
    elif total_marks >= 50: return 2.50
    elif total_marks >= 45: return 2.25
    elif total_marks >= 40: return 2.00
    else: return 0.00

def get_grade_from_marks(total_marks):
    if total_marks >= 80: return "A+"
    elif total_marks >= 75: return "A"
    elif total_marks >= 70: return "A-"
    elif total_marks >= 65: return "B+"
    elif total_marks >= 60: return "B"
    elif total_marks >= 55: return "B-"
    elif total_marks >= 50: return "C+"
    elif total_marks >= 45: return "C"
    elif total_marks >= 40: return "D"
    else: return "F"

def get_current_semester():
    bd_tz = pytz.timezone('Asia/Dhaka')
    now = datetime.now(bd_tz)
    if 1 <= now.month <= 6:
        return f"Spring {now.year}"
    else:
        return f"Summer {now.year}"

# ============================================================================
# EEE EXCEL PARSER - FIXED FOR YOUR FORMAT
# ============================================================================
class EEEExcelParser:
    """
    Parser for EEE Department Excel files.
    Handles: Midterm Exam, Final Exam, Assignment, Analysis of CO, Analysis of PO
    """
    
    @staticmethod
    def parse(uploaded_file) -> dict:
        """Main parse function"""
        wb = load_workbook(uploaded_file, data_only=True)
        sheets = wb.sheetnames
        print(f"Found sheets: {sheets}")
        
        result = {
            'course_info': {
                'university': 'Stamford University Bangladesh',
                'trimester': '',
                'section': '',
                'course_code': '',
                'course_title': '',
                'teacher': ''
            },
            'students': {},
            'co_columns': ['CO1', 'CO2', 'CO3', 'CO4'],
            'po_columns': [],
            'co_po_mapping': {},
            'co_attainment': {},
            'po_attainment': {},
            'max_co_marks': {'CO1': 0, 'CO2': 0, 'CO3': 0, 'CO4': 0}
        }
        
        # Parse course info from Midterm Exam
        if 'Midterm Exam' in sheets:
            EEEExcelParser._parse_course_info(wb['Midterm Exam'], result)
        
        # Parse Analysis of CO (has combined CO marks from all exams)
        if 'Analysis of CO' in sheets:
            EEEExcelParser._parse_analysis_of_co(wb['Analysis of CO'], result)
        
        # Parse Analysis of PO (has CO-PO mapping matrix and individual PO)
        if 'Analysis of PO' in sheets:
            EEEExcelParser._parse_analysis_of_po(wb['Analysis of PO'], result)
        
        wb.close()
        
        # Calculate course statistics
        if result['students']:
            EEEExcelParser._calculate_statistics(result)
        
        return result
    
    @staticmethod
    def _parse_course_info(sheet, result):
        """Extract course info from rows 1-8"""
        for row in sheet.iter_rows(min_row=1, max_row=8, max_col=4, values_only=True):
            for cell in row:
                if cell and isinstance(cell, str) and ':' in cell:
                    key, val = cell.split(':', 1)
                    key = key.strip().lower()
                    val = val.strip().replace(':', '').strip()
                    if 'trimester' in key:
                        result['course_info']['trimester'] = val
                    elif 'section' in key:
                        result['course_info']['section'] = val
                    elif 'course code' in key:
                        result['course_info']['course_code'] = val
                    elif 'course title' in key:
                        result['course_info']['course_title'] = val
                    elif 'teacher' in key:
                        result['course_info']['teacher'] = val
    
    @staticmethod
    def _parse_analysis_of_co(sheet, result):
        """
        Parse Analysis of CO sheet.
        Structure:
        - Row 10: Headers (SL, Student ID, Name, Status, CO1, CO2, CO3, CO4, Total, CO1%, Y/N, CO2%, Y/N, CO3%, Y/N, CO4%, Y/N)
        - Row 11: Max marks for each CO (E11:H11)
        - Row 12-30: Student data
          A: SL, B: Student ID, C: Name, D: Status
          E: CO1 marks, F: CO2 marks, G: CO3 marks, H: CO4 marks
          I: Total CO Marks
          J: CO1%, K: CO1>=50%, L: CO2%, M: CO2>=50%, N: CO3%, O: CO3>=50%, P: CO4%, Q: CO4>=50%
        """
        student_id_pattern = re.compile(r'EEE\s*\d{3}\s*\d{5}', re.IGNORECASE)
        
        # Find the data start row (where SL = 1)
        data_start = None
        max_marks_row = None
        
        for row_idx in range(1, min(50, sheet.max_row + 1)):
            a_val = str(sheet.cell(row=row_idx, column=1).value or '').strip()
            
            # Find max marks row
            if a_val == '':
                e_val = sheet.cell(row=row_idx, column=5).value
                f_val = sheet.cell(row=row_idx, column=6).value
                g_val = sheet.cell(row=row_idx, column=7).value
                h_val = sheet.cell(row=row_idx, column=8).value
                if e_val and f_val and g_val and h_val:
                    try:
                        result['max_co_marks']['CO1'] = float(e_val)
                        result['max_co_marks']['CO2'] = float(f_val)
                        result['max_co_marks']['CO3'] = float(g_val)
                        result['max_co_marks']['CO4'] = float(h_val)
                        max_marks_row = row_idx
                    except:
                        pass
            
            # Find first student row
            if a_val == '1' or a_val == '1.0':
                data_start = row_idx
                break
        
        if not data_start:
            # Try to find by SL header
            for row_idx in range(1, min(50, sheet.max_row + 1)):
                a_val = str(sheet.cell(row=row_idx, column=1).value or '').strip().upper()
                if a_val == 'SL':
                    # Data starts 2 rows after SL header
                    data_start = row_idx + 2
                    break
        
        if not data_start:
            return
        
        # Parse each student row
        for row_idx in range(data_start, sheet.max_row + 1):
            sl = sheet.cell(row=row_idx, column=1).value
            if not sl:
                continue
            
            student_id_cell = sheet.cell(row=row_idx, column=2).value
            if not student_id_cell:
                continue
            
            student_id = str(student_id_cell).strip()
            
            # Validate student ID format
            if not student_id_pattern.search(student_id):
                if not re.match(r'^\d{3}\s*\d{5}$', student_id):
                    continue
            
            student_name = str(sheet.cell(row=row_idx, column=3).value or '').strip()
            student_status = str(sheet.cell(row=row_idx, column=4).value or '').strip()
            
            # Read CO marks from columns E, F, G, H
            co_marks = {}
            for col, co_name in [(5, 'CO1'), (6, 'CO2'), (7, 'CO3'), (8, 'CO4')]:
                val = sheet.cell(row=row_idx, column=col).value
                co_marks[co_name] = float(val) if val is not None else 0.0
            
            total_co_marks = sum(co_marks.values())
            
            # Read CO percentages from columns J, L, N, P
            co_pct = {}
            for col, co_name in [(10, 'CO1'), (12, 'CO2'), (14, 'CO3'), (16, 'CO4')]:
                val = sheet.cell(row=row_idx, column=col).value
                co_pct[co_name] = float(val) if val is not None else 0.0
            
            # Calculate total marks (scale to 100 if needed)
            max_total = sum(result['max_co_marks'].values())
            if max_total > 0:
                total_marks_scaled = (total_co_marks / max_total) * 100
            else:
                total_marks_scaled = total_co_marks
            
            result['students'][student_id] = {
                'id': student_id,
                'name': student_name,
                'status': student_status,
                'co_marks': co_marks,
                'total_marks': round(total_marks_scaled, 2),
                'co_attainment_pct': co_pct,
                'sgpa': calculate_sgpa(total_marks_scaled),
                'grade': get_grade_from_marks(total_marks_scaled),
                'status_final': 'Pass' if total_marks_scaled >= 40 else 'Fail'
            }
    
    @staticmethod
    def _parse_analysis_of_po(sheet, result):
        """
        Parse Analysis of PO sheet.
        Structure:
        - Row 10: CO-PO matrix headers
        - Row 11-14: CO-PO mapping (CO1-CO4 with weights for PO(a)-PO(l))
        - Row 19-22: Weightage calculation
        - Row 30-50: Individual student PO data
          A: SL, B: Student ID, C: Name, D: Status
          E: PO(a) marks, F: PO(b) marks, G: PO(d) marks
          H: PO(a)%, I: PO(a)>=50%, J: PO(b)%, K: PO(b)>=50%, L: PO(d)%, M: PO(d)>=50%
        """
        # Parse CO-PO mapping matrix (rows 11-14)
        co_po_mapping = {}
        po_headers = {}
        
        # Find PO headers row
        for row_idx in range(1, min(20, sheet.max_row + 1)):
            cell_b = str(sheet.cell(row=row_idx, column=2).value or '').strip()
            if 'CO-PO' in cell_b.upper() or 'CO-PO matrix' in cell_b.upper():
                # Read PO headers from columns D-O
                for col in range(4, 16):
                    header = str(sheet.cell(row=row_idx, column=col).value or '').strip()
                    if header and 'PO' in header.upper():
                        po_headers[col] = header.lower()
                
                # Read CO-PO mapping from next 4 rows
                for offset in range(1, 5):
                    co_row = row_idx + offset
                    co_name = str(sheet.cell(row=co_row, column=3).value or '').strip()
                    if co_name and co_name.upper().startswith('CO'):
                        co_name = co_name.upper()
                        mapping = {}
                        for col, po_name in po_headers.items():
                            val = sheet.cell(row=co_row, column=col).value
                            if val and float(val) > 0:
                                mapping[po_name] = float(val)
                        if mapping:
                            co_po_mapping[co_name] = mapping
                break
        
        result['co_po_mapping'] = co_po_mapping
        result['po_columns'] = sorted(set(
            po for mapping in co_po_mapping.values() for po in mapping.keys()
        ))
        
        # Parse individual student PO data (starts around row 30)
        student_id_pattern = re.compile(r'EEE\s*\d{3}\s*\d{5}', re.IGNORECASE)
        
        # Find student data section
        student_start = None
        for row_idx in range(25, min(60, sheet.max_row + 1)):
            a_val = str(sheet.cell(row=row_idx, column=1).value or '').strip()
            if a_val == '1' or a_val == '1.0':
                # Check if column B has a student ID
                b_val = str(sheet.cell(row=row_idx, column=2).value or '').strip()
                if student_id_pattern.search(b_val) or re.match(r'^\d{3}\s*\d{5}$', b_val):
                    student_start = row_idx
                    break
        
        if not student_start:
            # Try finding SL header
            for row_idx in range(25, min(60, sheet.max_row + 1)):
                a_val = str(sheet.cell(row=row_idx, column=1).value or '').strip().upper()
                if a_val == 'SL':
                    student_start = row_idx + 2
                    break
        
        if not student_start:
            return
        
        # Parse each student's PO data
        for row_idx in range(student_start, sheet.max_row + 1):
            sl = sheet.cell(row=row_idx, column=1).value
            if not sl:
                continue
            
            student_id_cell = sheet.cell(row=row_idx, column=2).value
            if not student_id_cell:
                continue
            
            student_id = str(student_id_cell).strip()
            if not student_id_pattern.search(student_id) and not re.match(r'^\d{3}\s*\d{5}$', student_id):
                continue
            
            if student_id not in result['students']:
                continue
            
            # Read PO marks from columns E, F, G
            po_marks = {}
            for col, po_name in [(5, 'po(a)'), (6, 'po(b)'), (7, 'po(d)')]:
                val = sheet.cell(row=row_idx, column=col).value
                po_marks[po_name] = float(val) if val is not None else 0.0
            
            # Read PO percentages from columns H, J, L
            po_pct = {}
            for col, po_name in [(8, 'po(a)'), (10, 'po(b)'), (12, 'po(d)')]:
                val = sheet.cell(row=row_idx, column=col).value
                po_pct[po_name] = float(val) if val is not None else 0.0
            
            result['students'][student_id]['po_marks'] = po_marks
            result['students'][student_id]['po_attainment_pct'] = po_pct
    
    @staticmethod
    def _calculate_statistics(result):
        """Calculate course-level CO and PO attainment"""
        students = result['students']
        
        # Calculate course-level CO attainment (average of individual CO%)
        co_attainment = {}
        for co in ['CO1', 'CO2', 'CO3', 'CO4']:
            scores = [s['co_attainment_pct'].get(co, 0) for s in students.values()]
            co_attainment[co] = round(np.mean(scores), 2) if scores else 0
        result['co_attainment'] = co_attainment
        
        # Calculate course-level PO attainment using CO-PO mapping
        co_po_mapping = result.get('co_po_mapping', {})
        if co_po_mapping and co_attainment:
            po_attainment = {}
            for po in result.get('po_columns', []):
                po_value = 0
                total_weight = 0
                for co, mappings in co_po_mapping.items():
                    if po in mappings:
                        po_value += co_attainment.get(co, 0) * mappings[po]
                        total_weight += mappings[po]
                if total_weight > 0:
                    po_attainment[po] = round(po_value / total_weight, 2)
                else:
                    po_attainment[po] = 0
            result['po_attainment'] = po_attainment
        
        # Calculate course statistics
        marks = [s['total_marks'] for s in students.values()]
        passing = len([m for m in marks if m >= 40])
        total = len(marks)
        
        result['course_stats'] = {
            'average_marks': round(np.mean(marks), 2),
            'highest_marks': round(max(marks), 2),
            'lowest_marks': round(min(marks), 2),
            'average_sgpa': round(np.mean([s['sgpa'] for s in students.values()]), 2),
            'total_students': total,
            'passing_students': passing,
            'pass_percentage': round((passing / total * 100) if total > 0 else 0, 1),
            'fail_percentage': round(((total - passing) / total * 100) if total > 0 else 0, 1),
            'std_deviation': round(np.std(marks), 2) if marks else 0
        }

# ============================================================================
# SESSION STATE INITIALIZATION
# ============================================================================
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'user_type' not in st.session_state: st.session_state.user_type = ""
if 'username' not in st.session_state: st.session_state.username = ""
if 'user_data' not in st.session_state: st.session_state.user_data = {}
if 'current_page' not in st.session_state: st.session_state.current_page = "login"
if 'results' not in st.session_state: st.session_state.results = {}
if 'processed' not in st.session_state: st.session_state.processed = False
if 'activity_log' not in st.session_state: st.session_state.activity_log = []
if 'teacher_uploads' not in st.session_state: st.session_state.teacher_uploads = {}
if 'ml_model' not in st.session_state: st.session_state.ml_model = None

# ============================================================================
# THEME
# ============================================================================
def apply_professional_theme():
    st.markdown("""
    <style>
    .stApp footer, footer, #MainMenu, header[data-testid="stHeader"] {
        display: none !important;
    }
    .main {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        font-family: 'Segoe UI', 'Inter', 'Roboto', sans-serif;
    }
    @keyframes gradientAnimation {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }
    .header {
        background: linear-gradient(-45deg, #ee7752, #e73c7e, #23a6d5, #23d5ab);
        background-size: 400% 400%;
        animation: gradientAnimation 15s ease infinite;
        color: white;
        padding: 1.8rem;
        border-radius: 0 0 25px 25px;
        margin-bottom: 2rem;
        text-align: center;
        box-shadow: 0 8px 32px rgba(31, 38, 135, 0.37);
    }
    .card {
        background: rgba(255, 255, 255, 0.95);
        border-radius: 20px;
        padding: 2rem;
        margin-bottom: 1.8rem;
        box-shadow: 0 8px 32px rgba(31, 38, 135, 0.15);
        transition: all 0.4s ease;
    }
    .card:hover {
        transform: translateY(-5px);
        box-shadow: 0 20px 40px rgba(31, 38, 135, 0.25);
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border-radius: 20px;
        padding: 1.5rem;
        text-align: center;
        box-shadow: 0 10px 30px rgba(102, 126, 234, 0.3);
    }
    .metric-value {
        font-size: 2.5rem;
        font-weight: 800;
        margin: 0.5rem 0;
    }
    .metric-label {
        font-size: 0.9rem;
        color: rgba(255,255,255,0.9);
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    .stButton > button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        padding: 0.8rem 1.5rem;
        border-radius: 12px;
        font-weight: 600;
        transition: all 0.3s ease;
        width: 100%;
    }
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 25px rgba(102, 126, 234, 0.4);
    }
    </style>
    """, unsafe_allow_html=True)

# ============================================================================
# SPIDER/ RADAR PLOT
# ============================================================================
def create_spider_plot(values, labels, title="PO Attainment Spider Plot"):
    """Create beautiful spider/ radar plot for PO attainment"""
    num_vars = len(labels)
    if num_vars < 3:
        return None
    
    angles = np.linspace(0, 2 * np.pi, num_vars, endpoint=False).tolist()
    values = list(values) + [values[0]]
    angles += angles[:1]
    
    fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(projection='polar'))
    ax.fill(angles, values, color='#667eea', alpha=0.25)
    ax.plot(angles, values, color='#667eea', linewidth=2, marker='o', markersize=8, markerfacecolor='#764ba2')
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(labels, fontsize=10)
    ax.set_ylim(0, 100)
    ax.set_yticks([20, 40, 60, 80, 100])
    ax.set_yticklabels(['20%', '40%', '60%', '80%', '100%'], fontsize=8, color='gray')
    ax.grid(True, alpha=0.3)
    ax.set_title(title, size=16, fontweight='bold', pad=20)
    
    for i, (angle, value) in enumerate(zip(angles[:-1], values[:-1])):
        ax.annotate(f'{value:.1f}%', xy=(angle, value), xytext=(5, 5),
                   textcoords='offset points', fontsize=9, fontweight='bold', color='#333')
    
    plt.tight_layout()
    return fig

# ============================================================================
# AUTHENTICATION
# ============================================================================
def load_users():
    default_users = {
        "admins": {
            "admin": {
                "username": "admin",
                "password": hash_password("admin123"),
                "email": "admin@stamford.edu.bd",
                "full_name": "System Administrator",
                "user_type": "admin",
                "is_active": True
            }
        },
        "teachers": {
            "teacher": {
                "username": "teacher",
                "password": hash_password("teacher123"),
                "email": "teacher@stamford.edu.bd",
                "full_name": "Teacher",
                "user_type": "teacher",
                "is_active": True
            }
        },
        "students": {},
        "parents": {}
    }
    
    try:
        if os.path.exists("users_enhanced.json"):
            with open("users_enhanced.json", 'r') as f:
                loaded_users = json.load(f)
            for user_type in default_users:
                if user_type not in loaded_users:
                    loaded_users[user_type] = {}
                for username, user_data in default_users[user_type].items():
                    if username not in loaded_users[user_type]:
                        loaded_users[user_type][username] = user_data
            return loaded_users
    except:
        pass
    
    return default_users

def save_users(users):
    try:
        with open("users_enhanced.json", 'w') as f:
            json.dump(users, f, indent=4)
        return True
    except:
        return False

def authenticate_user(username, password, user_type):
    users = load_users()
    user_category = user_type + "s"
    if user_category in users and username in users[user_category]:
        user_data = users[user_category][username]
        if verify_password(password, user_data["password"]):
            if not user_data.get("is_active", True):
                return False, "Account deactivated"
            return True, user_data
    return False, "Invalid credentials"

# ============================================================================
# PAGES
# ============================================================================

def show_login_page():
    apply_professional_theme()
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<h1 style='text-align:center;'>🎓 EduTrack Pro</h1>", unsafe_allow_html=True)
        st.markdown("<p style='text-align:center;color:#666;'>EEE Department | Stamford University Bangladesh</p>", unsafe_allow_html=True)
        
        with st.form("login_form"):
            user_type = st.selectbox("Account Type", ["admin", "teacher", "student", "parent"])
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            
            if st.form_submit_button("🔐 Login", use_container_width=True):
                if username and password:
                    success, user_data = authenticate_user(username, password, user_type)
                    if success:
                        st.session_state.logged_in = True
                        st.session_state.user_type = user_type
                        st.session_state.username = username
                        st.session_state.user_data = user_data
                        st.session_state.activity_log.append({
                            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "username": username,
                            "action": "login"
                        })
                        st.session_state.current_page = "upload" if user_type in ["teacher", "admin"] else "student_analytics"
                        st.rerun()
                    else:
                        st.error("Invalid credentials")
                else:
                    st.warning("Please enter credentials")

def show_upload_page():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("## 📤 Upload Student Data")
    st.markdown("Upload EEE Department Excel file with CO-PO analysis sheets")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        semester = st.text_input("Semester*", value=get_current_semester())
    with col2:
        course_code = st.text_input("Course Code*", value="EEE 321")
    with col3:
        course_name = st.text_input("Course Name*", value="Power System I")
    
    uploaded_file = st.file_uploader(
        "Choose Excel file (.xlsx or .xls)",
        type=['xlsx', 'xls'],
        help="Upload EEE Department Excel file with Analysis of CO and Analysis of PO sheets"
    )
    
    if uploaded_file:
        st.success(f"✅ File uploaded: {uploaded_file.name}")
        
        if st.button("🚀 Process & Analyze Data", type="primary", use_container_width=True):
            if not semester or not course_code:
                st.error("Please fill semester and course code")
            else:
                with st.spinner("🔄 Parsing Excel file..."):
                    try:
                        parsed = EEEExcelParser.parse(uploaded_file)
                        
                        if parsed and parsed.get('students'):
                            student_count = len(parsed['students'])
                            st.success(f"✅ Successfully parsed {student_count} students!")
                            
                            st.session_state.results = parsed
                            st.session_state.processed = True
                            
                            # Save course data
                            save_path = Config.DATA_DIR / f"course_{semester.replace(' ', '_')}_{course_code.replace(' ', '_')}.pkl"
                            with open(save_path, 'wb') as f:
                                pickle.dump(parsed, f)
                            
                            # Track upload
                            if st.session_state.username not in st.session_state.teacher_uploads:
                                st.session_state.teacher_uploads[st.session_state.username] = []
                            st.session_state.teacher_uploads[st.session_state.username].append({
                                'semester': semester,
                                'course_code': course_code,
                                'filename': uploaded_file.name,
                                'student_count': student_count,
                                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            })
                            
                            # Show summary metrics
                            stats = parsed.get('course_stats', {})
                            col_m1, col_m2, col_m3, col_m4 = st.columns(4)
                            with col_m1:
                                st.metric("📚 Students", stats.get('total_students', 0))
                            with col_m2:
                                st.metric("📊 Avg Marks", f"{stats.get('average_marks', 0):.1f}")
                            with col_m3:
                                st.metric("✅ Pass Rate", f"{stats.get('pass_percentage', 0):.1f}%")
                            with col_m4:
                                st.metric("🎯 Avg SGPA", f"{stats.get('average_sgpa', 0):.2f}")
                        else:
                            st.error("❌ No student data found. Please check the Excel file format.")
                            
                    except Exception as e:
                        st.error(f"❌ Error parsing file: {str(e)}")
                        import traceback
                        st.code(traceback.format_exc())
    
    st.markdown('</div>', unsafe_allow_html=True)

def show_course_reports():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("## 📊 Course Reports & Analytics")
    
    # Load all saved courses
    all_courses = {}
    for file in Config.DATA_DIR.glob("course_*.pkl"):
        try:
            with open(file, 'rb') as f:
                course_data = pickle.load(f)
                key = f"{course_data.get('course_info', {}).get('trimester', 'N/A')} - {course_data.get('course_info', {}).get('course_code', 'N/A')}"
                all_courses[key] = course_data
        except:
            continue
    
    # Also check session state
    if st.session_state.processed and st.session_state.results:
        results = st.session_state.results
        key = f"{results.get('course_info', {}).get('trimester', 'N/A')} - {results.get('course_info', {}).get('course_code', 'N/A')}"
        all_courses[key] = results
    
    if not all_courses:
        st.info("📭 No course data available. Please upload data first.")
        if st.button("Go to Upload Page", use_container_width=True):
            st.session_state.current_page = "upload"
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
        return
    
    selected_course = st.selectbox("Select Course", list(all_courses.keys()))
    
    if selected_course:
        course_data = all_courses[selected_course]
        stats = course_data.get('course_stats', {})
        students = course_data.get('students', {})
        
        # KPI Metrics
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("📚 Total Students", stats.get('total_students', 0))
        with col2:
            st.metric("📊 Average Marks", f"{stats.get('average_marks', 0):.1f}")
        with col3:
            st.metric("✅ Pass Rate", f"{stats.get('pass_percentage', 0):.1f}%")
        with col4:
            st.metric("🎯 Average SGPA", f"{stats.get('average_sgpa', 0):.2f}")
        
        # Tabs
        tab1, tab2, tab3, tab4 = st.tabs(["📊 Dashboard", "📈 Marks Distribution", "🎯 CO-PO Attainment", "📋 Student List"])
        
        with tab1:
            st.markdown("### Real-Time Dashboard")
            
            # Grade Distribution Pie Chart
            grades = [s.get('grade', 'F') for s in students.values()]
            grade_counts = pd.Series(grades).value_counts()
            
            fig_pie = go.Figure(data=[go.Pie(
                labels=grade_counts.index.tolist(),
                values=grade_counts.values.tolist(),
                hole=0.4,
                marker=dict(colors=['#4CAF50', '#8BC34A', '#CDDC39', '#FFC107', '#FF9800', '#FF5722', '#9C27B0', '#673AB7', '#3F51B5', '#F44336'])
            )])
            fig_pie.update_layout(height=400, title="Grade Distribution")
            st.plotly_chart(fig_pie, use_container_width=True)
            
            # Marks Distribution Histogram
            marks = [s.get('total_marks', 0) for s in students.values()]
            fig_hist = px.histogram(
                x=marks, nbins=10, title="Marks Distribution",
                labels={'x': 'Total Marks', 'y': 'Number of Students'},
                color_discrete_sequence=['#667eea']
            )
            fig_hist.update_layout(template='plotly_white', height=400)
            st.plotly_chart(fig_hist, use_container_width=True)
        
        with tab2:
            st.markdown("### Marks Distribution Analysis")
            marks = [s.get('total_marks', 0) for s in students.values()]
            
            col_m1, col_m2, col_m3, col_m4 = st.columns(4)
            with col_m1:
                st.metric("Highest", f"{max(marks):.1f}")
            with col_m2:
                st.metric("Lowest", f"{min(marks):.1f}")
            with col_m3:
                st.metric("Std Dev", f"{np.std(marks):.2f}")
            with col_m4:
                st.metric("Median", f"{np.median(marks):.1f}")
            
            fig_box = px.box(y=marks, title="Marks Box Plot", color_discrete_sequence=['#667eea'])
            fig_box.update_layout(template='plotly_white', height=400)
            st.plotly_chart(fig_box, use_container_width=True)
        
        with tab3:
            st.markdown("### CO-PO Attainment Analysis")
            
            col_co, col_po = st.columns(2)
            
            with col_co:
                co_attainment = course_data.get('co_attainment', {})
                if co_attainment:
                    st.markdown("#### CO Attainment (Course Average)")
                    cos = list(co_attainment.keys())
                    vals = list(co_attainment.values())
                    
                    fig_co = go.Figure(data=[go.Bar(
                        x=cos, y=vals,
                        text=[f'{v:.1f}%' for v in vals],
                        textposition='auto',
                        marker_color=['#4CAF50' if v >= 80 else '#FFC107' if v >= 60 else '#F44336' for v in vals]
                    )])
                    fig_co.update_layout(template='plotly_white', height=350, yaxis_title="Percentage (%)")
                    st.plotly_chart(fig_co, use_container_width=True)
                else:
                    st.info("No CO attainment data available")
            
            with col_po:
                po_attainment = course_data.get('po_attainment', {})
                if po_attainment:
                    st.markdown("#### PO Attainment (Course Average)")
                    pos = list(po_attainment.keys())
                    po_vals = list(po_attainment.values())
                    
                    if len(pos) >= 3:
                        fig = create_spider_plot(po_vals, pos, "PO Attainment Spider Plot")
                        if fig:
                            st.pyplot(fig)
                    else:
                        fig_po = go.Figure(data=[go.Bar(
                            x=pos, y=po_vals,
                            text=[f'{v:.1f}%' for v in po_vals],
                            textposition='auto',
                            marker_color='#764ba2'
                        )])
                        fig_po.update_layout(template='plotly_white', height=350, yaxis_title="Percentage (%)")
                        st.plotly_chart(fig_po, use_container_width=True)
                else:
                    st.info("No PO attainment data available")
        
        with tab4:
            st.markdown("### Student Results")
            
            search_term = st.text_input("🔍 Search by Name or ID", placeholder="Type to filter...")
            
            sorted_students = sorted(students.items(), key=lambda x: x[1].get('total_marks', 0), reverse=True)
            
            if search_term:
                sorted_students = [(sid, s) for sid, s in sorted_students
                                  if search_term.lower() in sid.lower() or
                                  search_term.lower() in s.get('name', '').lower()]
            
            data = []
            for rank, (sid, s) in enumerate(sorted_students, 1):
                data.append({
                    'Rank': rank,
                    'Student ID': sid,
                    'Name': s.get('name', 'N/A'),
                    'Total Marks': f"{s.get('total_marks', 0):.1f}",
                    'Grade': s.get('grade', 'N/A'),
                    'SGPA': f"{s.get('sgpa', 0):.2f}",
                    'Status': s.get('status_final', 'N/A')
                })
            
            if data:
                df = pd.DataFrame(data)
                st.dataframe(df, use_container_width=True, height=500, hide_index=True)
                
                csv = df.to_csv(index=False)
                st.download_button("📥 Download CSV", csv, f"student_list_{course_data.get('course_info', {}).get('course_code', 'COURSE')}.csv", "text/csv")
    
    st.markdown('</div>', unsafe_allow_html=True)

def show_teacher_panel():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("## 👨‍🏫 Teacher Panel")
    
    tab1, tab2 = st.tabs(["👥 Create Student & Parent Account", "📤 Upload History"])
    
    with tab1:
        with st.form("create_accounts"):
            st.markdown("### Student Information")
            col1, col2 = st.columns(2)
            with col1:
                student_id = st.text_input("Student ID*", placeholder="EEE 078 07759")
                student_name = st.text_input("Student Full Name*")
                student_username = st.text_input("Student Username*")
                student_password = st.text_input("Student Password*", type="password")
            with col2:
                student_semester = st.text_input("Semester*", value=get_current_semester())
                student_email = st.text_input("Student Email")
            
            st.markdown("---")
            st.markdown("### Parent/Guardian Information")
            col3, col4 = st.columns(2)
            with col3:
                parent_name = st.text_input("Parent Full Name*")
                parent_username = st.text_input("Parent Username*")
                parent_password = st.text_input("Parent Password*", type="password")
            with col4:
                parent_email = st.text_input("Parent Email*")
                parent_contact = st.text_input("Parent Contact Number*")
            
            if st.form_submit_button("✅ Create Accounts", use_container_width=True, type="primary"):
                errors = []
                if not student_id: errors.append("Student ID required")
                if not student_name: errors.append("Student Name required")
                if not student_username: errors.append("Student Username required")
                if not student_password: errors.append("Student Password required")
                if not parent_name: errors.append("Parent Name required")
                if not parent_email: errors.append("Parent Email required")
                if not parent_username: errors.append("Parent Username required")
                if not parent_password: errors.append("Parent Password required")
                
                if errors:
                    for e in errors:
                        st.error(f"❌ {e}")
                else:
                    users = load_users()
                    
                    if student_username in users.get('students', {}):
                        st.error(f"Username '{student_username}' already exists!")
                    elif parent_username in users.get('parents', {}):
                        st.error(f"Username '{parent_username}' already exists!")
                    else:
                        users['students'][student_username] = {
                            "username": student_username,
                            "password": hash_password(student_password),
                            "email": student_email or parent_email,
                            "full_name": student_name,
                            "student_id": student_id,
                            "user_type": "student",
                            "is_active": True,
                            "semester": student_semester,
                            "parent_email": parent_email,
                            "parent_contact": parent_contact,
                            "parent_name": parent_name,
                            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "created_by": st.session_state.username
                        }
                        
                        users['parents'][parent_username] = {
                            "username": parent_username,
                            "password": hash_password(parent_password),
                            "email": parent_email,
                            "full_name": parent_name,
                            "student_linked": student_id,
                            "user_type": "parent",
                            "is_active": True,
                            "contact_no": parent_contact,
                            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "created_by": st.session_state.username
                        }
                        
                        if save_users(users):
                            st.success(f"✅ Accounts created! Student: `{student_username}`, Parent: `{parent_username}`")
                            st.session_state.activity_log.append({
                                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "username": st.session_state.username,
                                "action": "created_accounts",
                                "details": f"Student: {student_username}"
                            })
                            time.sleep(1.5)
                            st.rerun()
    
    with tab2:
        st.markdown("### Upload History")
        uploads = st.session_state.teacher_uploads.get(st.session_state.username, [])
        if uploads:
            df = pd.DataFrame(reversed(uploads))
            st.dataframe(df, use_container_width=True)
        else:
            st.info("No uploads yet")
    
    st.markdown('</div>', unsafe_allow_html=True)

def show_admin_panel():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("## 🔧 Admin Panel")
    
    tab1, tab2, tab3 = st.tabs(["👥 User Management", "📊 System Overview", "📝 Activity Log"])
    
    with tab1:
        users = load_users()
        user_type = st.selectbox("Filter by Type", ["teachers", "students", "parents", "admins"])
        
        if user_type in users:
            user_list = []
            for u, d in users[user_type].items():
                user_list.append({
                    'Username': u,
                    'Name': d.get('full_name', ''),
                    'Email': d.get('email', ''),
                    'Status': 'Active' if d.get('is_active', True) else 'Inactive',
                    'Created': d.get('created_at', 'N/A')
                })
            
            if user_list:
                df = pd.DataFrame(user_list)
                st.dataframe(df, use_container_width=True, hide_index=True)
            else:
                st.info(f"No {user_type} found")
    
    with tab2:
        st.markdown("### System Overview")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            total_users = sum(len(v) for v in users.values())
            st.metric("Total Users", total_users)
        with col2:
            course_files = list(Config.DATA_DIR.glob("course_*.pkl"))
            st.metric("Courses Stored", len(course_files))
        with col3:
            st.metric("Activity Logs", len(st.session_state.activity_log))
    
    with tab3:
        st.markdown("### Activity Log")
        logs = st.session_state.activity_log
        if logs:
            df = pd.DataFrame(reversed(logs))
            st.dataframe(df, use_container_width=True, height=400)
        else:
            st.info("No activity logged yet")
    
    st.markdown('</div>', unsafe_allow_html=True)

def show_sidebar():
    with st.sidebar:
        st.markdown("<h1 style='text-align:center;'>🎓</h1>", unsafe_allow_html=True)
        st.markdown("<h3 style='text-align:center;'>EduTrack Pro</h3>", unsafe_allow_html=True)
        
        st.markdown(f"""
        <div style="background: #667eea; color: white; padding: 1rem; border-radius: 10px; margin-bottom: 1rem;">
            <strong>{st.session_state.user_data.get('full_name', 'User')}</strong><br>
            <small>{st.session_state.user_type.title()}</small>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("### 📍 Navigation")
        
        if st.session_state.user_type in ["teacher", "admin"]:
            if st.button("📤 Upload Data", use_container_width=True):
                st.session_state.current_page = "upload"
                st.rerun()
            if st.button("📊 Course Reports", use_container_width=True):
                st.session_state.current_page = "reports"
                st.rerun()
        
        if st.session_state.user_type == "teacher":
            if st.button("👨‍🏫 Teacher Panel", use_container_width=True):
                st.session_state.current_page = "teacher_panel"
                st.rerun()
        
        if st.session_state.user_type == "admin":
            if st.button("🔧 Admin Panel", use_container_width=True):
                st.session_state.current_page = "admin_panel"
                st.rerun()
        
        st.markdown("---")
        if st.button("🚪 Logout", use_container_width=True):
            for key in ['logged_in', 'user_type', 'username', 'user_data']:
                st.session_state[key] = False if key == 'logged_in' else ""
            st.session_state.current_page = "login"
            st.rerun()

# ============================================================================
# MAIN APP
# ============================================================================
def main():
    st.set_page_config(
        page_title=Config.APP_NAME,
        page_icon="🎓",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    init_db()
    apply_professional_theme()
    
    if not st.session_state.logged_in:
        show_login_page()
        return
    
    show_sidebar()
    
    page = st.session_state.current_page
    
    if page == "upload" and st.session_state.user_type in ["teacher", "admin"]:
        show_upload_page()
    elif page == "reports" and st.session_state.user_type in ["teacher", "admin"]:
        show_course_reports()
    elif page == "teacher_panel" and st.session_state.user_type == "teacher":
        show_teacher_panel()
    elif page == "admin_panel" and st.session_state.user_type == "admin":
        show_admin_panel()
    else:
        show_upload_page() if st.session_state.user_type in ["teacher", "admin"] else show_course_reports()
    
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; padding: 10px; color: #666;">
        <p>EduTrack Pro 2026 | EEE Department | Stamford University Bangladesh</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    Config.DATA_DIR.mkdir(parents=True, exist_ok=True)
    main()

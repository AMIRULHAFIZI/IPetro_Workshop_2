import os
import re
import json
import click
import uuid 
import time
import copy
import io
from pptx.enum.shapes import MSO_SHAPE_TYPE
from flask import Flask, request, render_template, send_from_directory, send_file, redirect, url_for, flash
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from flask_migrate import Migrate
from datetime import datetime, time
from sqlalchemy import func
from dotenv import load_dotenv

import pandas as pd
import pdf2image
from PIL import Image
import io
import requests
import base64
from openpyxl import load_workbook
# --- IMPORT ADDED: Needed for centering text in merged cells ---
from openpyxl.styles import Alignment 
import shutil
import time as sleep_time 

from flask_wtf import FlaskForm
from wtforms import StringField, PasswordField, SubmitField, SelectField, TextAreaField, BooleanField
from wtforms.validators import DataRequired, Email, EqualTo, Length, ValidationError, Optional
from email_validator import validate_email, EmailNotValidError
from flask_wtf.file import FileField, FileAllowed
from werkzeug.utils import secure_filename
from functools import wraps
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# --- LOAD API KEY ---
load_dotenv()
GEMINI_API_KEY = os.environ.get('GEMINI_API_KEY')
if not GEMINI_API_KEY:
    raise ValueError("CRITICAL ERROR: 'GEMINI_API_KEY' not set.")

# --- APP SETUP ---
app = Flask(__name__)
app.config['SECRET_KEY'] = 'a-super-secret-key-that-is-hard-to-guess'
app.config['SQLALCHEMY_DATABASE_URI'] = os.environ.get('DATABASE_URL')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)
migrate = Migrate(app, db)

login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'
login_manager.login_message = 'Please log in to access this page.'
login_manager.login_message_category = 'error'

# --- FOLDER CONFIG ---
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
ANNOUNCEMENT_FOLDER = 'announcements'
POPPLER_PATH = r'C:\poppler-25.07.0\Library\bin' # Verify this path on your machine!

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(ANNOUNCEMENT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['ANNOUNCEMENT_FOLDER'] = ANNOUNCEMENT_FOLDER

# --- EXCEL TEMPLATE CONFIG ---
TEMPLATE_FILE = 'EXCEL.xlsx'
TARGET_SHEET = 'Sheet1'
template_column_order = [
    'NO.', 'EQUIPMENT NO. ', 'PMT NO.', 'EQUIPMENT DESCRIPTION', 'PARTS',
    'PHASE', 'FLUID', 'TYPE', 'SPEC.', 'GRADE',
    'INSULATION',
    'TEMP. (°C) ', 'PRESSURE (Mpa) ', 
    'TEMP. (°C)', 'PRESSURE (Mpa)' 
]

def replace_text_exact(slide, search_str, new_str):
    """
    Searches for specific text in all text boxes and replaces it.
    """
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    if search_str in run.text:
                        run.text = run.text.replace(search_str, str(new_str))

# ==================== ROUTES ==================

# ==================== HELPER FUNCTIONS ==================

def _set_cell_text(cell, text, font_size=10):
    """
    Helper to set text in a table cell with specific font size and alignment.
    """
    # 1. Clean the text
    if text is None:
        text = "-"
    
    text_str = str(text).strip()
    
    if text_str == '' or text_str.lower() == 'nan':
        text_str = "-"
        
    # 2. Set the text
    cell.text_frame.text = text_str
    
    # 3. Apply styling to every paragraph in the cell
    for paragraph in cell.text_frame.paragraphs:
        paragraph.font.size = Pt(font_size)
        paragraph.font.name = 'Arial'  # Standard font
        
    # 4. Center vertically
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE

def add_new_row(table):
    """
    Copies the last row of the table and appends it to the end.
    """
    import copy # This is required for the deepcopy to work
    
    # 1. Get the last row's XML element
    new_row_xml = copy.deepcopy(table._tbl.tr_lst[-1])
    
    # 2. Append it to the table's XML list
    table._tbl.append(new_row_xml)
    
    # 3. Clear the text in the new row so it's blank
    new_row = table.rows[len(table.rows) - 1]
    for cell in new_row.cells:
        cell.text_frame.text = ""
        
    return new_row

# --- HELPER TO DUPLICATE SLIDE (HANDLES IMAGES CORRECTLY) ---
def duplicate_slide(pres, index):
    """
    Robustly duplicate the slide at 'index', handling images correctly by 
    extracting their data and re-adding them.
    """
    source = pres.slides[index]
    # Create new slide with same layout
    dest = pres.slides.add_slide(source.slide_layout)

    # Remove any default placeholders/shapes on the new slide so it's clean
    for shp in dest.shapes:
        shp.element.getparent().remove(shp.element)

    # Iterate through source shapes to preserve order
    for shp in source.shapes:
        if shp.shape_type == MSO_SHAPE_TYPE.PICTURE:
            # --- IT'S AN IMAGE: Handle specially ---
            try:
                # 1. Extract raw image data
                blob = shp.image.blob
                # 2. Wrap data in a stream so pptx can read it
                image_stream = io.BytesIO(blob)
                # 3. Add as a NEW picture shape at same coordinates
                dest.shapes.add_picture(
                    image_stream, shp.left, shp.top, shp.width, shp.height
                )
            except Exception as e:
                print(f"Warning: Could not copy image on slide duplication: {e}")
        else:
            # --- IT'S A TABLE/TEXTBOX/SHAPE: Deepcopy works ---
            new_el = copy.deepcopy(shp.element)
            dest.shapes._spTree.append(new_el)

    return dest
# ==================== MAIN ROUTE ==================

@app.route('/generate_ppt/<filename>')
@login_required
def generate_ppt(filename):
    # 1. Path Setup
    excel_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
    template_path = os.path.join(app.root_path, 'Inspection Plan Powerpoint Template.pptx')
    output_pptx_name = f"Inspection_Plan_{filename.replace('.xlsx', '.pptx')}"
    output_pptx_path = os.path.join(app.config['OUTPUT_FOLDER'], output_pptx_name)

    # 2. Check Files
    if not os.path.exists(excel_path):
        flash("Excel file missing.", "error")
        return redirect(url_for('index'))
    if not os.path.exists(template_path):
        flash("PPT template missing.", "error")
        return redirect(url_for('index'))

    # ==========================================
    # PASS 1: EXTRACT METADATA (Tag, PMT, Desc)
    # ==========================================
    tag_val, pmt_val, desc_val = "", "", ""
    try:
        df_meta = pd.read_excel(excel_path, header=None, nrows=15)
        for r_idx, row in df_meta.iterrows():
            for c_idx, cell in enumerate(row):
                s = str(cell).strip().upper()
                if any(x in s for x in ["TAG NO", "TAG. NO", "EQUIPMENT NO"]) and not tag_val:
                    if c_idx+1 < len(row) and str(row[c_idx+1]).strip() not in ['nan', '']: tag_val = str(row[c_idx+1])
                    elif c_idx+2 < len(row): tag_val = str(row[c_idx+2])
                if any(x in s for x in ["PMT NO", "PMT. NO"]) and not pmt_val:
                    if c_idx+1 < len(row) and str(row[c_idx+1]).strip() not in ['nan', '']: pmt_val = str(row[c_idx+1])
                    elif c_idx+2 < len(row): pmt_val = str(row[c_idx+2])
                if any(x in s for x in ["DESCRIPTION", "DESC."]) and not desc_val:
                    if c_idx+1 < len(row) and str(row[c_idx+1]).strip() not in ['nan', '']: desc_val = str(row[c_idx+1])
                    elif c_idx+2 < len(row): desc_val = str(row[c_idx+2])
    except: pass

    # ==========================================
    # PASS 2: CONTENT-AWARE SCANNING (Victory Logic)
    # ==========================================
    try:
        df_raw = pd.read_excel(excel_path, header=None)
        
        # 1. Find the Anchor Row (Start of Data Table)
        anchor_idx = None
        for idx, row in df_raw.iterrows():
            r_str = row.astype(str).str.upper().tolist()
            if "TYPE" in r_str and "SPEC." in r_str:
                anchor_idx = idx
                break
        
        if anchor_idx is None:
            flash("Could not find table headers (TYPE/SPEC).", "error")
            return redirect(url_for('index'))

        # 2. Slice Data
        df_data = df_raw.iloc[anchor_idx + 1:].copy().reset_index(drop=True)
        df_data = df_data.fillna('-')
        
        # Get Header Rows for keyword matching
        header_row_sub = df_raw.iloc[anchor_idx].astype(str).str.upper().tolist()
        header_row_main = []
        if anchor_idx > 0:
            header_row_main = df_raw.iloc[anchor_idx-1].astype(str).str.upper().tolist()

        # 3. IDENTIFY COLUMNS
        idx_fluid, idx_part = None, None
        idx_type, idx_spec, idx_grade = None, None, None
        idx_insul = None
        idx_temp, idx_press = None, None

        num_cols = df_data.shape[1]
        
        # A. Find PARTS (Look for words like HEAD, SHELL, FLANGE)
        # Avoid columns that are just Tag Numbers (V-001)
        best_part_score = -1
        for c in range(num_cols):
            score = 0
            for val in df_data.iloc[:10, c]:
                v = str(val).upper()
                if any(k in v for k in ['HEAD', 'SHELL', 'PIPE', 'FLANGE', 'PLATE']): score += 2
                if 'V-' in v: score -= 5 # Penalty for Tag No column
            
            # Header boost
            if c < len(header_row_main) and 'PARTS' in header_row_main[c]: score += 10
            
            if score > best_part_score:
                best_part_score = score
                idx_part = c

        # B. Find MATERIAL TYPE (Stainless, Carbon)
        best_type_score = -1
        for c in range(num_cols):
            score = 0
            for val in df_data.iloc[:10, c]:
                v = str(val).upper()
                if any(k in v for k in ['STAINLESS', 'CARBON', 'STEEL']): score += 2
            
            # Header boost
            if c < len(header_row_sub) and 'TYPE' in header_row_sub[c]: score += 10
            
            if score > best_type_score:
                best_type_score = score
                idx_type = c

        # C. Find FLUID (Condensate, Water, Gas - or Header 'FLUID')
        # Check Headers First!
        for c in range(min(len(header_row_main), num_cols)):
            if "FLUID" in header_row_main[c] or "MEDIA" in header_row_main[c]:
                idx_fluid = c
                break
        
        # Fallback: Look for fluid names if header failed
        if idx_fluid is None:
            best_fluid_score = -1
            for c in range(num_cols):
                if c == idx_part or c == idx_type: continue
                score = 0
                for val in df_data.iloc[:10, c]:
                    v = str(val).upper()
                    if any(k in v for k in ['WATER', 'GAS', 'CONDENSATE', 'OIL', 'AIR']): score += 2
                if score > best_fluid_score and score > 0:
                    best_fluid_score = score
                    idx_fluid = c
            # Last resort: Column 0
            if idx_fluid is None: idx_fluid = 0

        # D. Find INSULATION (N/A, YES, NO)
        # Scan Headers First
        for c in range(min(len(header_row_main), num_cols)):
            if "INSUL" in header_row_main[c]: 
                idx_insul = c
                break
        
        # Scan Data content if header missing
        if idx_insul is None:
            best_ins_score = -1
            for c in range(num_cols):
                if c in [idx_part, idx_type, idx_fluid]: continue
                score = 0
                for val in df_data.iloc[:10, c]:
                    v = str(val).upper()
                    # Strong signal for insulation column
                    if v in ['N/A', 'YES', 'NO', 'NA']: score += 3
                if score > best_ins_score and score > 0:
                    best_ins_score = score
                    idx_insul = c

        # E. Find SPEC and GRADE (Relative to Type usually)
        # Scan Headers first
        for c in range(min(len(header_row_sub), num_cols)):
            if "SPEC" in header_row_sub[c]: idx_spec = c
            if "GRADE" in header_row_sub[c] or "GR." in header_row_sub[c] or header_row_sub[c] == "GR": idx_grade = c
            
        # Fallback relative to Type
        if idx_type is not None:
            if idx_spec is None: idx_spec = idx_type + 1
            if idx_grade is None: idx_grade = idx_type + 2
            
        # F. Find TEMP / PRESS (Look for Units)
        for c in range(num_cols):
            # Header Check
            h_main = header_row_main[c] if c < len(header_row_main) else ""
            h_sub = header_row_sub[c] if c < len(header_row_sub) else ""
            full_head = (h_main + " " + h_sub)
            
            # Skip Design columns if possible
            if "DESIGN" in full_head: continue

            if "TEMP" in full_head or "T (C)" in full_head: idx_temp = c
            if "PRESS" in full_head or "P (MPA)" in full_head: idx_press = c

        # 4. Filter Data (Use valid part column)
        if idx_part is not None:
             df_data = df_data[
                (df_data[idx_part].astype(str).str.strip() != '-') & 
                (df_data[idx_part].astype(str).str.strip() != '') &
                (df_data[idx_part].astype(str).str.strip() != 'nan') &
                (df_data[idx_part].astype(str).str.strip().str.upper() != 'COMPONENT') 
            ]

    except Exception as e:
        flash(f"Error processing data: {e}", "error")
        return redirect(url_for('index'))

    # ==========================================
    # GENERATE POWERPOINT
    # ==========================================
    prs = Presentation(template_path)
    MAX_ROWS = 5
    data_rows = [row for _, row in df_data.iterrows()]
    chunks = [data_rows[i:i + MAX_ROWS] for i in range(0, len(data_rows), MAX_ROWS)] if data_rows else [[]]

    for i, chunk in enumerate(chunks):
        if i == 0: slide = prs.slides[0]
        else: slide = duplicate_slide(prs, 0)
        
        # Fill Header
        for shape in slide.shapes:
            if not shape.has_table: continue
            tbl = shape.table
            try:
                r0 = " ".join([c.text_frame.text.upper() for c in tbl.rows[0].cells])
                if "TAG" in r0 or "DESC" in r0:
                    for row in tbl.rows:
                        for idx, cell in enumerate(row.cells):
                            txt = cell.text_frame.text.strip().upper()
                            if "TAG" in txt and idx+1 < len(row.cells): _set_cell_text(row.cells[idx+1], tag_val)
                            elif "PMT" in txt and idx+1 < len(row.cells): _set_cell_text(row.cells[idx+1], pmt_val)
                            elif "DESCRIPTION" in txt and idx+1 < len(row.cells): _set_cell_text(row.cells[idx+1], desc_val)
            except: pass

        # Fill Data
        main_table = None
        for shape in slide.shapes:
            if shape.has_table:
                try:
                    r0 = " ".join([c.text_frame.text.upper() for c in shape.table.rows[0].cells])
                    if "COMPONENT" in r0 and "FLUID" in r0:
                        main_table = shape.table
                        break
                except: continue
        
        if main_table:
            start_row = 2
            while len(main_table.rows) < (start_row + MAX_ROWS):
                add_new_row(main_table)
            
            for idx_in_chunk, row_data in enumerate(chunk):
                curr_idx = start_row + idx_in_chunk
                cells = main_table.rows[curr_idx].cells
                
                # USE INTELLIGENT INDICES
                val_fluid = row_data.get(idx_fluid, '') if idx_fluid is not None else ''
                _set_cell_text(cells[0], val_fluid, 9)
                
                val_part = row_data.get(idx_part, '') if idx_part is not None else ''
                _set_cell_text(cells[1], val_part, 9)
                
                _set_cell_text(cells[2], "", 9) 
                
                _set_cell_text(cells[3], row_data.get(idx_type, ''), 9)
                _set_cell_text(cells[4], row_data.get(idx_spec, ''), 9)
                _set_cell_text(cells[5], row_data.get(idx_grade, ''), 9)
                
                val_ins = row_data.get(idx_insul, '') if idx_insul is not None else ''
                _set_cell_text(cells[6], val_ins, 9)
                
                val_temp = row_data.get(idx_temp, '') if idx_temp is not None else ''
                _set_cell_text(cells[7], val_temp, 9)
                
                val_press = row_data.get(idx_press, '') if idx_press is not None else ''
                _set_cell_text(cells[8], val_press, 9)

            for j in range(len(chunk), MAX_ROWS):
                curr_idx = start_row + j
                if curr_idx < len(main_table.rows):
                    for cell in main_table.rows[curr_idx].cells:
                        _set_cell_text(cell, "", 9)

    prs.save(output_pptx_path)
    return send_file(output_pptx_path, as_attachment=True)
# ==================== DATABASE MODELS ==================

class Role(db.Model):
    __tablename__ = 'roles'
    role_id = db.Column(db.Integer, primary_key=True)
    role_name = db.Column(db.String(50), unique=True, nullable=False)
    users = db.relationship('User', backref='role', lazy=True)

class User(db.Model, UserMixin):
    __tablename__ = 'users'
    user_id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(256), nullable=False)
    name = db.Column(db.String(100), nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    phone_num = db.Column(db.String(20), nullable=True)
    role_id = db.Column(db.Integer, db.ForeignKey('roles.role_id'), nullable=False)
    
    equipment_data = db.relationship('EquipmentData', backref='user', lazy=True)
    history_entries = db.relationship('History', backref='user', lazy=True)
    announcements = db.relationship('Announcement', backref='user', lazy=True)

    def get_id(self):
        return (self.user_id)
    def set_password(self, password):
        self.password_hash = generate_password_hash(password)
    def check_password(self, password):
        return check_password_hash(self.password_hash, password)
    def __repr__(self):
        return f'<User {self.username}>'

class History(db.Model):
    __tablename__ = 'history'
    history_id = db.Column(db.Integer, primary_key=True)
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    excel_filename = db.Column(db.String(255), nullable=True)
    created_by_user_id = db.Column(db.Integer, db.ForeignKey('users.user_id'), nullable=False)
    equipment_data_rows = db.relationship('EquipmentData', backref='history_entry', lazy=True)

class EquipmentData(db.Model):
    __tablename__ = 'equipment_data'
    id = db.Column(db.Integer, primary_key=True)
    source_drawing = db.Column(db.String(255), nullable=False)
    part_name = db.Column(db.String(100), nullable=False)
    fluid = db.Column(db.String(100))
    material_type = db.Column(db.String(100))
    material_spec = db.Column(db.String(100))
    material_grade = db.Column(db.String(100))
    design_temp = db.Column(db.String(50))
    design_pressure = db.Column(db.String(50))
    operating_temp = db.Column(db.String(50))
    operating_pressure = db.Column(db.String(50))
    insulation = db.Column(db.String(50))
    no = db.Column(db.String(50), default='')
    equipment_no = db.Column(db.String(50), default='')
    pmt_no = db.Column(db.String(50), default='')
    equipment_description = db.Column(db.String(255), default='')
    phase = db.Column(db.String(50), default='')
    created_by_user_id = db.Column(db.Integer, db.ForeignKey('users.user_id'), nullable=True)
    history_id = db.Column(db.Integer, db.ForeignKey('history.history_id'), nullable=False)

class Announcement(db.Model):
    __tablename__ = 'announcements'
    announcement_id = db.Column(db.Integer, primary_key=True)
    message = db.Column(db.Text, nullable=False)
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    attachment_filename = db.Column(db.String(255), nullable=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.user_id'), nullable=False)
    visible_to_manager = db.Column(db.Boolean, default=False, nullable=False)
    visible_to_engineer = db.Column(db.Boolean, default=False, nullable=False)

# ==================== FORMS ==================

class CreateUserForm(FlaskForm):
    username = StringField('Username', validators=[DataRequired(), Length(min=4, max=80)])
    role = SelectField('Role', choices=[('Engineer', 'Engineer'), ('Manager', 'Manager')], validators=[DataRequired()])
    password = PasswordField('Password', validators=[DataRequired(), Length(min=6)])
    confirm_password = PasswordField('Confirm Password', validators=[DataRequired(), EqualTo('password')])
    submit = SubmitField('Create User')
    def validate_username(self, username):
        user = User.query.filter_by(username=username.data).first()
        if user: raise ValidationError('Username taken.')

class UpdateProfileForm(FlaskForm):
    username = StringField('Username', validators=[DataRequired(), Length(min=4, max=80)])
    name = StringField('Full Name', validators=[DataRequired(), Length(min=2, max=100)])
    email = StringField('Email', validators=[DataRequired(), Email()])
    phone_num = StringField('Phone Number', validators=[Optional(), Length(min=6, max=20)])
    submit = SubmitField('Save Changes')
    def validate_username(self, username):
        user = User.query.filter_by(username=username.data).first()
        if user and user.user_id != current_user.user_id: raise ValidationError('Username taken.')
    def validate_email(self, email):
        user = User.query.filter_by(email=email.data).first()
        if user and user.user_id != current_user.user_id: raise ValidationError('Email in use.')

class AnnouncementForm(FlaskForm):
    message = TextAreaField('Message', validators=[DataRequired(), Length(min=1, max=5000)])
    attachment = FileField('Attachment (Optional)', validators=[FileAllowed(['pdf', 'png', 'jpg', 'xlsx'])])
    visible_to_manager = BooleanField('Managers')
    visible_to_engineer = BooleanField('Engineers')
    submit = SubmitField('Post Announcement')
    def validate(self, **kwargs):
        if not super().validate(**kwargs): return False
        if not self.visible_to_manager.data and not self.visible_to_engineer.data:
            self.visible_to_manager.errors.append('Select at least one role.')
            return False
        return True

# ==================== HELPERS ==================

def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        # FIX: Allow BOTH 'Manager' and 'Admin' roles
        if not current_user.is_authenticated or current_user.role.role_name not in ['Manager', 'Admin']:
            flash('Permission denied.', 'error')
            return redirect(url_for('index'))
        return f(*args, **kwargs)
    return decorated_function

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

def get_gemini_prompt():
    return """
    You are an expert engineering assistant. Analyze these technical drawing images.
    Extract the following data points and return them as a clean JSON object.
    1. design_pressure 2. design_temperature 3. operating_pressure 4. operating_temperature 5. fluid
    6. parts_list: list of objects with "part_name", "material_spec", "material_grade".
    Example: {"design_pressure": "14 Bar", "parts_list": [{"part_name": "Shell", "material_spec": "SA-516", "material_grade": "70"}]}
    """

def call_gemini_api(images, prompt, api_key):
    api_url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
    headers = {"Content-Type": "application/json"}
    parts = [{"text": prompt}]
    for img in images:
        img_byte_arr = io.BytesIO()
        img.save(img_byte_arr, format='JPEG')
        img_base64 = base64.b64encode(img_byte_arr.getvalue()).decode('utf-8')
        parts.append({"inline_data": {"mime_type": "image/jpeg", "data": img_base64}})
    
    payload = {"contents": [{"parts": parts}]}
    max_retries = 3
    for attempt in range(max_retries):
        try:
            response = requests.post(api_url, headers=headers, data=json.dumps(payload), timeout=125)
            response.raise_for_status()
            return response.json()['candidates'][0]['content']['parts'][0]['text']
        except Exception as e:
            if attempt == max_retries - 1: raise e
            sleep_time.sleep(1)

def clean_gemini_response(response_text):
    match = re.search(r'```json\s*([\s\S]+?)\s*```', response_text)
    return match.group(1) if match else response_text.strip()

def refine_material_type(spec, grade, ai_suggested_type="Not Found"):
    """
    Determines Material Type by looking at BOTH Spec and Grade.
    Includes rules for Bolts, Nuts, Structural, and JIS standards.
    """
    # 1. Clean up the inputs (Safety first)
    s = str(spec).upper() if spec else ""
    g = str(grade).upper() if grade else ""
    t = str(ai_suggested_type).upper() if ai_suggested_type else ""

    # --- PRIORITY RULES (Specific items first) ---

    # 1. Structural Steel (S275JR)
    if "S275" in s or "S275" in g or "JR" in g:
        return "Structural Steel"

    # 2. Stainless Steel Bolting (SA-193 / A193)
    if "193" in s or "193" in g:
        return "Stainless Steel Bolting"

    # 3. Heavy Hex Nuts (SA-194 / A194)
    if "194" in s or "194" in g:
        return "Heavy Hex Nuts"

    # 4. JIS Carbon Steel (JIS G3507)
    if "G3507" in s or "G3507" in g:
        return "Carbon Steel"

    # --- GENERAL RULES ---

    # 5. Stainless Steel General
    if "304" in g or "316" in g or "321" in g or "347" in g:
        return "Stainless Steel"
    if "SA-240" in s or "SA-312" in s or "SA-182" in s:
        # Double check SA-182 isn't actually Chrome-Moly Alloy
        if "F11" not in g and "F22" not in g:
            return "Stainless Steel"

    # 6. Duplex
    if "2205" in g or "S31803" in g:
        return "Duplex SS"

    # 7. Carbon Steel General
    if "SA-106" in s or "SA-105" in s or "SA-516" in s or "API 5L" in s or "A106" in s or "A516" in s:
        return "Carbon Steel"
    
    # 8. Alloy Steel
    if "F11" in g or "F22" in g or "P11" in g or "P22" in g:
        return "Alloy Steel"

    # 9. Fallback: If we found nothing, trust the AI (unless it said "Other")
    if t and t != "NOT FOUND" and t != "NONE" and t != "OTHER":
        return ai_suggested_type

    return "Not Found"


def parse_gemini_response(json_text, drawing_name):
    extracted_data = []
    try: 
        data = json.loads(json_text)
    except: 
        data = {"parts_list": []}
    
    parts = data.get("parts_list", []) or [{"part_name": "Not Found"}]
    
    # Extract general data once
    fluid = data.get("fluid", "Not Found")
    design_temp = data.get("design_temperature", "Not Found")
    design_press = data.get("design_pressure", "Not Found")
    op_temp = data.get("operating_temperature", "Not Found")
    op_press = data.get("operating_pressure", "Not Found")

    for part in parts:
        # Get raw values
        raw_spec = part.get("material_spec", "Not Found")
        raw_grade = part.get("material_grade", "Not Found")
        
        # --- CRITICAL FIX: CALL THE INFERENCE FUNCTION ---
        final_type = refine_material_type(raw_spec, raw_grade)
        # -------------------------------------------------

        extracted_data.append({
            "source_drawing": drawing_name,
            "part_name": part.get("part_name", "Not Found"),
            "fluid": fluid,
            "material_type": final_type, # Use the calculated type here
            "material_spec": raw_spec,
            "material_grade": raw_grade,
            'TEMP. (°C) ': design_temp,
            'PRESSURE (Mpa) ': design_press,
            'TEMP. (°C)': op_temp,
            'PRESSURE (Mpa)': op_press,
            "insulation": "N/A"
        })
    return extracted_data

# ==================== ROUTES ==================

@app.route('/', methods=['GET', 'POST'])
def login():
    # 1. If already logged in, redirect based on role
    if current_user.is_authenticated:
        if current_user.role.role_name in ['Admin', 'Manager']:
            return redirect(url_for('admin_dashboard'))
        else:
            return redirect(url_for('index'))
    
    # 2. Handle Login Form Submission
    if request.method == 'POST':
        user = User.query.filter((User.username == request.form.get('username')) | (User.email == request.form.get('username'))).first()
        if user and user.check_password(request.form.get('password')) and user.role.role_name.lower() == request.form.get('role').lower():
            login_user(user)
            
            # --- FIX IS HERE: Redirect Admins/Managers to Dashboard, others to Index ---
            if user.role.role_name in ['Admin', 'Manager']:
                return redirect(url_for('admin_dashboard'))
            else:
                return redirect(url_for('index'))
            # ---------------------------------------------------------------------------
            
        flash('Invalid credentials.', 'error')
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))

@app.route('/home')
@login_required
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
@login_required
def upload_file():
    # 1. Basic validation
    if 'drawings' not in request.files:
        flash("No file part", "error")
        return redirect(url_for('index'))
    
    files = request.files.getlist('drawings')
    
    print(f"DEBUG: System received {len(files)} files.") 
    
    if not files or files[0].filename == '':
        flash("No selected file", "error")
        return redirect(url_for('index'))
    
    files.sort(key=lambda x: x.filename)
    all_data = [] 
    
    try:
        # 2. Process each file
        for index, file in enumerate(files):
            if file.filename == '': 
                continue
            
            safe_name = secure_filename(file.filename)
            path = os.path.join(app.config['UPLOAD_FOLDER'], safe_name)
            file.save(path)
            
            print(f"--- Processing File {index + 1}/{len(files)}: {safe_name} ---")

            try:
                # Convert PDF to Image & Call Gemini
                images = pdf2image.convert_from_path(path, poppler_path=app.config.get('POPPLER_PATH') or POPPLER_PATH)
                text = call_gemini_api(images, get_gemini_prompt(), GEMINI_API_KEY)
                clean_text = clean_gemini_response(text)
                
                # Parse and Append
                parsed = parse_gemini_response(clean_text, safe_name)
                
                if parsed:
                    all_data.extend(parsed) 
                    print(f"DEBUG: Successfully extracted {len(parsed)} rows from {safe_name}")
                else:
                    print(f"DEBUG: AI returned empty data for {safe_name}")

                # --- WAIT FOR GOOGLE SERVERS ---
                if index < len(files) - 1:
                    print("Waiting 20 seconds for API cooldown...")
                    sleep_time.sleep(20)  # <--- THIS IS THE FIX
                # -----------------------------

            except Exception as e:
                print(f"CRITICAL ERROR processing file {safe_name}: {e}")
                continue 
        
        # 3. Check if we got any data
        if not all_data:
            flash("No data could be extracted. Check terminal for errors.", "error")
            return redirect(url_for('index'))
            
        # 4. Save
        print(f"DEBUG: Final saving. Total rows collected: {len(all_data)}")
        temp_name = f"temp_{uuid.uuid4().hex}.json"
        with open(os.path.join(app.config['OUTPUT_FOLDER'], temp_name), 'w') as f:
            json.dump(all_data, f)
            
        return redirect(url_for('preview_data', temp_file=temp_name))

    except Exception as e:
        flash(f"System Error: {e}", "error")
        return redirect(url_for('index'))

@app.route('/preview')
@login_required
def preview_page():
    return render_template('preview.html', data_rows=[], equipment_count=0)

@app.route('/preview/<temp_file>')
@login_required
def preview_data(temp_file):
    try:
        with open(os.path.join(app.config['OUTPUT_FOLDER'], temp_file), 'r') as f:
            raw = json.load(f)
        
        preview_rows = []
        equip_count = 0
        row_count = 1
        curr_drawing = ""
        
        for d in raw:
            if d['source_drawing'] != curr_drawing:
                equip_count += 1
                curr_drawing = d['source_drawing']
            
            row = {
                'NO.': equip_count,
                'EQUIPMENT NO. ': f"V-{equip_count:03d}",
                'PMT NO.': os.path.splitext(d['source_drawing'])[0],
                'EQUIPMENT DESCRIPTION': "",
                'PARTS': d.get('part_name'),
                'PHASE': "",
                'FLUID': d.get('fluid'),
                'TYPE': d.get('material_type'),
                'SPEC.': d.get('material_spec'),
                'GRADE': d.get('material_grade'),
                'INSULATION': d.get('insulation'),
                'TEMP. (°C) ': d.get('TEMP. (°C) '),
                'PRESSURE (Mpa) ': d.get('PRESSURE (Mpa) '),
                'TEMP. (°C)': d.get('TEMP. (°C)'),
                'PRESSURE (Mpa)': d.get('PRESSURE (Mpa)'),
                'source_drawing': d.get('source_drawing')
            }
            preview_rows.append(row)
            row_count += 1
            
        return render_template('preview.html', data_rows=preview_rows, equipment_count=equip_count, temp_file=temp_file)
    except Exception as e:
        flash(f"Error loading preview: {e}", "error")
        return redirect(url_for('index'))

@app.route('/view_uploaded_file/<filename>')
@login_required
def view_uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], secure_filename(filename), as_attachment=False)

@app.route('/manual-input')
@login_required
def manual_input():
    return render_template('manual_input.html')

# --- THIS IS THE UPDATED SAVE FUNCTION WITH EXCEL MERGING ---
@app.route('/save_data', methods=['POST'])
@login_required
def save_data():
    try:
        def get_vals(name): return [x.strip() for x in request.form.getlist(name)]
        
        # 1. Collect Data from Form
        rows = []
        parts = get_vals('PARTS') # Use parts as length reference
        
        for i in range(len(parts)):
            rows.append({
                'NO.': get_vals('NO.')[i],
                'EQUIPMENT NO. ': get_vals('EQUIPMENT NO. ')[i],
                'PMT NO.': get_vals('PMT NO.')[i],
                'EQUIPMENT DESCRIPTION': get_vals('EQUIPMENT DESCRIPTION')[i],
                'PARTS': parts[i],
                'PHASE': get_vals('PHASE')[i],
                'FLUID': get_vals('FLUID')[i],
                'TYPE': get_vals('TYPE')[i],
                'SPEC.': get_vals('SPEC.')[i],
                'GRADE': get_vals('GRADE')[i],
                'INSULATION': get_vals('INSULATION')[i],
                'TEMP. (°C) ': get_vals('TEMP. (°C) ')[i],
                'PRESSURE (Mpa) ': get_vals('PRESSURE (Mpa) ')[i],
                'TEMP. (°C)': get_vals('TEMP. (°C)')[i],
                'PRESSURE (Mpa)': get_vals('PRESSURE (Mpa)')[i],
                'source_drawing': request.form.getlist('source_drawing')[i]
            })

        # 2. Save to Database (Full Data)
        history = History(created_by_user_id=current_user.user_id)
        db.session.add(history)
        db.session.flush()
        
        for r in rows:
            db.session.add(EquipmentData(
                history_id=history.history_id,
                created_by_user_id=current_user.user_id,
                source_drawing=r['source_drawing'],
                part_name=r['PARTS'],
                fluid=r['FLUID'],
                material_type=r['TYPE'],
                material_spec=r['SPEC.'],
                material_grade=r['GRADE'],
                design_temp=r['TEMP. (°C) '],
                design_pressure=r['PRESSURE (Mpa) '],
                operating_temp=r['TEMP. (°C)'],
                operating_pressure=r['PRESSURE (Mpa)'],
                insulation=r['INSULATION'],
                no=r['NO.'],
                equipment_no=r['EQUIPMENT NO. '],
                pmt_no=r['PMT NO.'],
                equipment_description=r['EQUIPMENT DESCRIPTION'],
                phase=r['PHASE']
            ))

        # 3. Prepare Data for Excel (Blanking duplicates for merging)
        excel_rows = []
        prev_equip_no = None
        
        for r in rows:
            row_copy = r.copy()
            curr_equip = r['EQUIPMENT NO. ']
            
            # Logic: If this Equipment No is the same as previous, 
            # make the grouping columns empty so the Merge logic works.
            if curr_equip == prev_equip_no:
                row_copy['NO.'] = ""
                row_copy['EQUIPMENT NO. '] = ""
                row_copy['PMT NO.'] = ""
                row_copy['EQUIPMENT DESCRIPTION'] = ""
                row_copy['PHASE'] = ""
            else:
                prev_equip_no = curr_equip
            
            excel_rows.append(row_copy)

        # 4. Write to Excel
        filename = f"{history.history_id}_output.xlsx"
        filepath = os.path.join(app.config['OUTPUT_FOLDER'], filename)
        
        shutil.copyfile(TEMPLATE_FILE, filepath)
        book = load_workbook(filepath)
        start_row = book[TARGET_SHEET].max_row
        book.close()
        
        df = pd.DataFrame(excel_rows).reindex(columns=template_column_order)
        
        with pd.ExcelWriter(filepath, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            df.to_excel(writer, sheet_name=TARGET_SHEET, startrow=start_row, index=False, header=False)

        # 5. Apply Merging Logic (Restored from your old code)
        print("Applying cell merging...")
        wb = load_workbook(filepath)
        ws = wb[TARGET_SHEET]
        
        # Columns to merge: A(1), B(2), C(3), D(4)
        cols_to_merge = [1, 2, 3, 4]
        
        # Start looking from where we just wrote data
        # Note: If start_row was 1 (header), data starts at 2. 
        # But allow scanning whole sheet to be safe or start from newly added area.
        row = 2 
        max_row = ws.max_row

        while row <= max_row:
            # Check Equipment No (Column 2)
            cell_val = ws.cell(row=row, column=2).value
            
            if cell_val:
                start_merge_row = row
                next_row = row + 1
                
                # Look ahead for empty cells in Col 2 (Created in Step 3)
                while next_row <= max_row and not ws.cell(row=next_row, column=2).value:
                    next_row += 1
                
                end_merge_row = next_row - 1
                
                # Merge logic
                if end_merge_row > start_merge_row:
                    for col_idx in cols_to_merge:
                        ws.merge_cells(start_row=start_merge_row, start_column=col_idx, end_row=end_merge_row, end_column=col_idx)
                        cell = ws.cell(row=start_merge_row, column=col_idx)
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                else:
                    # Center single row items too
                    for col_idx in cols_to_merge:
                        cell = ws.cell(row=start_merge_row, column=col_idx)
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                        
                row = next_row
            else:
                row += 1
        
        wb.save(filepath)
        print("Merging complete.")

        # 6. Finalize
        history.excel_filename = filename
        db.session.commit()
        
        if request.form.get('temp_file'):
            try: os.remove(os.path.join(app.config['OUTPUT_FOLDER'], request.form.get('temp_file')))
            except: pass
            
        flash('Data saved successfully!', 'success')
        return render_template('preview.html', data_rows=rows, equipment_count=len(set(get_vals('EQUIPMENT NO. '))), excel_file=filename)

    except Exception as e:
        db.session.rollback()
        print(f"Error saving: {e}")
        flash(f"Error saving: {e}", "error")
        return redirect(url_for('manual_input'))

@app.route('/download/<filename>')
@login_required
def download_file(filename):
    return send_from_directory(app.config['OUTPUT_FOLDER'], filename, as_attachment=True)

@app.route('/notifications')
@login_required
def notifications():
    anns = Announcement.query.filter_by(visible_to_engineer=True).order_by(Announcement.created_at.desc()).all()
    return render_template('notification.html', announcements=anns)

@app.route('/profile', methods=['GET', 'POST'])
@login_required
def personal_info():
    form = UpdateProfileForm()
    if form.validate_on_submit():
        current_user.username = form.username.data
        current_user.name = form.name.data
        current_user.email = form.email.data
        current_user.phone_num = form.phone_num.data
        db.session.commit()
        flash('Profile updated.', 'success')
        return redirect(url_for('personal_info'))
    form.username.data = current_user.username
    form.name.data = current_user.name
    form.email.data = current_user.email
    form.phone_num.data = current_user.phone_num
    return render_template('personal_info.html', form=form)

@app.route('/download/announcement/<filename>')
@login_required
def download_announcement(filename):
    return send_from_directory(app.config['ANNOUNCEMENT_FOLDER'], filename, as_attachment=True)

# ==================== ADMIN ROUTES ==================

@app.route('/admin/dashboard')
@login_required
@admin_required
def admin_dashboard():
    u_count = User.query.count()
    a_count = User.query.join(Role).filter(Role.role_name == 'Manager').count()
    e_count = User.query.join(Role).filter(Role.role_name == 'Engineer').count()
    f_total = History.query.count()
    today = datetime.utcnow().date()
    f_today = History.query.filter(History.created_at >= datetime.combine(today, time.min), History.created_at <= datetime.combine(today, time.max)).count()
    recent = User.query.order_by(User.user_id.desc()).limit(5).all()
    return render_template('dashboard_admin.html', user_count=u_count, admin_count=a_count, engineer_count=e_count, file_count_total=f_total, file_count_today=f_today, recent_users=recent)

@app.route('/admin/create-user', methods=['GET', 'POST'])
@login_required
@admin_required
def admin_create_user():
    form = CreateUserForm()
    if form.validate_on_submit():
        role = Role.query.filter_by(role_name=form.role.data).first()
        u = User(username=form.username.data, name=form.username.data, email=f"{form.username.data}@ipetro.com", role_id=role.role_id)
        u.set_password(form.password.data)
        db.session.add(u)
        db.session.commit()
        flash('User created.', 'success')
        return redirect(url_for('admin_create_user'))
    return render_template('createuser.html', form=form)

@app.route('/admin/announcement', methods=['GET', 'POST'])
@login_required
@admin_required
def admin_announcement():
    form = AnnouncementForm()
    if form.validate_on_submit():
        fname = None
        if form.attachment.data:
            fname = secure_filename(form.attachment.data.filename)
            form.attachment.data.save(os.path.join(app.config['ANNOUNCEMENT_FOLDER'], fname))
        db.session.add(Announcement(message=form.message.data, attachment_filename=fname, user_id=current_user.user_id, visible_to_manager=form.visible_to_manager.data, visible_to_engineer=form.visible_to_engineer.data))
        db.session.commit()
        flash('Announcement posted.', 'success')
        return redirect(url_for('admin_announcement'))
    return render_template('announcement.html', form=form, announcements=Announcement.query.order_by(Announcement.created_at.desc()).all())

@app.route('/admin/history')
@login_required
@admin_required
def admin_history():
    # Fetch all history records, newest first
    items = History.query.order_by(History.created_at.desc()).all()
    return render_template('history_admin.html', histories=items)

@app.route('/admin/reports')
@login_required
@admin_required
def admin_reports():
    # Placeholder for future Reports functionality
    flash("The Reports module is currently under construction.", "info")
    return redirect(url_for('admin_dashboard'))

@app.route('/admin/review-queue')
@login_required
@admin_required
def admin_review_queue():
    # Placeholder for future Review Queue functionality
    flash("The Review Queue module is currently under construction.", "info")
    return redirect(url_for('admin_dashboard'))

@app.route('/admin/statistics')
@login_required
@admin_required
def admin_statistics():
    # This page is not ready yet, so we redirect back to the dashboard
    flash("The Statistics page is currently under construction.", "info")
    return redirect(url_for('admin_dashboard'))

# --- CLI COMMANDS ---
@app.cli.command('init-db')
@click.option('--drop', is_flag=True)
def init_db(drop):
    if drop: db.drop_all()
    db.create_all()
    if not Role.query.first():
        # 1. Create Roles
        m = Role(role_name='Manager')
        e = Role(role_name='Engineer')
        a = Role(role_name="Admin")
        db.session.add_all([m, e, a])
        db.session.commit()
        
        # 2. Create Users (With CORRECT variable names)
        
        # Manager
        manager_user = User(username='manager@ipetro.com', name='Manager', email='manager@ipetro.com', role_id=m.role_id)
        manager_user.set_password('abc1234')
        
        # Engineer
        eng_user = User(username='engineer@ipetro.com', name='Eng', email='engineer@ipetro.com', role_id=e.role_id)
        eng_user.set_password('abc1234')
        
        # Admin
        admin_user = User(username='admin@ipetro.com', name='Admin', email='Admin@ipetro.com', role_id=a.role_id)
        admin_user.set_password('abc1234')
        
        # 3. Add all unique variables to session
        db.session.add_all([manager_user, eng_user, admin_user])
        db.session.commit()
        
    print("DB Initialized.")

if __name__ == '__main__':
    app.run(debug=True)
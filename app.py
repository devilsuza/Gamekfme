import base64
import io
import json
import math
import os
import sys
import re
import unicodedata
import uuid
from datetime import datetime, time, timedelta
from pathlib import Path
from flask import Flask, render_template_string, request, Response, send_file, session, redirect, url_for, flash, send_from_directory, jsonify
from github import Github
import openpyxl
import qrcode
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from werkzeug.utils import secure_filename
from zk import ZK, const
from pypdf import PdfReader, PdfWriter

app = Flask(__name__)
APP_BUILD = '2026-10-10-POPUP-SERVER-FIX-V7'
app.secret_key = os.getenv('SECRET_KEY', 'gamek_fresmart_secret_key_sonu')

UPLOAD_FOLDER = os.getenv('UPLOAD_FOLDER', 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = int(os.getenv('MAX_UPLOAD_MB','50')) * 1024 * 1024
ALLOWED_EXTENSIONS = {'pdf', 'png', 'jpg', 'jpeg', 'doc', 'docx'}

# Machine Configurations with Multiple Stores
MACHINES = {
    'LM11': {'ip': os.getenv('MACHINE_IP', '192.168.1.153'), 'port': int(os.getenv('MACHINE_PORT', 4370)), 'name': 'Gamek HRMS'},
    'LF07': {'ip': '192.168.88.16', 'port': 4370, 'name': 'Benfica LF07 HRMS'},
    'DEV': {'ip': '127.0.0.1', 'port': 4370, 'name': 'Fresmart'}
}

GITHUB_TOKEN = os.getenv('GITHUB_TOKEN', '')
GITHUB_REPO_NAME = os.getenv('GITHUB_REPO_NAME', 'devilsuza/Gamekfme')
GITHUB_BRANCH = os.getenv('GITHUB_BRANCH', 'main')

# Persistent Storage Files
LEAVE_JSON_FILE = 'leave_requests.json'
LEAVE_EXCEL_FILE = 'leave_records.xlsx'
LEAVE_CODE_NAMES = {
    'F01;1': 'Baixa Médica',
    'F03;1': 'Falta Injustificada',
    'F05;1': 'Licença sem vencimento',
    'F10;1': 'Falta Justificada',
    'F51;1': 'Casamento',
    'F60;1': 'Nascimento',
    'F61;1': 'Obito',
    'F62;1': 'Gravidez',
}

ATTENDANCE_OVERRIDES_FILE = 'attendance_overrides.json'
MANUAL_PUNCHES_FILE = 'manual_punches.json'
DUPLICATE_PUNCHES_FILE = 'duplicate_punches.json'
DUPLICATE_PUNCH_WINDOW_SECONDS = 60
ROSTER_JSON_FILE = 'roster.json'
CORRECTIONS_FILE = 'attendance_corrections.json'
LEAVE_BALANCES_FILE = 'leave_balances.json'
DOCUMENTS_FILE = 'employee_documents.json'
PROMOTION_DOCUMENTS_FILE = 'promotion_documents.json'
PROMOTION_PERIODS_FILE = 'promotion_periods.json'
KVI_POWER_DOCUMENTS_FILE = 'kvi_power_documents.json'
KVI_POWER_DATA_FILE = 'kvi_power_data.json'
BUSINESS_DOC_FOLDER = os.path.join(UPLOAD_FOLDER, 'business_documents')
os.makedirs(BUSINESS_DOC_FOLDER, exist_ok=True)

SYNC_HISTORY_FILE = 'sync_history.json'
STAFFING_RULES_FILE = 'staffing_rules.json'
NOTIFICATIONS_FILE = 'notifications.json'
OVERTIME_APPROVALS_FILE = 'overtime_approvals.json'
OVERTIME_MIN_MINUTES = int(os.getenv('OVERTIME_MIN_MINUTES', '15'))

AUDIT_LOG_FILE = 'audit_log.json'
EMPLOYEE_DOC_FOLDER = os.path.join(UPLOAD_FOLDER, 'employee_documents')
os.makedirs(EMPLOYEE_DOC_FOLDER, exist_ok=True)
 
SHIFT_REQUESTS_FILE = 'shift_requests.json' 
PASSWORD_RESETS_FILE = 'password_resets.json'
USERS_DB_FILE = 'users_db.json'
SALARY_SLIPS_FILE = 'salary_slips.json'
MACHINES_DB_FILE = 'machines_db.json'
PERMISSIONS = ['dashboard','attendance_view','attendance_edit','employee_id_cards','employee_manage','roster','shift_approvals','overtime_approval','overtime_history','password_management','calendar','leave_management','payroll','employee_documents','attendance_corrections','notifications','workforce_hub','all_stores_attendance','biometric_sync','leave_balances','roster_copy','staffing_rules','audit_log','qr_directory','user_management','machine_management','profile']
PERMISSION_LABELS = {
 'dashboard':'Dashboard','attendance_view':'Attendance View','attendance_edit':'Edit Attendance Times','employee_id_cards':'Employee ID Cards','employee_manage':'Add / Manage Employees','roster':'Roster Planner','shift_approvals':'Shift Approvals','overtime_approval':'Overtime Approval','overtime_history':'Overtime History','password_management':'Password Management','calendar':'Calendar & Rota','leave_management':'Leave Management','payroll':'Payroll & Reports','employee_documents':'Employee Documents','attendance_corrections':'Attendance Corrections','notifications':'Notifications','workforce_hub':'Workforce Automation Hub','all_stores_attendance':'All Stores Attendance','biometric_sync':'Biometric Sync History','leave_balances':'Leave Balances','roster_copy':'Roster Copy','staffing_rules':'Staffing Rules','audit_log':'Audit Log','qr_directory':'Employee QR Directory','user_management':'User Management','machine_management':'Machine Management','profile':'My Profile'
}

# FIX: Added 'payroll' permission to EMPLOYEE role
ROLE_PRESETS = {
    'HR':['dashboard','attendance_view','leave_management','payroll','employee_id_cards','employee_documents','notifications','profile'],
    'AREA MANAGER':['dashboard','attendance_view','attendance_edit','employee_id_cards','employee_manage','leave_management','roster','shift_approvals','overtime_approval','overtime_history','calendar','notifications','profile'],
    'OPERATION HEAD':['dashboard','attendance_view','attendance_edit','employee_id_cards','employee_manage','leave_management','roster','shift_approvals','overtime_approval','overtime_history','payroll','calendar','employee_documents','attendance_corrections','notifications','workforce_hub','profile'],
    'STORE MANAGER':['dashboard','attendance_view','attendance_edit','employee_id_cards','employee_manage','leave_management','roster','shift_approvals','overtime_approval','overtime_history','calendar','notifications','profile'],
    'EMPLOYEE':['dashboard','attendance_view','calendar','payroll','employee_documents','attendance_corrections','notifications','profile']
}
DEFAULT_MACHINES = dict(MACHINES)

SYNCED_ATTENDANCE_LOGS = []
# Dictionary based device sync tracker for multiple stores
LAST_DEVICE_SYNC_TIME = {}

# Master Employees (Fallback / Default Setup)
EMPLOYEE_IDENTITY_DATA = {'NWC8364': {'data_de_contrato': '07/15/26', 'identificacao': '0008858606UE045'}, 'NWC2652': {'data_de_contrato': '04/02/21', 'identificacao': '0007320697UE040'}, 'NWC4554': {'data_de_contrato': '07/27/23', 'identificacao': '0007010800KS048'}, 'NWC3381': {'data_de_contrato': '10/07/22', 'identificacao': '001491835UE035'}, 'NWC2788': {'data_de_contrato': '06/27/22', 'identificacao': '0003316212BA038'}, 'NWC1010': {'data_de_contrato': '10/10/18', 'identificacao': '0003726506UE038'}, 'NWC1983': {'data_de_contrato': '06/27/22', 'identificacao': '000841297LA033'}, 'NWC5187': {'data_de_contrato': '06/04/24', 'identificacao': '0006083083LA043'}, 'NWC1168': {'data_de_contrato': '03/19/19', 'identificacao': '0006278742BA042'}, 'NWC1525': {'data_de_contrato': '06/24/21', 'identificacao': '0005957722LA049'}, 'NWC1553': {'data_de_contrato': '07/01/21', 'identificacao': '0002518985LA039'}, 'NWC3596': {'data_de_contrato': '10/14/22', 'identificacao': '002869217LA033'}, 'NWC3127': {'data_de_contrato': '08/26/22', 'identificacao': '0009463847LA048'}, 'NWC2005': {'data_de_contrato': '09/01/21', 'identificacao': '0009206303LA041'}, 'NWC5168': {'data_de_contrato': '06/04/24', 'identificacao': '0003096236LA038'}, 'NWC5713': {'data_de_contrato': '11/14/24', 'identificacao': '0001358868LA036'}, 'NWC5186': {'data_de_contrato': '06/01/24', 'identificacao': '0001144465LA011'}, 'NWC3318': {'data_de_contrato': '09/20/22', 'identificacao': '0009972719LA042'}, 'NWC2300': {'data_de_contrato': '03/15/22', 'identificacao': '0003367879LA034'}, 'NWC5830': {'data_de_contrato': '01/17/25', 'identificacao': '0005999061LA043'}, 'NWC5529': {'data_de_contrato': '08/19/24', 'identificacao': '0002286363LA034'}, 'NWC2624': {'data_de_contrato': '04/18/22', 'identificacao': '007833178LA049'}, 'NWC5396': {'data_de_contrato': '06/20/24', 'identificacao': '000419510KN032'}, 'NWC3711': {'data_de_contrato': '10/28/22', 'identificacao': '000137207LA032'}, 'NWC2757': {'data_de_contrato': '05/16/22', 'identificacao': '0004945742LN043'}, 'NWC3791': {'data_de_contrato': '11/17/22', 'identificacao': '0005968810CA047'}, 'NWC4281': {'data_de_contrato': '03/10/23', 'identificacao': '000955344LA030'}, 'NWC4481': {'data_de_contrato': '05/31/23', 'identificacao': '0004858352LA044'}, 'NWC5222': {'data_de_contrato': '06/20/24', 'identificacao': '0004785268LA046'}, 'NWC2981': {'data_de_contrato': '07/13/22', 'identificacao': '0005280299LA040'}, 'NWC6661': {'data_de_contrato': '08/15/25', 'identificacao': '0005518818BA046'}, 'NWC6444': {'data_de_contrato': '06/20/25', 'identificacao': '0003579623LA037'}, 'NWC6638': {'data_de_contrato': '08/16/25', 'identificacao': '0007074038KS044'}, 'NWC6702': {'data_de_contrato': '08/18/25', 'identificacao': '0009148747LA047'}, 'NWC8362': {'data_de_contrato': '07/15/26', 'identificacao': '0006864584UE041'}, 'NWC8328': {'data_de_contrato': '07/10/26', 'identificacao': '006246050LA048'}, 'NWC8350': {'data_de_contrato': '07/15/26', 'identificacao': '0002789758LA031'}, 'NWC6715': {'data_de_contrato': '08/18/25', 'identificacao': '006753746LA046'}, 'NWC7347': {'data_de_contrato': '12/19/25', 'identificacao': '007247013HO044'}, 'NWC8361': {'data_de_contrato': '07/15/26', 'identificacao': '0007046482ME043'}}

MASTER_EMPLOYEES = {
    'NWC2981': {'name': 'ANTONIO JOSE BANDOLA', 'off': 'SUNDAY', 'dept': 'ADMIN - MANAGER', 'shift': 'morning'},
    'NWC3127': {'name': 'MATEUS ANTONIO DA COSTA BALMIRO', 'off': 'FRIDAY', 'dept': 'ADMIN - MANAGER', 'shift': 'morning'},
    'NWC1525': {'name': 'ETY JOSÉ BANDUA MONTEIRO', 'off': 'MONDAY', 'dept': 'ADMIN - MANAGER', 'shift': 'morning'},
    'NWC8328': {'name': 'JOAO MATIAS DOMINGOS', 'off': 'SATURDAY', 'dept': 'ADMIN - CCTV', 'shift': 'morning'},
    'NWC6661': {'name': 'FRANCISCO MUNDELE CHIVELA', 'off': 'WEDNESDAY', 'dept': 'ADMIN - CCTV', 'shift': 'morning'},
    'NWC1553': {'name': 'TIAGO SANDALA CHISSANHA', 'off': 'SUNDAY', 'dept': 'ADMIN - AUDITOR', 'shift': 'morning'},
    'NWC8350': {'name': 'LOLIVALDO ALBERTO MADEIRA', 'off': 'SUNDAY', 'dept': 'ADMIN - EDP', 'shift': 'morning'},
    'NWC5187': {'name': 'VICTOR NSOSI JOAO', 'off': 'MONDAY', 'dept': 'CASH - HEAD', 'shift': 'morning'},
    'NWC1168': {'name': 'ADELIA MBALOMBO CHIPEPI', 'off': 'SUNDAY', 'dept': 'CASH - HEAD', 'shift': 'morning'},
    'NWC2652': {'name': 'DULCE DOROTEIA GARCIA LUSITANO', 'off': 'MONDAY', 'dept': 'CASH - HEAD', 'shift': 'morning'},
    'NWC3381': {'name': 'ANDRE DE JESUS NGOLA JOSE', 'off': 'TUESDAY', 'dept': 'CASH', 'shift': 'morning'},
    'NWC1983': {'name': 'PATRICIA SOLANGE FRANCISCO', 'off': 'THURSDAY', 'dept': 'CASH', 'shift': 'second'},
    'NWC8364': {'name': 'DIELUMBAKA AUGUSTO', 'off': 'WEDNESDAY', 'dept': 'CASH', 'shift': 'second'},
    'NWC2788': {'name': 'INES NACHINGOLO FELICIANO NAMBELO', 'off': 'TUESDAY', 'dept': 'CASH', 'shift': 'morning'},
    'NWC1010': {'name': 'TERESA PEDRO LEAO', 'off': 'FRIDAY', 'dept': 'CASH', 'shift': 'morning'},
    'NWC6638': {'name': 'CLAUDIO JANUARIO MANUEL AVELINO', 'off': 'SUNDAY', 'dept': 'TALHO', 'shift': 'morning'},
    'NWC5830': {'name': 'REGINA DE FATIMA VIDAL', 'off': 'MONDAY', 'dept': 'TALHO', 'shift': 'morning'},
    'NWC5529': {'name': 'ALEXANDRE LUIS CORREIA', 'off': 'FRIDAY', 'dept': 'TALHO', 'shift': 'morning'},
    'NWC5713': {'name': 'ROSA GARNEIRA BUMBA', 'off': 'WEDNESDAY', 'dept': 'TALHO', 'shift': 'morning'},
    'NWC5396': {'name': 'COSTA BEBIANO HEBO', 'off': 'THURSDAY', 'dept': 'TALHO', 'shift': 'morning'},
    'NWC8361': {'name': 'AGOSTINHO JOAQUIM KUANGO DA COSTA', 'off': 'SUNDAY', 'dept': 'SECU', 'shift': 'morning'},
    'NWC2300': {'name': 'JOANA CARDOSO JOAQUIM AFONSO', 'off': 'MONDAY', 'dept': 'SECU', 'shift': 'morning'},
    'NWC5168': {'name': 'JOAO NVUNDA DALA', 'off': 'THURSDAY', 'dept': 'F & V', 'shift': 'morning'},
    'NWC3711': {'name': 'ANGELA MARIA BUMBA', 'off': 'FRIDAY', 'dept': 'F & V', 'shift': 'morning'},
    'NWC5186': {'name': 'HELIA DOMINGOS DE CARVALHO', 'off': 'WEDNESDAY', 'dept': 'SECU', 'shift': 'morning'},
    'NWC6702': {'name': 'DOMINGOS GAMA PEREIRA', 'off': 'FRIDAY', 'dept': 'SECU', 'shift': 'morning'},
    'NWC3596': {'name': 'ALDAIR FERNANDES FERREIRA', 'off': 'TUESDAY', 'dept': 'SECU', 'shift': 'morning'},
    'NWC2757': {'name': 'JOSEFA KUELUNGA MUASSOKA', 'off': 'THURSDAY', 'dept': 'SECU', 'shift': 'morning'},
    'NWC4554': {'name': 'DOMINGOS ANTONIO FERNANDO', 'off': 'THURSDAY', 'dept': 'CASH', 'shift': 'morning'},
    'NWC2624': {'name': 'CECILIA JORGE FAMOSO', 'off': 'FRIDAY', 'dept': 'CASH', 'shift': 'morning'},
    'NWC3318': {'name': 'JOSE MANUEL KAZOLA', 'off': 'SUNDAY', 'dept': 'FRESCO', 'shift': 'morning'},
    'NWC7347': {'name': 'ARMANDO CHICOVO SAMBA', 'off': 'TUESDAY', 'dept': 'FRESCO', 'shift': 'morning'},
    'NWC8362': {'name': 'ARAUJO PAULOMENDES', 'off': 'FRIDAY', 'dept': 'STOCK', 'shift': 'morning'},
    'NWC6715': {'name': 'RIBEIRO ANTONIO FRANCISCO', 'off': 'THURSDAY', 'dept': 'STOCK', 'shift': 'morning'},
    'NWC6444': {'name': 'HENRIQUES BRANDAO', 'off': 'WEDNESDAY', 'dept': 'STOCK', 'shift': 'morning'}
}

EMPLOYEE_OVERRIDES = {
    '8364': {'code': 'NWC8364', 'name': 'DIELUMBAKA AUGUSTO'},
    '1': {'code': 'NWC8350', 'name': 'LOLIVALDO ALBERTO MADEIRA'},
    '8362': {'code': 'NWC8362', 'name': 'ARAUJO PAULOMENDES'},
    '6661': {'code': 'NWC6661', 'name': 'FRANCISCO MUNDELE CHIVELA'}
}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def load_json_file(filepath):
    if os.path.exists(filepath):
        try:
            with open(filepath, 'r', encoding='utf-8') as f: return json.load(f)
        except Exception: pass
    return [] if 'list' in filepath or 'requests' in filepath or 'resets' in filepath or 'slips' in filepath else {}

def save_json_file(filepath, data):
    try:
        with open(filepath, 'w', encoding='utf-8') as f: json.dump(data, f, indent=4, ensure_ascii=False)
    except Exception as e: print(f"Error saving {filepath}: {e}")

def load_machines_db():
    global MACHINES
    stored = load_json_file(MACHINES_DB_FILE)
    if not isinstance(stored, dict) or not stored:
        stored = {k: dict(v) for k, v in MACHINES.items()}
        save_json_file(MACHINES_DB_FILE, stored)
    changed = False
    for code, info in DEFAULT_MACHINES.items():
        if code not in stored:
            stored[code] = dict(info); changed = True
    if changed: save_json_file(MACHINES_DB_FILE, stored)
    MACHINES = stored
    return MACHINES

load_machines_db()

def normalize_user_record(info):
    info = dict(info or {})
    info.setdefault('status','active'); info.setdefault('role','employee')
    info.setdefault('designation', info.get('dept','Employee'))
    info.setdefault('job_role', 'EMPLOYEE' if info.get('role') == 'employee' else 'ADMIN')
    stores = info.get('stores')
    if not isinstance(stores,list): stores=[info.get('store','LM11')]
    info['stores']=[str(x).upper() for x in stores if x]
    if not info['stores']: info['stores']=['LM11']
    info['store']=info['stores'][0]
    perms=info.get('permissions')
    if not isinstance(perms,list):
        preset=ROLE_PRESETS.get(str(info.get('job_role','')).upper())
        perms=(PERMISSIONS[:] if info.get('role') in ['admin','developer'] else (preset[:] if preset else ROLE_PRESETS['EMPLOYEE'][:]))
    info['permissions']=[p for p in perms if p in PERMISSIONS]
    info.setdefault('profile_photo',''); info.setdefault('off','SUNDAY'); info.setdefault('dept',info.get('designation','General')); info.setdefault('shift','morning')
    statutory=EMPLOYEE_IDENTITY_DATA.get(str(info.get('emp_code','')).upper(),{})
    info.setdefault('identificacao', statutory.get('identificacao',''))
    info.setdefault('data_de_contrato', statutory.get('data_de_contrato',''))
    return info

def user_has_permission(user_id, permission):
    if session.get('role')=='developer' or user_id=='NCSA0608': return True
    info=normalize_user_record(load_users_db().get(user_id,{}))
    return permission in info.get('permissions',[])

def session_has_permission(permission): return user_has_permission(session.get('user_id'), permission)

def get_user_stores(user_id=None):
    user_id=user_id or session.get('user_id')
    if session.get('role')=='developer' or user_id=='NCSA0608':
        return [x for x in MACHINES.keys() if x != 'DEV']
    info=normalize_user_record(load_users_db().get(user_id,{}))
    primary=(info.get('store') or session.get('store') or '').upper()
    if primary in MACHINES and primary != 'DEV':
        return [primary]
    valid=[x for x in info.get('stores',[]) if x in MACHINES and x != 'DEV']
    return valid[:1]

def save_leave_requests(data): save_json_file(LEAVE_JSON_FILE,data)

def save_leave_to_excel(user_id,name,start_date,end_date,leave_type,filename):
    try:
        if os.path.exists(LEAVE_EXCEL_FILE):
            try: wb=openpyxl.load_workbook(LEAVE_EXCEL_FILE); ws=wb.active
            except Exception: wb=openpyxl.Workbook(); ws=wb.active
        else: wb=openpyxl.Workbook(); ws=wb.active
        if ws.max_row==1 and ws.cell(1,1).value is None:
            ws.append(['User ID','Name','Start Date','End Date','Leave Type','Document','Saved At'])
        elif ws.max_row==1 and ws.cell(1,1).value != 'User ID':
            ws.insert_rows(1); ws.append(['User ID','Name','Start Date','End Date','Leave Type','Document','Saved At'])
        ws.append([user_id,name,start_date,end_date,leave_type,filename,datetime.now().strftime('%Y-%m-%d %H:%M:%S')]); wb.save(LEAVE_EXCEL_FILE)
    except Exception as e: print(f'Leave Excel save error: {e}')

# Database initialization
def load_users_db():
    db=load_json_file(USERS_DB_FILE)
    if not db:
        db={}
        for k,v in MASTER_EMPLOYEES.items():
            db[k]=normalize_user_record({'name':v['name'],'password':'123','store':'LM11','status':'active','role':'employee','off':v.get('off','SUNDAY'),'dept':v.get('dept','General'),'shift':v.get('shift','morning'),'designation':v.get('dept','Employee'),'job_role':'EMPLOYEE'})
        db['LM11']=normalize_user_record({'name':'Admin (LM11)','password':os.getenv('ADMIN_PWD','Gamek@789'),'store':'LM11','status':'active','role':'admin','off':'SUNDAY','dept':'Admin','shift':'morning','designation':'Store Admin','job_role':'ADMIN'})
        db['LF07']=normalize_user_record({'name':'Admin (LF07)','password':'123','store':'LF07','status':'active','role':'admin','off':'SUNDAY','dept':'Admin','shift':'morning','designation':'Store Admin','job_role':'ADMIN'})
        save_json_file(USERS_DB_FILE,db)
    changed=False
    for uid in list(db.keys()):
        n=normalize_user_record(db[uid])
        if n!=db[uid]: db[uid]=n; changed=True
    if changed: save_json_file(USERS_DB_FILE,db)
    return db

LEAVE_REQUESTS = load_json_file(LEAVE_JSON_FILE)
def load_overrides(): return load_json_file(ATTENDANCE_OVERRIDES_FILE)
def save_overrides(data): save_json_file(ATTENDANCE_OVERRIDES_FILE, data)
def load_manual_punches(): return load_json_file(MANUAL_PUNCHES_FILE) if isinstance(load_json_file(MANUAL_PUNCHES_FILE), list) else []
def load_roster(): return load_json_file(ROSTER_JSON_FILE)
def save_roster(data): save_json_file(ROSTER_JSON_FILE, data)
def load_shift_requests(): return load_json_file(SHIFT_REQUESTS_FILE) if isinstance(load_json_file(SHIFT_REQUESTS_FILE), list) else []
def save_shift_requests(data): save_json_file(SHIFT_REQUESTS_FILE, data)
def load_salary_slips(): return load_json_file(SALARY_SLIPS_FILE) if isinstance(load_json_file(SALARY_SLIPS_FILE), list) else []
def save_salary_slips(data): save_json_file(SALARY_SLIPS_FILE, data)

def load_overtime_approvals():
    data=load_json_file(OVERTIME_APPROVALS_FILE)
    return data if isinstance(data,list) else []

def save_overtime_approvals(data):
    save_json_file(OVERTIME_APPROVALS_FILE,data)

def overtime_record(user_id,date_str):
    return next((x for x in load_overtime_approvals() if x.get('user_id')==user_id and x.get('date')==date_str),None)

def ensure_overtime_request(user_id,name,store,date_str,minutes,actual_seconds,target_seconds):
    rows=load_overtime_approvals()
    rec=next((x for x in rows if x.get('user_id')==user_id and x.get('date')==date_str),None)
    now=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    if rec:
        # Do not reopen a rejected item until punches are edited to a different overtime amount.
        old_minutes=int(rec.get('minutes',0))
        if rec.get('status')=='Rejected' and old_minutes==minutes:return rec
        if rec.get('status')=='Rejected' and old_minutes!=minutes:
            rec.update({'status':'Pending','minutes':minutes,'actual_seconds':actual_seconds,'target_seconds':target_seconds,'updated_at':now,'reviewed_by':'','reviewed_at':''})
        elif rec.get('status')=='Pending':
            rec.update({'minutes':minutes,'actual_seconds':actual_seconds,'target_seconds':target_seconds,'updated_at':now})
        save_overtime_approvals(rows)
        return rec
    rec={'id':uuid.uuid4().hex,'user_id':user_id,'name':name,'store':store,'date':date_str,'minutes':minutes,'actual_seconds':actual_seconds,'target_seconds':target_seconds,'status':'Pending','created_at':now,'updated_at':now,'reviewed_by':'','reviewed_at':''}
    rows.append(rec);save_overtime_approvals(rows)
    # Store-targeted notification appears in that store admin portal; developers also see it.
    suite_notify(store,'Overtime approval required',f'{name} ({user_id}) worked {minutes} extra minutes on {date_str}.',store)
    return rec



def _normalize_salary_text(value):
    value = value or ''
    value = unicodedata.normalize('NFKD', value)
    value = ''.join(ch for ch in value if not unicodedata.combining(ch))
    value = value.upper()
    value = re.sub(r'[^A-Z0-9]+', ' ', value)
    return re.sub(r'\s+', ' ', value).strip()

def _salary_employee_candidates(page_text):
    text = _normalize_salary_text(page_text)
    if not text:
        return []
    candidates = []
    for emp_code, emp_info in MASTER_EMPLOYEES.items():
        clean_code = emp_code.upper().replace('NWC', '')
        if (re.search(r'\bNWC\s*' + re.escape(clean_code) + r'\b', text)
                or re.search(r'\b' + re.escape(clean_code) + r'\b', text)):
            candidates.append(emp_code)
    if candidates:
        return list(dict.fromkeys(candidates))
    for emp_code, emp_info in MASTER_EMPLOYEES.items():
        normalized_name = _normalize_salary_text(emp_info.get('name', ''))
        if normalized_name and normalized_name in text:
            candidates.append(emp_code)
    return list(dict.fromkeys(candidates))

def _build_salary_pdf_groups(reader):
    employee_order = sorted(MASTER_EMPLOYEES.keys(), key=lambda x: get_emp_info(x)['name'])
    groups = {code: [] for code in employee_order}
    unmatched_pages = []
    current_emp = None

    for page_index, page in enumerate(reader.pages):
        try:
            page_text = page.extract_text() or ''
        except Exception:
            page_text = ''
        candidates = _salary_employee_candidates(page_text)
        if candidates:
            emp_code = current_emp if current_emp in candidates else candidates[0]
            groups[emp_code].append(page_index)
            current_emp = emp_code
        else:
            unmatched_pages.append(page_index)

    if unmatched_pages and current_emp:
        assigned = {code for code, pages in groups.items() if pages}
        missing = [code for code in employee_order if code not in assigned]
        reserve = min(len(unmatched_pages), len(missing))
        continuation_pages = unmatched_pages[reserve:]
        groups[current_emp].extend(continuation_pages)
        unmatched_pages = unmatched_pages[:reserve]

    missing = [code for code in employee_order if not groups[code]]
    for page_index, emp_code in zip(unmatched_pages, missing):
        groups[emp_code].append(page_index)

    consumed = min(len(unmatched_pages), len(missing))
    leftover_pages = unmatched_pages[consumed:]
    if leftover_pages:
        fallback_emp = current_emp or (missing[-1] if missing else employee_order[-1])
        groups[fallback_emp].extend(leftover_pages)

    return {code: pages for code, pages in groups.items() if pages}

def get_emp_info(emp_code):
    emp_str = str(emp_code).strip()
    if not emp_str.startswith('NWC') and f"NWC{emp_str}" in load_users_db(): emp_str = f"NWC{emp_str}"
    
    db = load_users_db()
    if emp_str in db: return db[emp_str]
    return {'name': f'Employee {emp_code}', 'off': 'SUNDAY', 'dept': 'General', 'shift': 'morning'}

def check_device_connectivity(store_code):
    machine = MACHINES.get(store_code, MACHINES['LM11'])
    try:
        zk = ZK(machine['ip'], port=machine['port'], timeout=2, password=0, force_udp=False, ommit_ping=False)
        conn = zk.connect()
        if conn:
            conn.disconnect()
            return True
    except Exception: pass
    
    # Store-specific sync tracking check
    last_sync = LAST_DEVICE_SYNC_TIME.get(store_code)
    if last_sync and (datetime.now() - last_sync).total_seconds() < 300: return True
    return False

def fetch_attendance_data(start_date_str, end_date_str, filter_user_id, store_code):
    machine = MACHINES.get(store_code, MACHINES['LM11'])
    device_online = check_device_connectivity(store_code)
    period_data = {}
    raw_punches_list = []
    
    db = load_users_db()
    users_map_temp = {k: {'code': k, 'name': v['name']} for k, v in db.items() if v.get('role') == 'employee'}
        
    for uid_override, over_data in EMPLOYEE_OVERRIDES.items():
        users_map_temp[str(uid_override)] = {'code': over_data['code'], 'name': over_data['name']}

    attendance_records = []
    try:
        zk = ZK(machine['ip'], port=machine['port'], timeout=2, password=0, force_udp=False, ommit_ping=False)
        conn = zk.connect()
        if conn:
            users = conn.get_users()
            for user in users:
                uid_str = str(user.user_id)
                emp_code = f"NWC{uid_str}" if not uid_str.startswith('NWC') else uid_str
                emp_name = user.name if user.name else get_emp_info(emp_code)['name']
                users_map_temp[uid_str] = {'code': emp_code, 'name': emp_name}
                
            attendance = conn.get_attendance()
            for att in attendance:
                attendance_records.append({'user_id': str(att.user_id), 'timestamp': att.timestamp})
            conn.disconnect()
    except Exception: pass

    if SYNCED_ATTENDANCE_LOGS:
        for log in SYNCED_ATTENDANCE_LOGS:
            ts = log['timestamp']
            if isinstance(ts, str):
                try: ts = datetime.strptime(ts, '%Y-%m-%d %H:%M:%S')
                except ValueError: 
                    try: ts = datetime.fromisoformat(ts)
                    except Exception: continue
            attendance_records.append({'user_id': str(log['user_id']), 'timestamp': ts})

    manual_punches = load_manual_punches()
    for mp in manual_punches:
        try:
            ts = datetime.strptime(mp['timestamp'], '%Y-%m-%d %H:%M:%S')
            attendance_records.append({'user_id': str(mp['user_id']), 'timestamp': ts})
        except Exception: continue

    # Process chronologically so the earliest punch is retained and later punches within 60 seconds are duplicates.
    attendance_records.sort(key=lambda x: x.get('timestamp') or datetime.min)
    last_accepted_punch = {}
    duplicate_batch = []
    for att in attendance_records:
        att_ts = att['timestamp']
        att_date_str = att_ts.strftime('%Y-%m-%d')
        if start_date_str <= att_date_str <= end_date_str:
            raw_uid = str(att['user_id'])
            if raw_uid in EMPLOYEE_OVERRIDES:
                emp_code = EMPLOYEE_OVERRIDES[raw_uid]['code']
                emp_name = EMPLOYEE_OVERRIDES[raw_uid]['name']
            elif raw_uid in users_map_temp:
                emp_code = users_map_temp[raw_uid]['code']
                emp_name = users_map_temp[raw_uid]['name']
            else:
                clean_uid = raw_uid.replace('NWC', '')
                emp_code = f"NWC{clean_uid}" if not clean_uid.startswith('NWC') else clean_uid
                emp_name = get_emp_info(emp_code)['name']
                
            # Filter users belonging to current store unless developer
            if store_code != 'DEV' and db.get(emp_code, {}).get('store') != store_code:
                continue

            if filter_user_id and filter_user_id != 'ALL' and emp_code != filter_user_id and raw_uid != filter_user_id:
                continue

            punch_key = (emp_code, att_date_str)
            previous_accepted = last_accepted_punch.get(punch_key)
            if previous_accepted is not None:
                gap_seconds = int((att_ts - previous_accepted).total_seconds())
                if 0 <= gap_seconds <= DUPLICATE_PUNCH_WINDOW_SECONDS:
                    duplicate_batch.append({'id':f'{store_code}|{emp_code}|{att_ts.strftime("%Y-%m-%d %H:%M:%S")}', 'store':store_code, 'user_id':emp_code, 'name':emp_name, 'date':att_date_str, 'duplicate_time':att_ts.strftime('%H:%M:%S'), 'kept_time':previous_accepted.strftime('%H:%M:%S'), 'gap_seconds':gap_seconds, 'reason':'Second punch within 1 minute', 'detected_at':datetime.now().strftime('%Y-%m-%d %H:%M:%S')})
                    continue
            last_accepted_punch[punch_key] = att_ts
            raw_punches_list.append({'date': att_date_str, 'time': att_ts.strftime('%H:%M:%S'), 'user_id': emp_code, 'name': emp_name, 'timestamp': att_ts})
            if att_date_str not in period_data: period_data[att_date_str] = {}
            if emp_code not in period_data[att_date_str]: period_data[att_date_str][emp_code] = {'name': emp_name, 'timestamps': []}
            period_data[att_date_str][emp_code]['timestamps'].append(att_ts)
            
    if duplicate_batch:
        existing_duplicates = load_json_file(DUPLICATE_PUNCHES_FILE)
        if not isinstance(existing_duplicates,list): existing_duplicates=[]
        known_ids={x.get('id') for x in existing_duplicates}
        existing_duplicates.extend(x for x in duplicate_batch if x.get('id') not in known_ids)
        save_json_file(DUPLICATE_PUNCHES_FILE,existing_duplicates[-10000:])

    users_list = []
    for k, v in db.items():
        if v.get('role') == 'employee' and (store_code == 'DEV' or v.get('store') == store_code):
            users_list.append({'user_id': k, 'name': v['name'], 'dept': v.get('dept', ''), 'off': v.get('off', '')})
    users_list = sorted(users_list, key=lambda x: x['name'])

    final_data = []
    total_duration_seconds, total_lunch_seconds, total_net_variance_seconds = 0, 0, 0
    present_count, absent_count, off_count, mis_punch_count, late_arrival_count, ml_count, shift_a_count, shift_b_count = 0, 0, 0, 0, 0, 0, 0, 0
    
    # Process ALL dates in range to show accurate weekly off records
    start_dt = datetime.strptime(start_date_str, '%Y-%m-%d')
    end_dt = datetime.strptime(end_date_str, '%Y-%m-%d')
    delta = end_dt - start_dt
    dates_to_process = [(start_dt + timedelta(days=i)).strftime('%Y-%m-%d') for i in range(delta.days + 1)]
    dates_to_process.sort(reverse=True)

    overrides = load_overrides()
    global_roster = load_roster()

    for date_str in dates_to_process:
        day_users_dict = period_data.get(date_str, {})
        current_dt = datetime.strptime(date_str, '%Y-%m-%d')
        current_day_name = current_dt.strftime('%A').upper()
        is_weekend = current_dt.weekday() >= 5

        present_records, absent_records, off_records, mispunch_records, ml_records = [], [], [], [], []

        for emp_code, emp_info in db.items():
            if emp_info.get('role') != 'employee': continue
            if store_code != 'DEV' and emp_info.get('store') != store_code: continue
            
            final_emp_code = emp_code
            emp_name, emp_dept, emp_off, emp_shift = emp_info['name'], emp_info.get('dept', ''), emp_info.get('off', '').upper(), emp_info.get('shift', 'morning')

            roster_val = global_roster.get(final_emp_code, {}).get(date_str)
            if roster_val:
                if roster_val == 'Weekly Off': emp_off = current_day_name
                elif roster_val in ['Shift A', 'Shift B']:
                    emp_off = 'NONE' 
                    emp_shift = 'morning' if roster_val == 'Shift A' else 'second'

            if filter_user_id and filter_user_id != 'ALL' and final_emp_code != filter_user_id:
                continue

            approved_leave_obj = next((l for l in LEAVE_REQUESTS if l['user_id'] == final_emp_code and l['status'] == 'Approved' and l['start_date'] <= date_str <= l['end_date']), None)

            if approved_leave_obj:
                ml_count += 1
                leave_type_code = approved_leave_obj.get('leave_type', 'F10;1')
                ml_records.append({
                    'date': date_str, 'user_id': final_emp_code, 'name': emp_name, 'dept': emp_dept,
                    'store_in': f'Approved Leave ({leave_type_code})', 'lunch_out': '-', 'lunch_in': '-', 'out_time': '-',
                    'total_lunch': '-', 'lunch_seconds': 3600, 'net_duration_seconds': 0, 'total_hours': '-', 'net_variance': '-', 'variance_type': 'neutral',
                    'status': f'{leave_type_code} (Leave)', 'is_late': 'No', 'shift_type': '-'
                })
                continue

            matched_key = final_emp_code if final_emp_code in day_users_dict else None
            day_ovr = overrides.get(date_str, {}).get(final_emp_code, {})
            has_override = len(day_ovr) > 0

            if matched_key or has_override:
                present_count += 1
                data_obj = day_users_dict.get(matched_key, {'timestamps': []}) if matched_key else {'timestamps': []}
                sorted_times = sorted(list(set(data_obj['timestamps'])))
                total_punches = len(sorted_times)
                
                store_in = sorted_times[0].strftime('%H:%M:%S') if total_punches > 0 else '-'
                lunch_out = sorted_times[1].strftime('%H:%M:%S') if total_punches >= 3 else '-'
                lunch_in = sorted_times[2].strftime('%H:%M:%S') if total_punches >= 3 else '-'
                out_time = sorted_times[-1].strftime('%H:%M:%S') if total_punches > 1 else '-'

                if 'store_in' in day_ovr: store_in = day_ovr['store_in']
                if 'lunch_out' in day_ovr: lunch_out = day_ovr['lunch_out']
                if 'lunch_in' in day_ovr: lunch_in = day_ovr['lunch_in']
                if 'out_time' in day_ovr: out_time = day_ovr['out_time']

                def get_sec(t_s):
                    if t_s == '-': return None
                    try:
                        h, m, s = map(int, t_s.split(':'))
                        return h * 3600 + m * 60 + s
                    except: return None

                s_in, l_o, l_i, s_out = get_sec(store_in), get_sec(lunch_out), get_sec(lunch_in), get_sec(out_time)
                shift_type, is_late = '-', False
                if s_in is not None:
                    if s_in <= 10 * 3600:
                        shift_type, shift_a_count = 'Shift A', shift_a_count + 1
                    else:
                        shift_type, shift_b_count = 'Shift B', shift_b_count + 1
                    
                    limit_sec = (13 * 3600 + 10 * 60) if emp_shift == 'second' else (7 * 3600)
                    if s_in > limit_sec:
                        late_arrival_count += 1; is_late = True

                lunch_seconds = 0
                has_lunch_punches = False
                if l_o is not None and l_i is not None:
                    act_l = l_i - l_o
                    lunch_seconds = 3600 if act_l < 3600 else act_l
                    has_lunch_punches = True

                if has_lunch_punches:
                    total_lunch_seconds += lunch_seconds
                    l_hrs = divmod(lunch_seconds, 3600)
                    total_lunch_str = f"{l_hrs[0]}h {l_hrs[1]//60}m"
                else: total_lunch_str = "-"
                
                # Working time is always first punch to last punch minus a valid lunch interval.
                # The old calculation incorrectly used a 7-hour target when lunch punches were absent.
                actual_net_seconds = max(0, (s_out - s_in) - lunch_seconds) if (s_in is not None and s_out is not None and s_out > s_in) else 0
                target_seconds = (8 * 3600) if has_lunch_punches else (7 * 3600)
                net_duration_seconds = actual_net_seconds
                net_variance_str, variance_type = "-", "neutral"
                overtime_status = "None"
                overtime_minutes = 0
                if s_in is not None and s_out is not None:
                    diff = actual_net_seconds - target_seconds
                    if diff >= 45 * 60:
                        raw_overtime_minutes = diff // 60
                        # 45m-1h44m => 1 credited OT hour; 1h45m-2h44m => 2 hours, etc.
                        overtime_hours = max(1, int((diff + 15 * 60) // 3600))
                        overtime_minutes = overtime_hours * 60
                        rec = ensure_overtime_request(final_emp_code,emp_name,emp_info.get('store',store_code),date_str,overtime_minutes,actual_net_seconds,target_seconds)
                        rec['raw_minutes'] = raw_overtime_minutes
                        overtime_status = rec.get('status','Pending')
                        if overtime_status == 'Approved':
                            net_duration_seconds = target_seconds + overtime_minutes * 60
                            net_variance_str, variance_type = f"Approved OT {overtime_minutes//60}h {overtime_minutes%60}m", 'positive'
                        elif session.get('role') == 'employee':
                            # Employees only see overtime after approval. Pending/rejected extra time is hidden and not credited.
                            net_duration_seconds = min(actual_net_seconds,target_seconds)
                            net_variance_str, variance_type = "0h 0m", "neutral"
                        elif overtime_status == 'Rejected':
                            # Admin/developer retain complete history; rejected time is not credited.
                            net_duration_seconds = min(actual_net_seconds,target_seconds)
                            net_variance_str, variance_type = f"Rejected OT {overtime_minutes//60}h {overtime_minutes%60}m", "rejected"
                        else:
                            net_variance_str, variance_type = f"Pending OT {overtime_minutes//60}h {overtime_minutes%60}m", 'pending'
                    elif diff < 0:
                        short=abs(diff);s_hrs=divmod(short,3600)
                        net_variance_str,variance_type=f"-{s_hrs[0]}h {s_hrs[1]//60}m",'negative'
                    else:
                        net_variance_str,variance_type="0h 0m",'neutral'
                total_duration_seconds += net_duration_seconds
                total_net_variance_seconds += (net_duration_seconds-target_seconds) if (s_in is not None and s_out is not None) else 0
                total_hours_str = f"{divmod(net_duration_seconds, 3600)[0]}h {divmod(net_duration_seconds, 3600)[1]//60}m" if s_in is not None and s_out is not None else "-"
                
                if s_in is not None and s_out is None:
                    status, mis_punch_count = 'Mis Punch', mis_punch_count + 1
                elif s_in is None and s_out is None: status = 'Absent'
                else: status = 'Present'
                if current_day_name == emp_off: status = 'Weekly Off'

                record = {
                    'date': date_str, 'user_id': final_emp_code, 'name': emp_name, 'dept': emp_dept,
                    'store_in': store_in, 'lunch_out': lunch_out, 'lunch_in': lunch_in, 'out_time': out_time, 
                    'total_lunch': total_lunch_str, 'lunch_seconds': lunch_seconds, 'net_duration_seconds': net_duration_seconds, 
                    'total_hours': total_hours_str, 'net_variance': net_variance_str, 'variance_type': variance_type,
                    'status': status, 'is_late': 'Yes' if is_late else 'No', 'shift_type': shift_type, 'overtime_status': overtime_status, 'overtime_minutes': overtime_minutes
                }
                if status == 'Weekly Off': off_records.append(record)
                elif status == 'Mis Punch': mispunch_records.append(record)
                else: present_records.append(record)
            else:
                if current_day_name == emp_off:
                    off_count += 1
                    off_records.append({
                        'date': date_str, 'user_id': final_emp_code, 'name': emp_name, 'dept': emp_dept,
                        'store_in': '-', 'lunch_out': '-', 'lunch_in': '-', 'out_time': '-', 'total_lunch': '-', 
                        'lunch_seconds': 0, 'net_duration_seconds': 0, 'total_hours': '-', 'net_variance': 'Off', 
                        'variance_type': 'neutral', 'status': 'Weekly Off', 'is_late': 'No', 'shift_type': '-'
                    })
                else:
                    absent_count += 1
                    absent_records.append({
                        'date': date_str, 'user_id': final_emp_code, 'name': emp_name, 'dept': emp_dept,
                        'store_in': '-', 'lunch_out': '-', 'lunch_in': '-', 'out_time': '-', 'total_lunch': '-', 
                        'lunch_seconds': 0, 'net_duration_seconds': 0, 'total_hours': '-', 'net_variance': '-', 
                        'variance_type': 'neutral', 'status': 'Absent', 'is_late': 'No', 'shift_type': '-'
                    })
        final_data.extend(present_records + mispunch_records + off_records + absent_records + ml_records)
            
    tot_hrs = divmod(total_duration_seconds, 3600)
    grand_total_hours = f"{tot_hrs[0]}h {tot_hrs[1]//60}m"
    tot_l_hrs = divmod(total_lunch_seconds, 3600)
    grand_total_lunch_hours = f"{tot_l_hrs[0]}h {tot_l_hrs[1]//60}m"

    v_sec = total_net_variance_seconds
    if v_sec >= 0: grand_total_variance, grand_variance_type = f"+{divmod(v_sec, 3600)[0]}h {divmod(v_sec, 3600)[1]//60}m", 'positive'
    else: grand_total_variance, grand_variance_type = f"-{divmod(abs(v_sec), 3600)[0]}h {divmod(abs(v_sec), 3600)[1]//60}m", 'negative'
    
    stats_summary = {
        'present': present_count, 'absent': absent_count, 'off': off_count, 'mispunch': mis_punch_count,
        'late_arrival': late_arrival_count, 'ml': ml_count, 'shift_a': shift_a_count, 'shift_b': shift_b_count,
        'total_hrs': grand_total_hours, 'total_lunch_hrs': grand_total_lunch_hours,
        'total_variance': grand_total_variance, 'variance_type': grand_variance_type, 'device_online': device_online
    }
    raw_punches_list = sorted(raw_punches_list, key=lambda x: x['timestamp'], reverse=True)
    return final_data, users_list, grand_total_hours, grand_total_lunch_hours, grand_total_variance, raw_punches_list, stats_summary

LOGIN_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Login | Attendance Portal</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>body { font-family: 'Inter', sans-serif; }</style>
</head>
<body class="bg-slate-900 min-h-screen flex items-center justify-center p-4 relative"><div class="absolute top-4 right-4"><select onchange="location.href='/set_language/'+this.value" class="bg-white text-slate-900 text-xs font-bold rounded-xl px-3 py-2"><option value="en" {% if session.get('ui_language','en')=='en' %}selected{% endif %}>English</option><option value="pt" {% if session.get('ui_language')=='pt' %}selected{% endif %}>Português</option></select></div>
    <div class="bg-white/95 backdrop-blur-md rounded-3xl shadow-2xl border border-slate-200/50 p-8 w-full max-w-md space-y-6">
        <div class="text-center space-y-2">
            <div class="inline-flex bg-[#78b13f] px-5 py-3 rounded-2xl shadow-lg mb-2 items-center justify-center">
                <img src="{{ url_for('static', filename='fresmart.png') }}" alt="Logo" class="h-12 object-contain" onerror="this.style.display='none'">
            </div>
            <h1 class="text-2xl font-black text-slate-900 tracking-tight">Attendance Portal</h1>
            <p class="text-xs text-slate-500 font-medium">Developed by Sonu Kumar <span class="text-emerald-600 font-semibold">(NCSA0608)</span></p>
        </div>
        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}{% for category, message in messages %}<div class="hrms-flash hidden" data-category="{{ category }}" data-message="{{ message|e }}"></div>{% endfor %}{% endif %}
        {% endwith %}
        {% if error %}<div class="hrms-flash hidden" data-category="danger" data-message="{{ error|e }}"></div>{% endif %}
        <form method="POST" action="/login" class="space-y-4">
            <div>
                <label class="block text-xs font-bold uppercase tracking-wider text-slate-500 mb-2">Store/Employee Code</label>
                <input type="text" name="user_id" required placeholder="NWC1234 or Store Code" class="w-full bg-slate-50 border border-slate-300 rounded-xl px-4 py-3 text-sm font-medium focus:ring-2 focus:ring-emerald-500 focus:outline-none">
            </div>
            <div>
                <label class="block text-xs font-bold uppercase tracking-wider text-slate-500 mb-2">Password</label>
                <input type="password" name="password" required placeholder="Enter Password" class="w-full bg-slate-50 border border-slate-300 rounded-xl px-4 py-3 text-sm font-medium focus:ring-2 focus:ring-emerald-500 focus:outline-none">
            </div>
            <button type="submit" class="w-full bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs uppercase tracking-widest py-3.5 rounded-xl shadow-lg transition duration-200">Secure Login 🚀</button>
            <div class="text-center mt-3">
                <a href="/reset_password" class="text-[11px] text-emerald-600 font-bold hover:underline transition">Forgot/Reset Password?</a>
            </div>
        </form>
    </div>
<script>window.HRMS_LANGUAGE={{ session.get('ui_language','en')|tojson }};window.HRMS_PT={'Save Changes':'Guardar Alterações','Manage Pictures':'Gerir Imagens','Upload More Pictures':'Carregar Mais Imagens','Delete Picture':'Eliminar Imagem','Edit / Upload':'Editar / Carregar','Select offer period':'Seleccionar período de oferta','Upload Pictures to Offer Period':'Carregar Imagens para o Período de Oferta','Create Offer Period':'Criar Período de Oferta','Create Offer Period Link':'Criar Ligação do Período de Oferta','Offer Period Links':'Ligações dos Períodos de Oferta','Select at least one article.':'Seleccione pelo menos um artigo.','Select all visible rows':'Seleccionar todas as linhas visíveis','Deletion is permanent.':'A eliminação é permanente.','Delete All Existing Data':'Eliminar Todos os Dados Existentes','Delete Selected':'Eliminar Seleccionados','Click to sort':'Clique para ordenar','Other':'Outro','Search article code or name':'Pesquisar código ou nome do artigo','All Article Types':'Todos os Tipos de Artigo','Automatic one-minute duplicate filtering':'Filtragem automática de duplicados no intervalo de um minuto','No duplicate punches found.':'Nenhuma marcação duplicada encontrada.','Gap':'Intervalo','Duplicate Punch':'Marcação Duplicada','Kept Punch':'Marcação Mantida','Duplicate Punches':'Marcações Duplicadas','Select one or more documents first.':'Seleccione primeiro um ou mais documentos.','file(s) selected':'ficheiro(s) seleccionado(s)','You can select multiple files at once.':'Pode seleccionar vários ficheiros ao mesmo tempo.','No Excel data imported yet.':'Ainda não foram importados dados do Excel.','Supporting Documents':'Documentos de Suporte','Upload & Import Excel':'Carregar e Importar Excel','Import Article Excel':'Importar Excel de Artigos','Article Name':'Nome do Artigo','Article Code':'Código do Artigo','Article Type':'Tipo de Artigo','Power SKU':'SKU de Potência','Total Articles':'Total de Artigos','Available Promotion Documents':'Documentos de Promoção Disponíveis','Upload Document':'Carregar Documento','KVI + Power SKU':'KVI + SKU de Potência','Promotion':'Promoção','Credited Overtime':'Horas Extra Creditadas','Credited Working Hours':'Horas de Trabalho Creditadas','Export Last 6 Months':'Exportar Últimos 6 Meses','Add / Manage Employees':'Adicionar / Gerir Trabalhadores','Edit Attendance Times':'Editar Horas de Assiduidade','Attendance View':'Ver Assiduidade','Portal Options for New User':'Opções do Portal para Novo Utilizador','Clear All':'Limpar Tudo','Select All':'Seleccionar Tudo','Store Access':'Acesso à Loja','Save Rights':'Guardar Direitos','Developer can allow or hide every portal option for existing and new users.':'O programador pode permitir ou ocultar cada opção do portal para utilizadores existentes e novos.','Portal Rights Control':'Controlo de Direitos do Portal','Identification':'Identificação','Contract Date':'Data de Contrato','Open / Download ID Card':'Abrir / Baixar Cartão de Identificação','All employee identification cards in one place':'Todos os cartões de identificação dos trabalhadores num único local','Employee ID Cards':'Cartões de Identificação dos Trabalhadores','No overtime history':'Sem histórico de horas extra','No pending overtime requests':'Sem pedidos de horas extra pendentes','Reviewed At':'Revisto Em','Detected Extra':'Tempo Extra Detectado','All Status':'Todos os Estados','Complete employee-wise Pending, Approved and Rejected history.':'Histórico completo por trabalhador: Pendente, Aprovado e Rejeitado.','Only pending overtime requests are shown here.':'Apenas os pedidos de horas extra pendentes são apresentados aqui.','Overtime History':'Histórico de Horas Extra','Overtime Approval':'Aprovação de Horas Extra','Rejected OT':'HE Rejeitada','Search employee name or code':'Pesquisar nome ou código do trabalhador','Photo updated':'Fotografia actualizada','Change Photo':'Alterar Fotografia','Latest ID Card':'Cartão de Identificação Actual','Approved OT':'HE Aprovada','Pending OT':'HE Pendente','No overtime records':'Sem registos de horas extra','Reviewed By':'Revisto Por','Extra Time':'Tempo Extra','Approve or reject automatically detected extra working time.':'Aprovar ou rejeitar horas extra detectadas automaticamente.','Overtime Approvals':'Aprovações de Horas Extra','Correction Rejected':'Correcção Rejeitada','Correction Approved':'Correcção Aprovada','Correction Request':'Pedido de Correcção','New Document':'Novo Documento','was Rejected':'foi rejeitada','was Approved':'foi aprovada','Your correction':'A sua correcção','Your roster':'A sua escala','Month select karke Load Month click karein':'Seleccione o mês e clique em Carregar Mês','Employee-wise Shift A, Shift B and Weekly Off allocation':'Alocação por trabalhador de Turno A, Turno B e Folga Semanal','Public Holiday':'Feriado Público','Half Day':'Meio Dia','Full Day':'Dia Completo','Full Month':'Mês Completo','Clear All':'Limpar Tudo','Select All':'Seleccionar Tudo','Uploaded By':'Carregado por','Updated By':'Actualizado por','Created On':'Criado em','Created By':'Criado por','Records':'Registos','Record':'Registo','Summary':'Resumo','Details':'Detalhes','Code':'Código','Address':'Endereço','Phone':'Telefone','Email':'E-mail','Photo':'Fotografia','Store Access':'Acesso à Loja','Apply Selected':'Aplicar Seleccionados','Copy Previous Month':'Copiar Mês Anterior','Copy Previous Week':'Copiar Semana Anterior','Monthly roster saved':'Escala mensal guardada','Kam se kam ek option select karein.':'Seleccione pelo menos uma opção.','Month select karein':'Seleccione o mês','Galat User ID ya Password!':'ID do utilizador ou palavra-passe incorrectos!','Sirf PDF files allowed hain!':'Apenas ficheiros PDF são permitidos!','Sabhi fields bharna zaroori hai!':'Todos os campos são obrigatórios!','This roster shows only your shifts and weekly-off dates for the current month.':'Esta escala mostra apenas os seus turnos e folgas semanais do mês actual.','Overtime Minutes':'Minutos de Horas Extra','Submit Overtime':'Enviar Horas Extra','Email - requires SMTP':'E-mail - requer SMTP','Internal Notification':'Notificação Interna','Download Center':'Centro de Transferências','Monthly':'Mensal','Weekly':'Semanal','Daily':'Diário','Delivery':'Entrega','Frequency':'Frequência','File':'Ficheiro','Uploaded At':'Carregado em','Created':'Criado','To':'Para','From':'De','Effective':'Efectivo','Estimate':'Estimativa','Employer':'Empregador','Gross':'Bruto','Configurable estimate using 2026 settings. Validate payroll categories, taxable base and exemptions before finalisation.':'Estimativa configurável com definições de 2026. Valide as categorias salariais, a base tributável e as isenções antes da finalização.','Identification masking':'Ocultação da Identificação','Document access log':'Registo de Acesso a Documentos','Access audit':'Auditoria de Acesso','Last backups':'Últimas Cópias de Segurança','Number of Users':'Número de Utilizadores','Application Version':'Versão da Aplicação','retry required':'nova tentativa necessária','Last saved attendance remains available':'A última assiduidade guardada continua disponível','Queue State':'Estado da Fila','Level':'Nível','Late Count':'Número de Atrasos','Rolling 30-day view.':'Vista móvel de 30 dias.','roster cells copied':'células de escala copiadas','assignments saved':'alocações guardadas','Copy failed':'Falha ao copiar','Save Lifecycle Event':'Guardar Evento do Ciclo de Vida','Employee Lifecycle':'Ciclo de Vida do Trabalhador','New Records':'Novos Registos','Records Received':'Registos Recebidos','Machine Status':'Estado do Equipamento','Store Code':'Código da Loja','Payroll, holidays, lifecycle, overtime, backups and privacy controls.':'Salários, feriados, ciclo de vida, horas extra, cópias de segurança e controlos de privacidade.','Privacy-safe QR verification cards.':'Cartões QR de verificação com protecção de privacidade.','Internal leave, shift, payroll and contract notices.':'Avisos internos de ausências, turnos, salários e contratos.','Track user, store, action, IP and changes.':'Acompanhar utilizador, loja, acção, IP e alterações.','Minimum staffing rules and shortage warnings.':'Regras de dotação mínima e avisos de falta de pessoal.','Copy week or month schedules.':'Copiar escalas semanais ou mensais.','Contracts, IDs and certificates.':'Contratos, documentos de identificação e certificados.','Annual, used, pending and available balances.':'Saldos anuais, usados, pendentes e disponíveis.','Approve missing-punch correction requests.':'Aprovar pedidos de correcção de marcações em falta.','Machine status, manual checks and sync history.':'Estado dos equipamentos, verificações manuais e histórico de sincronização.','Combined attendance with store code and Excel export.':'Assiduidade combinada com código da loja e exportação Excel.','Multi-store operations, compliance and employee self-service':'Operações multi-loja, conformidade e auto-serviço do trabalhador','No contract alerts.':'Sem alertas de contrato.','Audit Events':'Eventos de Auditoria','Contracts':'Contratos','Corrections':'Correcções','Machine':'Equipamento','Store Comparison':'Comparação de Lojas','Monthly Summary':'Resumo Mensal','Default Schedule':'Horário Padrão','All Allocations':'Todas as Alocações','Designations Selected':'Funções Seleccionadas','All Designations Selected':'Todas as Funções Seleccionadas','Same Weekday for Full Month':'Mesmo Dia da Semana durante Todo o Mês','Monday-Sunday':'Segunda a Domingo','Last Week':'Última Semana','5th Week':'5.ª Semana','4th Week':'4.ª Semana','3rd Week':'3.ª Semana','2nd Week':'2.ª Semana','1st Week':'1.ª Semana','All Weeks (Monday-Sunday)':'Todas as Semanas (Segunda a Domingo)','Not in this month':'Não existe neste mês','options selected':'opções seleccionadas','option selected':'opção seleccionada','Apply to selected date':'Aplicar à data seleccionada','Month & Year':'Mês e Ano','load month':'carregar mês','Select Month':'Seleccionar Mês','No shift requests':'Sem pedidos de turno','Current Assignment':'Alocação Actual','Requested':'Solicitado','Shift Request':'Pedido de Turno','Request Shift / Off':'Pedir Turno / Folga','Select Employee':'Seleccionar Trabalhador','Bulk Upload':'Carregamento em Massa','Standard Row':'Linha Padrão','Monthly grid':'Grelha mensal','Master':'Principal','Individual Slip':'Recibo Individual','Split':'Separar','All Personnel':'Todo o Pessoal','No requests':'Sem pedidos','Minutes':'Minutos','ALL':'TODOS','Change':'Alteração','Emp':'Trab.','Remove Photo':'Remover Fotografia','STORE ACCESS':'ACESSO À LOJA','Create login and employee profile':'Criar início de sessão e perfil do trabalhador','No employees found.':'Nenhum trabalhador encontrado.','Select Designation':'Seleccionar Função','Add Employee':'Adicionar Trabalhador','Search, review and maintain workforce profiles':'Pesquisar, rever e manter perfis dos trabalhadores','Open Roster Planner':'Abrir Planeador de Escala','Use Roster Planner to assign shifts and weekly-off rotations.':'Utilize o Planeador de Escala para atribuir turnos e rotações de folga semanal.','Rota Rotation and Shift Schedules are integrated with the':'A rotação da escala e os horários de turno estão integrados com o','No salary slips uploaded yet.':'Ainda não foram carregados recibos de salário.','View':'Ver','PDF File':'Ficheiro PDF','-- Select Employee --':'-- Seleccionar Trabalhador --','Individual Slip Upload':'Carregar Recibo Individual','Split & Upload':'Separar e Carregar','Master PDF':'PDF Principal','Bulk Upload (Merged PDF)':'Carregamento em Massa (PDF Unido)','Monthly grid with Leave Codes':'Grelha mensal com códigos de ausência','Employee Matrix':'Matriz de Trabalhadores','Individual records per date':'Registos individuais por data','Standard Row Export':'Exportação em Linhas','Download Attendance Reports':'Baixar Relatórios de Assiduidade','Total Summary':'Resumo Total','No attendance records found for this selection.':'Nenhum registo de assiduidade encontrado para esta selecção.','Status ↕':'Estado ↕','Working Hrs ↕':'Horas Trabalhadas ↕','Total Lunch ↕':'Total de Almoço ↕','Out Time ↕':'Hora de Saída ↕','Store In ↕':'Entrada na Loja ↕','Dept ↕':'Departamento ↕','Employee Name ↕':'Nome do Trabalhador ↕','Date ↕':'Data ↕','Last Month':'Mês Passado','This Month':'Este Mês','This Week':'Esta Semana','Yesterday':'Ontem','Quick Range':'Intervalo Rápido','Rota':'Escala','Export':'Exportar','Logged In As':'Sessão Iniciada Como','-- All Personnel --':'-- Todo o Pessoal --','Employee Filter':'Filtro de Trabalhadores','Tot Hrs':'Total de Horas','Mis Punch':'Marcação em Falta','Mis-Punch':'Marcação em Falta','Late Arr.':'Chegada Tardia','Week Off':'Folga Semanal','Shutdown':'Encerrar','Sync':'Sincronização','Device':'Equipamento','Leaves':'Ausências','Dev':'Programador','Biometric live tracking active for Attendance Portal.':'Acompanhamento biométrico em tempo real activo no Portal de Assiduidade.','Announcements':'Comunicados','Employees Info (ID Card)':'Informações dos Trabalhadores (Cartão de Identificação)','Manage Passwords':'Gerir Palavras-passe','Biometric Machines':'Equipamentos Biométricos','Shift Approvals':'Aprovações de Turno','Reset':'Repor','Off':'Folga','Daily manpower summary':'Resumo diário de efectivos','Loading monthly roster...':'A carregar a escala mensal...','Send Request for Approval':'Enviar Pedido para Aprovação','Developed by':'Desenvolvido por','Privacy-safe QR verification cards.':'Cartões QR de verificação com protecção de privacidade.','Internal leave, shift, payroll and contract notices.':'Avisos internos de ausências, turnos, salários e contratos.','Track user, store, action, IP and changes.':'Acompanhar utilizador, loja, acção, IP e alterações.','Contracts, IDs and certificates.':'Contratos, documentos de identificação e certificados.','Approve missing-punch correction requests.':'Aprovar pedidos de correcção de marcações em falta.','Combined attendance with store code and Excel export.':'Assiduidade combinada com código da loja e exportação Excel.','Machine status, manual checks and sync history.':'Estado dos equipamentos, verificações manuais e histórico de sincronização.','Annual, used, pending and available balances.':'Saldos anuais, usados, pendentes e disponíveis.','Copy week or month schedules.':'Copiar escalas semanais ou mensais.','Open module':'Abrir módulo','Framework':'Enquadramento','Attendance Portal':'Portal de Assiduidade','Sign in to access your dashboard':'Inicie sessão para aceder ao seu painel','User ID':'ID do Utilizador','Password':'Palavra-passe','Secure Login':'Iniciar Sessão','Forgot/Reset Password?':'Esqueceu/Redefinir Palavra-passe?','Reset Password':'Redefinir Palavra-passe','Create a new password request':'Criar um novo pedido de palavra-passe','Submit Reset Request':'Enviar Pedido de Redefinição','Back to Login':'Voltar ao Início de Sessão','Current Password':'Palavra-passe Actual','New Password':'Nova Palavra-passe','Confirm Password':'Confirmar Palavra-passe','English':'Inglês','Portuguese':'Português'};function hrmsTranslate(root=document.body){
 if(window.HRMS_LANGUAGE!=='pt'||!root)return;
 const keys=Object.keys(window.HRMS_PT).sort((a,b)=>b.length-a.length);
 const cv=(value)=>{let t=value||'';keys.forEach(k=>{t=t.split(k).join(window.HRMS_PT[k])});return t};
 const walker=document.createTreeWalker(root,NodeFilter.SHOW_TEXT);const nodes=[];while(walker.nextNode())nodes.push(walker.currentNode);
 nodes.forEach(n=>{if(!n.parentElement||['SCRIPT','STYLE','TEXTAREA'].includes(n.parentElement.tagName))return;n.nodeValue=cv(n.nodeValue)});
 root.querySelectorAll('[placeholder],[title],[aria-label],input[type="button"],input[type="submit"]').forEach(el=>{
   ['placeholder','title','aria-label','value'].forEach(a=>{const v=el.getAttribute(a);if(v)el.setAttribute(a,cv(v))});
 });
 document.documentElement.lang='pt';document.title=cv(document.title);
}document.addEventListener('DOMContentLoaded',()=>hrmsTranslate());</script>
<style id="hrms-toast-style">
#hrms-toast-host{position:fixed;top:18px;right:18px;z-index:99999;width:min(420px,calc(100vw - 36px));display:flex;flex-direction:column;gap:10px;pointer-events:none}
.hrms-toast{pointer-events:auto;display:flex;align-items:flex-start;gap:12px;padding:14px 16px;border-radius:14px;color:#fff;font:700 12px/1.45 Inter,system-ui,sans-serif;box-shadow:0 18px 45px rgba(15,23,42,.24);transform:translateX(0);opacity:1;transition:opacity .3s ease,transform .3s ease;overflow-wrap:anywhere}
.hrms-toast-success{background:#059669}.hrms-toast-danger,.hrms-toast-error{background:#e11d48}.hrms-toast-warning{background:#d97706}.hrms-toast-info{background:#2563eb}
.hrms-toast-hide{opacity:0;transform:translateX(30px)}.hrms-toast-message{flex:1}.hrms-toast-close{border:0;background:transparent;color:#fff;font-size:21px;line-height:1;cursor:pointer;min-height:0;padding:0 0 0 6px}
@media(max-width:640px){ #hrms-toast-host{top:10px;right:10px;width:calc(100vw - 20px)}}
</style>
<script id="hrms-toast-script">
(function(){
  function host(){let h=document.getElementById('hrms-toast-host');if(!h){h=document.createElement('div');h.id='hrms-toast-host';h.setAttribute('aria-live','polite');h.setAttribute('aria-atomic','false');document.body.appendChild(h)}return h}
  window.showToast=function(message,type='success'){
    if(message===undefined||message===null||String(message).trim()==='')return;
    const normalized=(type==='error'?'danger':type)||'info';
    const t=document.createElement('div');t.className='hrms-toast hrms-toast-'+normalized;t.setAttribute('role',normalized==='danger'?'alert':'status');
    const m=document.createElement('div');m.className='hrms-toast-message';m.textContent=String(message);
    const x=document.createElement('button');x.type='button';x.className='hrms-toast-close';x.innerHTML='&times;';x.setAttribute('aria-label','Close notification');
    const remove=()=>{if(t.dataset.closing)return;t.dataset.closing='1';t.classList.add('hrms-toast-hide');setTimeout(()=>t.remove(),300)};
    x.addEventListener('click',remove);t.append(m,x);host().appendChild(t);setTimeout(remove,5000);
  };
  document.addEventListener('DOMContentLoaded',()=>document.querySelectorAll('.hrms-flash').forEach(el=>{showToast(el.dataset.message,el.dataset.category);el.remove()}));
})();
</script>
</body>
</html>
"""

RESET_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Reset Password | Attendance Portal</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>body { font-family: 'Inter', sans-serif; }</style>
</head>
<body class="bg-slate-900 min-h-screen flex items-center justify-center p-4 relative"><div class="absolute top-4 right-4"><select onchange="location.href='/set_language/'+this.value" class="bg-white text-slate-900 text-xs font-bold rounded-xl px-3 py-2"><option value="en" {% if session.get('ui_language','en')=='en' %}selected{% endif %}>English</option><option value="pt" {% if session.get('ui_language')=='pt' %}selected{% endif %}>Português</option></select></div>
    <div class="bg-white/95 backdrop-blur-md rounded-3xl shadow-2xl border border-slate-200/50 p-8 w-full max-w-md space-y-6">
        <div class="text-center space-y-2">
            <div class="inline-flex bg-amber-500 px-5 py-3 rounded-2xl shadow-lg mb-2 items-center justify-center">
                <span class="text-3xl">🔑</span>
            </div>
            <h1 class="text-2xl font-black text-slate-900 tracking-tight">Reset Password</h1>
            <p class="text-xs text-slate-500 font-medium">Create a new password request</p>
        </div>

        {% with messages = get_flashed_messages(with_categories=true) %}{% if messages %}{% for category, message in messages %}<div class="hrms-flash hidden" data-category="{{ category }}" data-message="{{ message|e }}"></div>{% endfor %}{% endif %}{% endwith %}

        <form method="POST" action="/reset_password" class="space-y-4">
            <div>
                <label class="block text-xs font-bold uppercase tracking-wider text-slate-500 mb-2">User ID (Employee Code / Admin)</label>
                <input type="text" name="user_id" required placeholder="NWC1234 or LM11" class="w-full bg-slate-50 border border-slate-300 rounded-xl px-4 py-3 text-sm font-medium focus:ring-2 focus:ring-amber-500 focus:outline-none">
            </div>
            <div>
                <label class="block text-xs font-bold uppercase tracking-wider text-slate-500 mb-2">New Password</label>
                <input type="password" name="new_password" required placeholder="Enter new strong password" class="w-full bg-slate-50 border border-slate-300 rounded-xl px-4 py-3 text-sm font-medium focus:ring-2 focus:ring-amber-500 focus:outline-none">
            </div>
            <button type="submit" class="w-full bg-amber-500 hover:bg-amber-600 text-white font-bold text-xs uppercase tracking-widest py-3.5 rounded-xl shadow-lg transition duration-200">
                Send Request for Approval 🚀
            </button>
            <div class="text-center mt-3">
                <a href="/login" class="text-[11px] text-slate-500 font-bold hover:text-slate-700 hover:underline transition">← Back to Login</a>
            </div>
        </form>
    </div>
<script>window.HRMS_LANGUAGE={{ session.get('ui_language','en')|tojson }};window.HRMS_PT={'Correction Rejected':'Correcção Rejeitada','Correction Approved':'Correcção Aprovada','Correction Request':'Pedido de Correcção','New Document':'Novo Documento','was Rejected':'foi rejeitada','was Approved':'foi aprovada','Your correction':'A sua correcção','Your roster':'A sua escala','Month select karke Load Month click karein':'Seleccione o mês e clique em Carregar Mês','Employee-wise Shift A, Shift B and Weekly Off allocation':'Alocação por trabalhador de Turno A, Turno B e Folga Semanal','Public Holiday':'Feriado Público','Half Day':'Meio Dia','Full Day':'Dia Completo','Full Month':'Mês Completo','Clear All':'Limpar Tudo','Select All':'Seleccionar Tudo','Uploaded By':'Carregado por','Updated By':'Actualizado por','Created On':'Criado em','Created By':'Criado por','Records':'Registos','Record':'Registo','Summary':'Resumo','Details':'Detalhes','Code':'Código','Address':'Endereço','Phone':'Telefone','Email':'E-mail','Photo':'Fotografia','Store Access':'Acesso à Loja','Apply Selected':'Aplicar Seleccionados','Copy Previous Month':'Copiar Mês Anterior','Copy Previous Week':'Copiar Semana Anterior','Monthly roster saved':'Escala mensal guardada','Kam se kam ek option select karein.':'Seleccione pelo menos uma opção.','Month select karein':'Seleccione o mês','Galat User ID ya Password!':'ID do utilizador ou palavra-passe incorrectos!','Sirf PDF files allowed hain!':'Apenas ficheiros PDF são permitidos!','Sabhi fields bharna zaroori hai!':'Todos os campos são obrigatórios!','This roster shows only your shifts and weekly-off dates for the current month.':'Esta escala mostra apenas os seus turnos e folgas semanais do mês actual.','Overtime Minutes':'Minutos de Horas Extra','Submit Overtime':'Enviar Horas Extra','Email - requires SMTP':'E-mail - requer SMTP','Internal Notification':'Notificação Interna','Download Center':'Centro de Transferências','Monthly':'Mensal','Weekly':'Semanal','Daily':'Diário','Delivery':'Entrega','Frequency':'Frequência','File':'Ficheiro','Uploaded At':'Carregado em','Created':'Criado','To':'Para','From':'De','Effective':'Efectivo','Estimate':'Estimativa','Employer':'Empregador','Gross':'Bruto','Configurable estimate using 2026 settings. Validate payroll categories, taxable base and exemptions before finalisation.':'Estimativa configurável com definições de 2026. Valide as categorias salariais, a base tributável e as isenções antes da finalização.','Identification masking':'Ocultação da Identificação','Document access log':'Registo de Acesso a Documentos','Access audit':'Auditoria de Acesso','Last backups':'Últimas Cópias de Segurança','Number of Users':'Número de Utilizadores','Application Version':'Versão da Aplicação','retry required':'nova tentativa necessária','Last saved attendance remains available':'A última assiduidade guardada continua disponível','Queue State':'Estado da Fila','Level':'Nível','Late Count':'Número de Atrasos','Rolling 30-day view.':'Vista móvel de 30 dias.','roster cells copied':'células de escala copiadas','assignments saved':'alocações guardadas','Copy failed':'Falha ao copiar','Save Lifecycle Event':'Guardar Evento do Ciclo de Vida','Employee Lifecycle':'Ciclo de Vida do Trabalhador','New Records':'Novos Registos','Records Received':'Registos Recebidos','Machine Status':'Estado do Equipamento','Store Code':'Código da Loja','Payroll, holidays, lifecycle, overtime, backups and privacy controls.':'Salários, feriados, ciclo de vida, horas extra, cópias de segurança e controlos de privacidade.','Privacy-safe QR verification cards.':'Cartões QR de verificação com protecção de privacidade.','Internal leave, shift, payroll and contract notices.':'Avisos internos de ausências, turnos, salários e contratos.','Track user, store, action, IP and changes.':'Acompanhar utilizador, loja, acção, IP e alterações.','Minimum staffing rules and shortage warnings.':'Regras de dotação mínima e avisos de falta de pessoal.','Copy week or month schedules.':'Copiar escalas semanais ou mensais.','Contracts, IDs and certificates.':'Contratos, documentos de identificação e certificados.','Annual, used, pending and available balances.':'Saldos anuais, usados, pendentes e disponíveis.','Approve missing-punch correction requests.':'Aprovar pedidos de correcção de marcações em falta.','Machine status, manual checks and sync history.':'Estado dos equipamentos, verificações manuais e histórico de sincronização.','Combined attendance with store code and Excel export.':'Assiduidade combinada com código da loja e exportação Excel.','Multi-store operations, compliance and employee self-service':'Operações multi-loja, conformidade e auto-serviço do trabalhador','No contract alerts.':'Sem alertas de contrato.','Audit Events':'Eventos de Auditoria','Contracts':'Contratos','Corrections':'Correcções','Machine':'Equipamento','Store Comparison':'Comparação de Lojas','Monthly Summary':'Resumo Mensal','Default Schedule':'Horário Padrão','All Allocations':'Todas as Alocações','Designations Selected':'Funções Seleccionadas','All Designations Selected':'Todas as Funções Seleccionadas','Same Weekday for Full Month':'Mesmo Dia da Semana durante Todo o Mês','Monday-Sunday':'Segunda a Domingo','Last Week':'Última Semana','5th Week':'5.ª Semana','4th Week':'4.ª Semana','3rd Week':'3.ª Semana','2nd Week':'2.ª Semana','1st Week':'1.ª Semana','All Weeks (Monday-Sunday)':'Todas as Semanas (Segunda a Domingo)','Not in this month':'Não existe neste mês','options selected':'opções seleccionadas','option selected':'opção seleccionada','Apply to selected date':'Aplicar à data seleccionada','Month & Year':'Mês e Ano','load month':'carregar mês','Select Month':'Seleccionar Mês','No shift requests':'Sem pedidos de turno','Current Assignment':'Alocação Actual','Requested':'Solicitado','Shift Request':'Pedido de Turno','Request Shift / Off':'Pedir Turno / Folga','Select Employee':'Seleccionar Trabalhador','Bulk Upload':'Carregamento em Massa','Standard Row':'Linha Padrão','Monthly grid':'Grelha mensal','Master':'Principal','Individual Slip':'Recibo Individual','Split':'Separar','All Personnel':'Todo o Pessoal','No requests':'Sem pedidos','Minutes':'Minutos','ALL':'TODOS','Change':'Alteração','Emp':'Trab.','Remove Photo':'Remover Fotografia','STORE ACCESS':'ACESSO À LOJA','Create login and employee profile':'Criar início de sessão e perfil do trabalhador','No employees found.':'Nenhum trabalhador encontrado.','Select Designation':'Seleccionar Função','Add Employee':'Adicionar Trabalhador','Search, review and maintain workforce profiles':'Pesquisar, rever e manter perfis dos trabalhadores','Open Roster Planner':'Abrir Planeador de Escala','Use Roster Planner to assign shifts and weekly-off rotations.':'Utilize o Planeador de Escala para atribuir turnos e rotações de folga semanal.','Rota Rotation and Shift Schedules are integrated with the':'A rotação da escala e os horários de turno estão integrados com o','No salary slips uploaded yet.':'Ainda não foram carregados recibos de salário.','View':'Ver','PDF File':'Ficheiro PDF','-- Select Employee --':'-- Seleccionar Trabalhador --','Individual Slip Upload':'Carregar Recibo Individual','Split & Upload':'Separar e Carregar','Master PDF':'PDF Principal','Bulk Upload (Merged PDF)':'Carregamento em Massa (PDF Unido)','Monthly grid with Leave Codes':'Grelha mensal com códigos de ausência','Employee Matrix':'Matriz de Trabalhadores','Individual records per date':'Registos individuais por data','Standard Row Export':'Exportação em Linhas','Download Attendance Reports':'Baixar Relatórios de Assiduidade','Total Summary':'Resumo Total','No attendance records found for this selection.':'Nenhum registo de assiduidade encontrado para esta selecção.','Status ↕':'Estado ↕','Working Hrs ↕':'Horas Trabalhadas ↕','Total Lunch ↕':'Total de Almoço ↕','Out Time ↕':'Hora de Saída ↕','Store In ↕':'Entrada na Loja ↕','Dept ↕':'Departamento ↕','Employee Name ↕':'Nome do Trabalhador ↕','Date ↕':'Data ↕','Last Month':'Mês Passado','This Month':'Este Mês','This Week':'Esta Semana','Yesterday':'Ontem','Quick Range':'Intervalo Rápido','Rota':'Escala','Export':'Exportar','Logged In As':'Sessão Iniciada Como','-- All Personnel --':'-- Todo o Pessoal --','Employee Filter':'Filtro de Trabalhadores','Tot Hrs':'Total de Horas','Mis Punch':'Marcação em Falta','Mis-Punch':'Marcação em Falta','Late Arr.':'Chegada Tardia','Week Off':'Folga Semanal','Shutdown':'Encerrar','Sync':'Sincronização','Device':'Equipamento','Leaves':'Ausências','Dev':'Programador','Biometric live tracking active for Attendance Portal.':'Acompanhamento biométrico em tempo real activo no Portal de Assiduidade.','Announcements':'Comunicados','Employees Info (ID Card)':'Informações dos Trabalhadores (Cartão de Identificação)','Manage Passwords':'Gerir Palavras-passe','Biometric Machines':'Equipamentos Biométricos','Shift Approvals':'Aprovações de Turno','Reset':'Repor','Off':'Folga','Daily manpower summary':'Resumo diário de efectivos','Loading monthly roster...':'A carregar a escala mensal...','Send Request for Approval':'Enviar Pedido para Aprovação','Developed by':'Desenvolvido por','Attendance Portal':'Portal de Assiduidade','Sign in to access your dashboard':'Inicie sessão para aceder ao seu painel','User ID':'ID do Utilizador','Password':'Palavra-passe','Secure Login':'Iniciar Sessão','Forgot/Reset Password?':'Esqueceu/Redefinir Palavra-passe?','Reset Password':'Redefinir Palavra-passe','Create a new password request':'Criar um novo pedido de palavra-passe','Submit Reset Request':'Enviar Pedido de Redefinição','Back to Login':'Voltar ao Início de Sessão','Current Password':'Palavra-passe Actual','New Password':'Nova Palavra-passe','Confirm Password':'Confirmar Palavra-passe','English':'Inglês','Portuguese':'Português'};function hrmsTranslate(root=document.body){
 if(window.HRMS_LANGUAGE!=='pt'||!root)return;
 const keys=Object.keys(window.HRMS_PT).sort((a,b)=>b.length-a.length);
 const cv=(value)=>{let t=value||'';keys.forEach(k=>{t=t.split(k).join(window.HRMS_PT[k])});return t};
 const walker=document.createTreeWalker(root,NodeFilter.SHOW_TEXT);const nodes=[];while(walker.nextNode())nodes.push(walker.currentNode);
 nodes.forEach(n=>{if(!n.parentElement||['SCRIPT','STYLE','TEXTAREA'].includes(n.parentElement.tagName))return;n.nodeValue=cv(n.nodeValue)});
 root.querySelectorAll('[placeholder],[title],[aria-label],input[type="button"],input[type="submit"]').forEach(el=>{
   ['placeholder','title','aria-label','value'].forEach(a=>{const v=el.getAttribute(a);if(v)el.setAttribute(a,cv(v))});
 });
 document.documentElement.lang='pt';document.title=cv(document.title);
}document.addEventListener('DOMContentLoaded',()=>hrmsTranslate());</script>
<style id="hrms-toast-style">
#hrms-toast-host{position:fixed;top:18px;right:18px;z-index:99999;width:min(420px,calc(100vw - 36px));display:flex;flex-direction:column;gap:10px;pointer-events:none}
.hrms-toast{pointer-events:auto;display:flex;align-items:flex-start;gap:12px;padding:14px 16px;border-radius:14px;color:#fff;font:700 12px/1.45 Inter,system-ui,sans-serif;box-shadow:0 18px 45px rgba(15,23,42,.24);transform:translateX(0);opacity:1;transition:opacity .3s ease,transform .3s ease;overflow-wrap:anywhere}
.hrms-toast-success{background:#059669}.hrms-toast-danger,.hrms-toast-error{background:#e11d48}.hrms-toast-warning{background:#d97706}.hrms-toast-info{background:#2563eb}
.hrms-toast-hide{opacity:0;transform:translateX(30px)}.hrms-toast-message{flex:1}.hrms-toast-close{border:0;background:transparent;color:#fff;font-size:21px;line-height:1;cursor:pointer;min-height:0;padding:0 0 0 6px}
@media(max-width:640px){ #hrms-toast-host{top:10px;right:10px;width:calc(100vw - 20px)}}
</style>
<script id="hrms-toast-script">
(function(){
  function host(){let h=document.getElementById('hrms-toast-host');if(!h){h=document.createElement('div');h.id='hrms-toast-host';h.setAttribute('aria-live','polite');h.setAttribute('aria-atomic','false');document.body.appendChild(h)}return h}
  window.showToast=function(message,type='success'){
    if(message===undefined||message===null||String(message).trim()==='')return;
    const normalized=(type==='error'?'danger':type)||'info';
    const t=document.createElement('div');t.className='hrms-toast hrms-toast-'+normalized;t.setAttribute('role',normalized==='danger'?'alert':'status');
    const m=document.createElement('div');m.className='hrms-toast-message';m.textContent=String(message);
    const x=document.createElement('button');x.type='button';x.className='hrms-toast-close';x.innerHTML='&times;';x.setAttribute('aria-label','Close notification');
    const remove=()=>{if(t.dataset.closing)return;t.dataset.closing='1';t.classList.add('hrms-toast-hide');setTimeout(()=>t.remove(),300)};
    x.addEventListener('click',remove);t.append(m,x);host().appendChild(t);setTimeout(remove,5000);
  };
  document.addEventListener('DOMContentLoaded',()=>document.querySelectorAll('.hrms-flash').forEach(el=>{showToast(el.dataset.message,el.dataset.category);el.remove()}));
})();
</script>
</body>
</html>
"""

@app.context_processor
def inject_portal_helpers(): return {'session_has_permission': session_has_permission}

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="{{ ui_language }}">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{{ portal_name }} | Attendance Portal</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        body { font-family: 'Inter', sans-serif; }
        td.nowrap-cell { white-space: nowrap; }
        .excel-table th, .excel-table td { border: 1px solid #e2e8f0 !important; white-space: nowrap; }
        th.sortable { cursor: pointer; user-select: none; transition: background-color 0.15s ease; }
        th.sortable:hover { background-color: #cbd5e1; }
        .stat-card { cursor: pointer; transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1); position: relative; overflow: hidden; }
        .stat-card:hover { transform: translateY(-2px); box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.05); }
        
        /* Custom scrollbar for better mobile view */
        ::-webkit-scrollbar { width: 6px; height: 6px; }
        ::-webkit-scrollbar-track { background: #f1f5f9; rounded-full }
        ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 10px; }
        ::-webkit-scrollbar-thumb:hover { background: #94a3b8; }
    </style>
    <script>
        let inactivityTimer;
        const INACTIVITY_LIMIT = 10 * 60 * 1000;

        function resetInactivityTimer() {
            clearTimeout(inactivityTimer);
            inactivityTimer = setTimeout(() => {
                showToast("Session timeout ho gaya hai. Dobara login karein.", "warning");
                window.location.href = "/logout";
            }, INACTIVITY_LIMIT);
        }

        window.onload = function() {
            const events = ['mousemove', 'keypress', 'click', 'scroll', 'touchstart'];
            events.forEach(eventName => {
                document.addEventListener(eventName, resetInactivityTimer, true);
            });
            resetInactivityTimer();
        };

        function updateLiveClock() {
            const now = new Date();
            const options = { weekday: 'short', year: 'numeric', month: 'short', day: 'numeric' };
            const dateStr = now.toLocaleDateString('en-US', options);
            let hours = String(now.getHours()).padStart(2, '0');
            let minutes = String(now.getMinutes()).padStart(2, '0');
            let seconds = String(now.getSeconds()).padStart(2, '0');
            const clockEl = document.getElementById('live-digital-clock');
            if (clockEl) clockEl.innerText = dateStr + ' | ' + hours + ':' + minutes + ':' + seconds;
        }
        
        function toggleMobileMenu(show) {
            const sidebar = document.getElementById('sidebar');
            const overlay = document.getElementById('sidebar-overlay');
            if (show) {
                sidebar.classList.remove('-translate-x-full');
                overlay.classList.remove('hidden');
            } else {
                sidebar.classList.add('-translate-x-full');
                overlay.classList.add('hidden');
            }
        }

        let sortDirections = {};
        function sortTable(columnIndex, isNumeric = false) {
            const table = document.getElementById("attendance-table");
            if (!table) return;
            const tbody = table.tBodies[0];
            const rows = Array.from(tbody.querySelectorAll("tr"));
            if (rows.length <= 1 && rows[0].cells.length <= 1) return;

            let dir = sortDirections[columnIndex] || 'asc';
            sortDirections[columnIndex] = (dir === 'asc') ? 'desc' : 'asc';

            rows.sort((rowA, rowB) => {
                let cellA = rowA.cells[columnIndex].innerText.trim();
                let cellB = rowB.cells[columnIndex].innerText.trim();
                if (isNumeric) {
                    let valA = parseFloat(cellA.replace(/[^0-9.-]+/g,"")) || 0;
                    let valB = parseFloat(cellB.replace(/[^0-9.-]+/g,"")) || 0;
                    return (dir === 'asc') ? valA - valB : valB - valA;
                } else {
                    return (dir === 'asc') ? cellA.localeCompare(cellB) : cellB.localeCompare(cellA);
                }
            });

            tbody.innerHTML = "";
            rows.forEach((row, index) => {
                if (row.cells[0]) row.cells[0].innerText = index + 1;
                tbody.appendChild(row);
            });
        }

        function filterByStatus(statusVal) {
            let table = document.getElementById('attendance-table');
            if (!table) return;
            let trs = table.tBodies[0].getElementsByTagName('tr');
            for (let i = 0; i < trs.length; i++) {
                if (statusVal === 'ALL') {
                    trs[i].style.display = "";
                } else if (statusVal === 'Late Arrival') {
                    trs[i].style.display = (trs[i].getAttribute('data-late') === 'Yes') ? "" : "none";
                } else if (statusVal === 'Shift A' || statusVal === 'Shift B') {
                    trs[i].style.display = (trs[i].getAttribute('data-shift') === statusVal) ? "" : "none";
                } else if (statusVal === 'ML') {
                    let statusCell = trs[i].getElementsByTagName('td')[12];
                    if (statusCell) {
                        let text = statusCell.textContent || statusCell.innerText;
                        trs[i].style.display = (text.includes('F01;1') || text.includes('F05;1') || text.includes('F10;1') || text.includes('F51;1') || text.includes('F60;1') || text.includes('F61;1') || text.includes('F62;1') || text.includes('Leave')) ? "" : "none";
                    }
                } else {
                    let statusCell = trs[i].getElementsByTagName('td')[12];
                    if (statusCell) {
                        trs[i].style.display = (statusCell.textContent || statusCell.innerText).includes(statusVal) ? "" : "none";
                    }
                }
            }
        }

        function filterTableSearch() {
            let input = document.getElementById('table-search-input').value.toLowerCase();
            let table = document.getElementById('attendance-table');
            if (!table) return;
            let trs = table.tBodies[0].getElementsByTagName('tr');
            for (let i = 0; i < trs.length; i++) {
                let idCell = trs[i].getElementsByTagName('td')[2];
                let nameCell = trs[i].getElementsByTagName('td')[3];
                let deptCell = trs[i].getElementsByTagName('td')[4];
                if (idCell && nameCell && deptCell) {
                    let text = (idCell.textContent + ' ' + nameCell.textContent + ' ' + deptCell.textContent).toLowerCase();
                    trs[i].style.display = (text.indexOf(input) > -1) ? "" : "none";
                }
            }
        }

        let timeLeft = 900;
        let timerInterval;
        function startTimer() {
            clearInterval(timerInterval);
            timeLeft = 900;
            timerInterval = setInterval(function() {
                if (timeLeft <= 0) { window.location.reload(); } else {
                    let m = Math.floor(timeLeft / 60);
                    let s = timeLeft % 60;
                    let timerEl = document.getElementById('countdown-timer');
                    if (timerEl) { timerEl.innerText = m + ':' + (s < 10 ? '0' : '') + s; }
                    timeLeft -= 1;
                }
            }, 1000);
        }

        function setQuickDate(type) {
            let today = new Date();
            let startInput = document.querySelector('input[name="start_date"]');
            let endInput = document.querySelector('input[name="end_date"]');
            let formatDate = (d) => {
                let month = '' + (d.getMonth() + 1), day = '' + d.getDate(), year = d.getFullYear();
                if (month.length < 2) month = '0' + month;
                if (day.length < 2) day = '0' + day;
                return [year, month, day].join('-');
            };

            if (type === 'today') { startInput.value = formatDate(today); endInput.value = formatDate(today); }
            else if (type === 'yesterday') { let yest = new Date(); yest.setDate(today.getDate() - 1); startInput.value = formatDate(yest); endInput.value = formatDate(yest); }
            else if (type === 'week') { let firstDay = new Date(today.setDate(today.getDate() - today.getDay())); startInput.value = formatDate(firstDay); endInput.value = formatDate(new Date()); }
            else if (type === 'month') { let firstDay = new Date(today.getFullYear(), today.getMonth(), 1); startInput.value = formatDate(firstDay); endInput.value = formatDate(new Date()); }
            else if (type === 'last_month') { let firstDay = new Date(today.getFullYear(), today.getMonth() - 1, 1); let lastDay = new Date(today.getFullYear(), today.getMonth(), 0); startInput.value = formatDate(firstDay); endInput.value = formatDate(lastDay); }
            
            document.getElementById('filter-form').submit();
        }

        function toggleModal(id, show) {
            let el = document.getElementById(id);
            if(el) { show ? el.classList.remove('hidden') : el.classList.add('hidden'); }
            if (show) toggleMobileMenu(false);
        }

        function inlineEdit(date, empId, field, currVal) {
            let title = field.replace('_', ' ').toUpperCase();
            let promptVal = currVal !== '-' ? currVal : '';
            let newVal = prompt(`Direct Editing: ${empId} | ${date}\\n\\nEnter new time for ${title} (HH:MM:SS) or leave blank to clear the overridden value:`, promptVal);
            if (newVal !== null) {
                window.location.href = `/quick_edit?date=${date}&emp_id=${empId}&field=${field}&time=${encodeURIComponent(newVal)}`;
            }
        }

        function filterOvertimeHistory(){const q=(document.getElementById('overtime-history-search')?.value||'').toLowerCase();const st=document.getElementById('overtime-history-status')?.value||'';document.querySelectorAll('.overtime-history-row').forEach(r=>r.style.display=(r.dataset.search.includes(q)&&(!st||r.dataset.status===st))?'':'none');}
        function setRights(btn,on){btn.closest('form').querySelectorAll('.rights-grid input[type=checkbox]').forEach(x=>x.checked=on);}
        function secureShutdown() {
            let pwd = prompt("Server band karne ke liye password enter karein:");
            if (pwd) window.location.href = "/shutdown?pwd=" + encodeURIComponent(pwd);
        }

        let rosterMatrixData = null;
        function rosterValueClass(value) {
            if (value === 'Shift A') return 'bg-emerald-50 text-emerald-700 border-emerald-200';
            if (value === 'Shift B') return 'bg-indigo-50 text-indigo-700 border-indigo-200';
            if (value === 'Weekly Off') return 'bg-rose-50 text-rose-700 border-rose-200';
            return 'bg-white text-slate-600 border-slate-200';
        }
        function loadRosterPlanner() {
            const monthStr = document.getElementById('rp-month').value;
            const table = document.getElementById('roster-matrix-table');
            if (!monthStr) { showToast('Month select karein','warning'); return; }
            table.innerHTML = '<div class="py-12 text-center text-slate-400">Loading monthly roster...</div>';
            fetch(`/api/roster_matrix?month=${encodeURIComponent(monthStr)}`)
            .then(r => { if(!r.ok) throw new Error('Roster load failed'); return r.json(); })
            .then(data => { rosterMatrixData=data; renderRosterMatrix(data); document.getElementById('save-roster-btn').classList.remove('hidden'); })
            .catch(err => table.innerHTML=`<div class="py-12 text-center text-rose-600">${err.message}</div>`);
        }
        function renderRosterMatrix(data) {
            const host=document.getElementById('roster-matrix-table');
            const dates=data.dates || [], employees=data.employees || [];
            let html='<table class="min-w-max w-full text-[10px] border-collapse"><thead class="sticky top-0 z-20">';
            html+='<tr class="bg-slate-900 text-white"><th class="sticky left-0 z-30 bg-slate-900 p-2 border border-slate-700 w-12">Sr.</th><th class="sticky left-12 z-30 bg-slate-900 p-2 border border-slate-700 min-w-[105px]">Employee Code</th><th class="sticky left-[153px] z-30 bg-slate-900 p-2 border border-slate-700 min-w-[190px]">Employee Name</th><th class="sticky left-[343px] z-30 bg-slate-900 p-2 border border-slate-700 min-w-[150px]">Designation</th>';
            dates.forEach(d=>{html+=`<th class="p-2 border border-slate-700 min-w-[102px] text-center"><div class="text-sm font-black">${d.day}</div><div class="text-[9px] text-cyan-200">${d.weekday}</div></th>`});
            html+='</tr><tr class="bg-slate-100 text-slate-700"><th colspan="4" class="sticky left-0 z-30 bg-slate-100 p-2 border border-slate-300 text-left">Daily manpower summary</th>';
            dates.forEach(d=>{const c=data.summary[d.date]||{};html+=`<th class="p-1 border border-slate-300"><div class="grid gap-1 text-[9px]"><span id="sum-a-${d.date}" class="rounded bg-emerald-100 text-emerald-800 px-1">A: ${c.shift_a||0}</span><span id="sum-b-${d.date}" class="rounded bg-indigo-100 text-indigo-800 px-1">B: ${c.shift_b||0}</span><span id="sum-off-${d.date}" class="rounded bg-rose-100 text-rose-800 px-1">Off: ${c.weekly_off||0}</span></div></th>`});
            html+='</tr></thead><tbody>';
            employees.forEach((e,i)=>{html+=`<tr class="roster-employee-row hover:bg-slate-50" data-code="${e.user_id.toLowerCase()}" data-name="${e.name.toLowerCase()}" data-designation="${(e.designation||e.dept||'employee').toLowerCase()}"><td class="sticky left-0 z-10 bg-white p-2 border border-slate-200 text-center font-bold">${i+1}</td><td class="sticky left-12 z-10 bg-white p-2 border border-slate-200 font-mono font-bold">${e.user_id}</td><td class="sticky left-[153px] z-10 bg-white p-2 border border-slate-200 font-bold"><div class="flex items-center justify-between gap-2"><span class="truncate" title="${e.name}">${e.name}</span><button type="button" onclick="resetEmployeeRoster('${e.user_id}')" class="shrink-0 rounded-md border border-rose-200 bg-rose-50 px-2 py-1 text-[8px] font-black text-rose-700 hover:bg-rose-100" title="Reset this employee's selected month roster">Reset</button></div></td><td class="sticky left-[343px] z-10 bg-white p-2 border border-slate-200">${e.designation||e.dept||'Employee'}</td>`;
                dates.forEach(d=>{const v=(e.roster&&e.roster[d.date])||'';html+=`<td class="p-1 border border-slate-200"><select data-emp="${e.user_id}" data-date="${d.date}" onchange="onRosterCellChange(this)" class="roster-matrix-select w-full rounded-lg border px-1 py-2 text-[9px] font-bold ${rosterValueClass(v)}"><option value="" ${!v?'selected':''}>Default</option><option value="Shift A" ${v==='Shift A'?'selected':''}>Shift A</option><option value="Shift B" ${v==='Shift B'?'selected':''}>Shift B</option><option value="Weekly Off" ${v==='Weekly Off'?'selected':''}>Weekly Off</option></select></td>`});
                html+='</tr>'});
            html+='</tbody></table>'; host.innerHTML=html; populateRosterDesignationFilter(); filterRosterRows();
        }
        let pendingRosterChange = null;
        function onRosterCellChange(el) {
            if(!rosterMatrixData) return;
            const employee=rosterMatrixData.employees.find(e=>e.user_id===el.dataset.emp);
            if(!employee) return;
            const oldValue=(employee.roster&&employee.roster[el.dataset.date])||'';
            const newValue=el.value;
            if(!newValue) {
                employee.roster[el.dataset.date]='';
                el.className='roster-matrix-select w-full rounded-lg border px-1 py-2 text-[9px] font-bold '+rosterValueClass('');
                updateRosterSummary(); filterRosterRows(); return;
            }
            pendingRosterChange={empId:el.dataset.emp,date:el.dataset.date,value:newValue,oldValue,element:el};
            const empName=employee.name||employee.user_id;
            document.getElementById('week-apply-title').textContent=`${empName} • ${newValue}`;
            document.getElementById('week-apply-subtitle').textContent=`Selected date: ${el.dataset.date}. Same ${newValue} kin week ranges me apply karna hai? Multiple options select kar sakte hain.`;
            document.querySelectorAll('input[name="week-apply-option"]').forEach(x=>x.checked=false);
            document.querySelector('input[name="week-apply-option"][value="selected_date"]').checked=true;
            updateWeekSelectionCount();
            toggleModal('week-apply-modal',true);
        }
        function cancelRosterWeekApply() {
            if(pendingRosterChange && pendingRosterChange.element) pendingRosterChange.element.value=pendingRosterChange.oldValue||'';
            pendingRosterChange=null; toggleModal('week-apply-modal',false);
        }
        function getRosterWeekDates(mode, selectedDate) {
            if(!rosterMatrixData) return [];
            const dates=rosterMatrixData.dates.map(d=>d.date);
            const selected=new Date(selectedDate+'T00:00:00');
            if(mode==='selected_date') return [selectedDate];
            if(mode==='same_weekday_all') {
                const weekday=selected.getDay(); return dates.filter(dt=>new Date(dt+'T00:00:00').getDay()===weekday);
            }
            let startDay=1,endDay=7;
            if(mode==='week_2'){startDay=8;endDay=14}
            else if(mode==='week_3'){startDay=15;endDay=21}
            else if(mode==='week_4'){startDay=22;endDay=28}
            else if(mode==='last_week'){startDay=29;endDay=99}
            return dates.filter(dt=>{const day=parseInt(dt.slice(-2));return day>=startDay&&day<=endDay});
        }
        function handleWeekOptionChange(changed) {
            const selectedDate=document.querySelector('input[name="week-apply-option"][value="selected_date"]');
            const sameWeekday=document.querySelector('input[name="week-apply-option"][value="same_weekday_all"]');
            const weekBoxes=[...document.querySelectorAll('input[name="week-apply-option"]')].filter(x=>!['selected_date','same_weekday_all'].includes(x.value));
            if(changed.value==='selected_date' && changed.checked) {
                weekBoxes.forEach(x=>x.checked=false); sameWeekday.checked=false;
            } else if(changed.value==='same_weekday_all' && changed.checked) {
                weekBoxes.forEach(x=>x.checked=false); selectedDate.checked=false;
            } else if(changed.checked) {
                selectedDate.checked=false; sameWeekday.checked=false;
            }
            if(!document.querySelector('input[name="week-apply-option"]:checked')) selectedDate.checked=true;
            updateWeekSelectionCount();
        }
        function updateWeekSelectionCount() {
            const selected=[...document.querySelectorAll('input[name="week-apply-option"]:checked')];
            const label=document.getElementById('week-selection-count');
            if(label) label.textContent=selected.length===1?'1 option selected':`${selected.length} options selected`;
        }
        function toggleAllRosterWeeks(master) {
            const values=['week_1','week_2','week_3','week_4','last_week'];
            document.querySelectorAll('input[name="week-apply-option"]').forEach(x=>x.checked=values.includes(x.value)?master.checked:false);
            updateWeekSelectionCount();
        }
        function applyRosterWeekChoice() {
            if(!pendingRosterChange) return;
            const modes=[...document.querySelectorAll('input[name="week-apply-option"]:checked')].map(x=>x.value);
            if(!modes.length){showToast('Kam se kam ek option select karein.','warning');return;}
            const employee=rosterMatrixData.employees.find(e=>e.user_id===pendingRosterChange.empId);
            if(!employee) return;
            const dateSet=new Set();
            modes.forEach(mode=>getRosterWeekDates(mode,pendingRosterChange.date).forEach(dt=>dateSet.add(dt)));
            dateSet.forEach(dt=>employee.roster[dt]=pendingRosterChange.value);
            pendingRosterChange=null; toggleModal('week-apply-modal',false); renderRosterMatrix(rosterMatrixData); updateRosterSummary(); filterRosterRows();
        }
        function updateRosterSummary() {
            if(!rosterMatrixData) return;
            rosterMatrixData.dates.forEach(d=>{rosterMatrixData.summary[d.date]={shift_a:0,shift_b:0,weekly_off:0}});
            rosterMatrixData.employees.forEach(e=>Object.entries(e.roster||{}).forEach(([dt,val])=>{
                const c=rosterMatrixData.summary[dt]; if(!c) return;
                if(val==='Shift A')c.shift_a++; else if(val==='Shift B')c.shift_b++; else if(val==='Weekly Off')c.weekly_off++;
            }));
            rosterMatrixData.dates.forEach(d=>{
                const c=rosterMatrixData.summary[d.date]||{};
                const a=document.getElementById('sum-a-'+d.date), b=document.getElementById('sum-b-'+d.date), o=document.getElementById('sum-off-'+d.date);
                if(a)a.textContent='A: '+(c.shift_a||0); if(b)b.textContent='B: '+(c.shift_b||0); if(o)o.textContent='Off: '+(c.weekly_off||0);
            });
        }
        function populateRosterDesignationFilter() {
            const menu=document.getElementById('roster-designation-menu'); if(!menu || !rosterMatrixData) return;
            const selected=new Set([...menu.querySelectorAll('input:checked')].map(x=>x.value));
            const values=[...new Set(rosterMatrixData.employees.map(e=>e.designation||e.dept||'Employee'))].sort();
            menu.innerHTML='<label class="flex gap-2 p-2 text-xs font-bold border-b"><input type="checkbox" id="designation-all" onchange="clearDesignationFilters()"> All Designations</label>'+values.map(v=>`<label class="flex gap-2 p-2 text-xs hover:bg-slate-50 rounded"><input class="roster-designation-check" type="checkbox" value="${v.toLowerCase()}" ${selected.has(v.toLowerCase())?'checked':''} onchange="filterRosterRows();updateDesignationButton()"> ${v}</label>`).join('');
            updateDesignationButton();
        }
        function toggleDesignationMenu(){document.getElementById('roster-designation-menu').classList.toggle('hidden')}
        document.addEventListener('click', function(event) {
            const menu=document.getElementById('roster-designation-menu');
            const button=document.getElementById('roster-designation-button');
            if(!menu || !button || menu.classList.contains('hidden')) return;
            if(!menu.contains(event.target) && !button.contains(event.target)) {
                menu.classList.add('hidden');
                filterRosterRows();
            }
        });
        function selectedRosterDesignations(){return [...document.querySelectorAll('.roster-designation-check:checked')].map(x=>x.value)}
        function clearDesignationFilters(){document.querySelectorAll('.roster-designation-check').forEach(x=>x.checked=false);filterRosterRows();updateDesignationButton()}
        function updateDesignationButton(){const values=selectedRosterDesignations();const label=document.querySelector('#roster-designation-button span');if(label)label.textContent=values.length?`${values.length} Designations Selected`:'All Designations'}
        function resetEmployeeRoster(empId) {
            if(!rosterMatrixData) return;
            const employee=rosterMatrixData.employees.find(e=>e.user_id===empId);
            if(!employee) return;
            const monthLabel=document.getElementById('rp-month')?.value||'selected month';
            if(!confirm(`${employee.name} ka ${monthLabel} ka complete roster Default par reset karna hai?`)) return;
            rosterMatrixData.dates.forEach(d=>employee.roster[d.date]='');
            renderRosterMatrix(rosterMatrixData);
            updateRosterSummary();
            filterRosterRows();
            const notice=document.getElementById('roster-unsaved-notice');
            if(notice){notice.classList.remove('hidden');notice.textContent=`${employee.name} ka roster reset hua hai. Changes permanently save karne ke liye Save Monthly Roster click karein.`;}
        }
        function filterRosterRows() {
            const q=(document.getElementById('roster-search')?.value||'').trim().toLowerCase();
            const designations=selectedRosterDesignations();
            const allocation=document.getElementById('roster-allocation-filter')?.value||'';
            updateDesignationButton();
            document.querySelectorAll('.roster-employee-row').forEach(row=>{
                const matchesText=!q || row.dataset.code.includes(q) || row.dataset.name.includes(q);
                const matchesDesignation=!designations.length || designations.includes(row.dataset.designation);
                let matchesAllocation=true;
                if(allocation) matchesAllocation=[...row.querySelectorAll('.roster-matrix-select')].some(x=>x.value===allocation);
                row.style.display=matchesText && matchesDesignation && matchesAllocation ? '' : 'none';
            });
        }
        function resetRosterFilters() {
            document.getElementById('roster-search').value='';
            document.querySelectorAll('.roster-designation-check').forEach(x=>x.checked=false); updateDesignationButton();
            document.getElementById('roster-allocation-filter').value='';
            filterRosterRows();
        }
        function exportRosterExcel() {
            const month=document.getElementById('rp-month').value;
            if(!month){showToast('Month select karein','warning');return;}
            window.location.href='/export_roster_excel?month='+encodeURIComponent(month);
        }
        function saveRoster() {
            const updates=[];
            document.querySelectorAll('.roster-matrix-select').forEach(x=>updates.push({emp_id:x.dataset.emp,date:x.dataset.date,value:x.value}));
            const btn=document.getElementById('save-roster-btn');btn.disabled=true;btn.textContent='Saving...';
            fetch('/api/save_roster_matrix',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({updates})})
            .then(r=>r.json()).then(res=>{showToast(res.message||'Monthly roster saved','success');btn.disabled=false;btn.textContent='Save Monthly Roster';const notice=document.getElementById('roster-unsaved-notice');if(notice)notice.classList.add('hidden');loadRosterPlanner()})
            .catch(e=>{showToast(e.message,'danger');btn.disabled=false;btn.textContent='Save Monthly Roster'});
        }

        document.addEventListener('DOMContentLoaded', function() {
            startTimer();
            const rpm=document.getElementById('rp-month'); if(rpm && !rpm.value) rpm.value=new Date().toISOString().slice(0,7);
            setInterval(updateLiveClock, 1000);
            updateLiveClock();
        });
    </script>

<style id="mobile-responsive-fixes">
  html, body { max-width: 100%; overflow-x: hidden; }
  button, a, select, input, textarea { touch-action: manipulation; }
  input, select, textarea { max-width: 100%; }
  @media (max-width: 640px) {
    body { font-size: 14px; }
    header { gap: 0.4rem; }
    header select { max-width: 112px; }
    .fixed.inset-0 { padding: 0.35rem !important; }
    .fixed.inset-0 > div { max-width: calc(100vw - 0.7rem) !important; max-height: 96vh !important; }
    #roster-matrix-table { -webkit-overflow-scrolling: touch; }
    #roster-matrix-table .sticky { position: static !important; left: auto !important; }
    #roster-matrix-table table { min-width: 980px; }
    #roster-designation-menu { position: fixed !important; left: 0.5rem !important; right: 0.5rem !important; top: 8rem !important; width: auto !important; max-height: 55vh !important; }
    .id-card { width: min(340px, 94vw) !important; transform-origin: top center; }
    .download-btn { width: min(340px, 94vw); }
    iframe { min-height: 65vh; }
    table { font-size: 11px; }
    th, td { white-space: nowrap; }
    input[type="date"], input[type="month"], input[type="time"], select, textarea { min-height: 42px; }
    button, a.rounded-xl { min-height: 40px; }
  }
</style>
</head>
<body class="bg-slate-50 text-slate-800 antialiased flex h-screen overflow-hidden">
    
    <!-- Mobile Sidebar Overlay -->
    <div id="sidebar-overlay" onclick="toggleMobileMenu(false)" class="fixed inset-0 bg-slate-900/50 z-40 hidden lg:hidden backdrop-blur-sm transition-opacity"></div>

    <!-- Sidebar Navigation -->
    <aside id="sidebar" class="w-64 bg-white border-r border-slate-200 flex flex-col justify-between fixed inset-y-0 left-0 transform -translate-x-full lg:relative lg:translate-x-0 transition duration-200 ease-in-out z-50 h-screen">
        <div class="overflow-y-auto">
            <!-- Logo Header -->
            <div class="p-5 flex items-center justify-between border-b border-slate-100">
                <div class="flex items-center space-x-3">
                    <div class="bg-[#78b13f] p-2 rounded-xl shadow-sm">
                        <img src="{{ url_for('static', filename='fresmart.png') }}" alt="Logo" class="h-6 object-contain" onerror="this.style.display='none'">
                    </div>
                    <div>
                        <h2 class="text-sm font-bold text-slate-900 leading-tight">{{ portal_name }}</h2>
                        <p class="text-[10px] text-slate-400 font-medium">Attendance Portal</p>
                    </div>
                </div>
                <!-- Close Button (Mobile Only) -->
                <button onclick="toggleMobileMenu(false)" class="lg:hidden p-2 text-slate-400 hover:text-slate-600 rounded-lg">✕</button>
            </div>

            <!-- Menu Links -->
            <div class="p-4 space-y-1">
                <p class="text-[10px] font-bold uppercase tracking-wider text-slate-400 px-3 mb-2">Main Menu</p>
                <a href="/" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl bg-slate-900 text-white font-semibold text-xs shadow-sm">
                    <span>📊</span>
                    <span>Dashboard</span>
                </a>
                
                {% if session_has_permission('employee_id_cards') %}
                <a href="#" onclick="toggleModal('employee-list-modal', true); return false;" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs transition border border-transparent hover:border-slate-200">
                    <span>🪪</span><span>Employee ID Cards</span>
                </a>
                {% endif %}
                {% if session_has_permission('roster') %}
                <a href="#" onclick="toggleModal('roster-planner-modal', true); return false;" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs transition border border-transparent hover:border-slate-200">
                    <span>🗓</span>
                    <span>Roster Planner</span>
                </a>
                {% endif %}
                {% if session_has_permission('shift_approvals') %}
                <a href="#" onclick="toggleModal('shift-approvals-modal', true); return false;" class="flex items-center justify-between px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs transition">
                    <div class="flex items-center space-x-3">
                        <span>🔔</span>
                        <span>Shift Approvals</span>
                    </div>
                    {% if pending_shifts_count > 0 %}
                    <span class="bg-indigo-500 text-white text-[10px] font-bold px-2 py-0.5 rounded-full animate-pulse">{{ pending_shifts_count }}</span>
                    {% endif %}
                </a>
                {% endif %}
                
                {% if role == 'developer' %}
                <a href="#" onclick="toggleModal('machine-mgmt-modal', true); return false;" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs transition"><span>🖥️</span><span>Biometric Machines</span></a>
                <a href="#" onclick="toggleModal('user-mgmt-modal', true); return false;" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs transition border border-transparent hover:border-slate-200">
                    <span>🪪</span>
                    <span>User Management</span>
                </a>
                {% endif %}
                
                {% if session_has_permission('password_management') %}
                <a href="#" onclick="toggleModal('reset-approvals-modal', true); return false;" class="flex items-center justify-between px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs transition border border-transparent hover:border-slate-200">
                    <div class="flex items-center space-x-3">
                        <span>🔑</span>
                        <span>Manage Passwords</span>
                    </div>
                    {% if pending_resets_count > 0 %}
                    <span class="bg-rose-500 text-white text-[10px] font-bold px-2 py-0.5 rounded-full animate-pulse">{{ pending_resets_count }}</span>
                    {% endif %}
                </a>
                {% endif %}
                
                {% if role == 'employee' %}
                <a href="#" onclick="toggleModal('shift-request-modal', true); return false;" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs transition border border-transparent hover:border-slate-200">
                    <span>🔄</span>
                    <span>Shift/Off Request</span>
                </a>
                {% endif %}

                <a href="#" onclick="toggleModal('calendar-modal', true); return false;" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs transition">
                    <span>📅</span>
                    <span>Calendar & Rota</span>
                </a>
                {% if session_has_permission('leave_management') %}<a href="#" onclick="toggleModal('leave-modal', true); return false;" class="flex items-center justify-between px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs transition">
                    <div class="flex items-center space-x-3">
                        <span>🏖️</span>
                        <span>Leave Management</span>
                    </div>
                    {% if role in ['admin', 'developer'] and pending_leaves_count > 0 %}
                    <span class="bg-rose-500 text-white text-[10px] font-bold px-2 py-0.5 rounded-full animate-pulse">{{ pending_leaves_count }}</span>
                    {% endif %}
                </a>{% endif %}
                

                {% if session_has_permission('overtime_approval') %}
                <a href="#" onclick="toggleModal('overtime-approval-modal',true);return false;" class="flex items-center justify-between px-3 py-2.5 rounded-xl text-amber-700 bg-amber-50 hover:bg-amber-100 font-bold text-xs border border-amber-200">
                    <div class="flex items-center space-x-3"><span>⏱️</span><span>Overtime Approval</span></div>
                    {% if pending_overtime_count > 0 %}<span class="bg-amber-500 text-white px-2 py-0.5 rounded-full text-[10px]">{{pending_overtime_count}}</span>{% endif %}
                </a>
{% if session_has_permission('overtime_history') %}                <a href="#" onclick="toggleModal('overtime-history-modal',true);return false;" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs">
                    <span>🕘</span><span>Overtime History</span>
                </a>{% endif %}
                {% endif %}
                {% if role == 'developer' %}
                <a href="#" onclick="toggleModal('permission-control-modal',true);return false;" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-violet-700 bg-violet-50 hover:bg-violet-100 font-black text-xs border border-violet-200"><span>🔐</span><span>Portal Rights Control</span></a>
                {% endif %}
                {% if session_has_permission('workforce_hub') %}
                <a href="/workforce_hub" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-cyan-700 bg-cyan-50 hover:bg-cyan-100 font-bold text-xs transition border border-cyan-200">
                    <span>⚙️</span><span>Workforce Automation Hub</span>
                </a>
                {% endif %}
                {% if session_has_permission('attendance_corrections') %}
                <a href="/suite/corrections" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs"><span>✍️</span><span>Attendance Correction</span></a>
                {% endif %}{% if session_has_permission('employee_documents') %}<a href="/suite/documents" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs"><span>📁</span><span>My Documents</span></a>
                {% endif %}{% if session_has_permission('notifications') %}<a href="/suite/notifications" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs"><span>🔔</span><span>Notifications</span></a>
                {% endif %}
                {% if role in ['admin','developer'] %}<a href="/duplicate_punches" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-orange-800 bg-orange-50 hover:bg-orange-100 font-black text-xs border border-orange-200"><span>👆</span><span>Duplicate Punches</span></a>{% endif %}
                <a href="/promotion" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-fuchsia-700 bg-fuchsia-50 hover:bg-fuchsia-100 font-black text-xs border border-fuchsia-200"><span>📣</span><span>Promotion</span></a>
                <a href="/kvi_power_sku" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-amber-800 bg-amber-50 hover:bg-amber-100 font-black text-xs border border-amber-200"><span>⚡</span><span>KVI + Power SKU</span></a>
                <p class="text-[10px] font-bold uppercase tracking-wider text-slate-400 px-3 mt-6 mb-2">Team Management</p>
                {% if session_has_permission('payroll') %}<a href="#" onclick="toggleModal('salary-modal', true); return false;" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs transition border border-transparent hover:border-slate-200">
                    <span>💰</span>
                    <span>Payroll & Reports</span>
                </a>{% endif %}
                {% if role in ['admin','developer'] and session_has_permission('payroll') %}<a href="/export_last_six_months" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-emerald-700 bg-emerald-50 hover:bg-emerald-100 font-bold text-xs border border-emerald-200"><span>📥</span><span>Export Last 6 Months</span></a>{% endif %}
                
                {% if role == 'employee' %}
                <a href="/employee_id/{{ session.get('user_id') }}" target="_blank" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-emerald-700 bg-emerald-50 hover:bg-emerald-100 font-bold text-xs transition border border-emerald-200">
                    <span>🪪</span>
                    <span>{% if ui_language|default('en') == 'pt' %}Baixar Cartão de Identificação{% else %}Download ID Card{% endif %}</span>
                </a>
                {% elif session_has_permission('employees_info') %}
                <a href="/employee_id/{{ session.get('user_id') }}" target="_blank" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs transition">
                    <span>🪪</span><span>Employees Info (ID Card)</span>
                </a>
                {% endif %}
            </div>
        </div>

        <!-- Sidebar Footer -->
        <div class="p-4 border-t border-slate-100">
            <div class="bg-emerald-50 border border-emerald-100 rounded-2xl p-3.5 space-y-2">
                <div class="flex items-center space-x-2 text-emerald-800 font-bold text-xs">
                    <span>📢</span>
                    <span>Announcements</span>
                </div>
                <p class="text-[11px] text-slate-600 leading-tight">Biometric live tracking active for Attendance Portal.</p>
                <div class="text-[10px] text-emerald-600 font-bold pt-1">Dev: Sonu Kumar (NCSA0608)</div>
            </div>
        </div>
    </aside>

    <!-- Main Wrapper -->
    <div class="flex-1 flex flex-col h-screen overflow-hidden">
        
        <!-- Top Navbar (Responsive) -->
        <header class="bg-white border-b border-slate-200 px-4 sm:px-6 py-3.5 flex justify-between items-center z-10">
            <div class="flex items-center flex-1">
                <!-- Mobile Menu Button -->
                <button onclick="toggleMobileMenu(true)" class="lg:hidden mr-3 p-2 rounded-lg bg-slate-100 text-slate-600 hover:bg-slate-200 transition">
                    <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 6h16M4 12h16M4 18h16"></path></svg>
                </button>
                <div class="flex flex-col sm:flex-row sm:items-center sm:space-x-3">
                    <h1 class="text-sm sm:text-base font-black text-slate-900 tracking-tight">Dashboard</h1>{% if accessible_stores|length > 1 %}<select onchange="location.href='/?store='+this.value" class="ml-2 text-xs border rounded-lg px-2 py-1">{% for st in accessible_stores %}<option value="{{st}}" {% if st==store %}selected{% endif %}>{{st}}</option>{% endfor %}</select>{% endif %}
                    <span class="hidden sm:inline-block text-xs text-slate-400 font-medium">| Good day, {{ logged_user_name }}</span>
                </div>
            </div>

            <div class="flex items-center gap-2 flex-wrap justify-end">
                <select onchange="window.location.href='/set_language/'+this.value" class="text-xs border border-slate-200 rounded-xl px-2 py-2 bg-white font-bold" title="Language / Idioma">
                    <option value="en" {% if ui_language=='en' %}selected{% endif %}>English</option>
                    <option value="pt" {% if ui_language=='pt' %}selected{% endif %}>Português</option>
                </select>
                {% if role == 'employee' %}
                <a href="/employee_id/{{ session.get('user_id') }}" target="_blank" class="bg-indigo-50 hover:bg-indigo-100 text-indigo-800 text-xs font-bold px-2 py-1.5 sm:px-3 sm:py-2 rounded-xl transition border border-indigo-200 flex items-center gap-1.5" title="{% if ui_language|default('en') == 'pt' %}Baixar Cartão de Identificação{% else %}Download ID Card{% endif %}">
                    <span>🪪</span><span class="hidden lg:inline">{% if ui_language|default('en') == 'pt' %}Baixar Cartão{% else %}ID Card{% endif %}</span>
                </a>
                {% endif %}
                <button onclick="toggleModal('leave-modal', true)" class="relative bg-emerald-50 hover:bg-emerald-100 text-emerald-800 text-xs font-bold px-2 py-1.5 sm:px-3 sm:py-2 rounded-xl transition border border-emerald-200 flex items-center space-x-1.5">
                    <span class="hidden sm:inline">🏖 Leave Portal</span>
                    <span class="sm:hidden">🏖 Leaves</span>
                    {% if role in ['admin', 'developer'] and pending_leaves_count > 0 %}
                    <span class="bg-rose-500 text-white text-[10px] px-1.5 py-0.2 rounded-full font-black animate-bounce">{{ pending_leaves_count }}</span>
                    {% endif %}
                </button>

                <!-- Clock and Status (Hidden on very small screens to save space) -->
                <div class="hidden md:flex text-xs bg-slate-50 px-3 py-2 rounded-xl border border-slate-200 items-center space-x-2">
                    {% if role in ['admin','developer'] %}<span class="h-2 w-2 {% if stats.device_online %}bg-emerald-500{% else %}bg-red-500{% endif %} rounded-full animate-pulse"></span><span class="text-slate-600 font-medium">Device: <strong class="{% if stats.device_online %}text-emerald-600{% else %}text-red-600{% endif %}">{% if stats.device_online %}Online{% else %}Offline{% endif %}</strong></span><span class="text-slate-300">|</span><span class="text-slate-500 font-mono text-[10px]">{{ current_machine.ip }}:{{ current_machine.port }}</span><span class="text-slate-300">|</span>{% endif %}
                    <span id="live-digital-clock" class="text-slate-700 font-semibold"></span>
                    {% if role in ['admin','developer'] %}<span class="text-slate-300">|</span><span class="text-slate-500">Sync: <strong id="countdown-timer" class="text-emerald-600 font-mono">03:00</strong></span>{% endif %}
                </div>

                <div class="flex items-center space-x-2 bg-slate-100 border border-slate-200 px-2.5 py-1.5 sm:px-3 sm:py-1.5 rounded-xl text-xs font-bold text-slate-700">
                    <button onclick="toggleModal('profile-modal',true)" class="flex items-center gap-1">{% if current_user.profile_photo %}<img src="/uploads/{{ current_user.profile_photo }}" class="w-7 h-7 rounded-full object-cover border">{% else %}<span>👤</span>{% endif %}<span class="hidden sm:inline">{{ logged_user_name }}</span></button>
                    <a href="/logout" class="text-rose-600 hover:text-rose-700 sm:ml-2 font-semibold">Logout 🔒</a>
                </div>

                {% if role == 'developer' %}
                <button onclick="secureShutdown()" class="bg-rose-50 hover:bg-rose-100 text-rose-700 text-xs font-bold px-2.5 py-1.5 sm:px-3 sm:py-2 rounded-xl transition border border-rose-200">
                    <span class="hidden sm:inline">🛑 Shutdown</span>
                    <span class="sm:hidden">🛑</span>
                </button>
                {% endif %}
            </div>
        </header>

        <!-- Main Content Area -->
        <main class="flex-1 overflow-y-auto p-4 sm:p-6 space-y-6">
            
            {% with messages = get_flashed_messages(with_categories=true) %}
                {% if messages %}{% for category, message in messages %}
                <div class="hrms-flash hidden" data-category="{{ category }}" data-message="{{ message|e }}"></div>
                {% endfor %}{% endif %}
            {% endwith %}

            <!-- Quick Top Cards / Stat Overview -->
            <div class="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-9 gap-3">
                <div onclick="filterByStatus('Present')" class="stat-card bg-white p-3.5 rounded-2xl shadow-sm border border-slate-200 flex items-center justify-between border-l-4 border-l-emerald-500">
                    <div>
                        <p class="text-[10px] font-bold uppercase tracking-wider text-slate-400">Present</p>
                        <h3 class="text-xl font-black text-emerald-600 mt-0.5">{{ stats.present }}</h3>
                    </div>
                </div>
                <div onclick="filterByStatus('Absent')" class="stat-card bg-white p-3.5 rounded-2xl shadow-sm border border-slate-200 flex items-center justify-between border-l-4 border-l-rose-500">
                    <div>
                        <p class="text-[10px] font-bold uppercase tracking-wider text-slate-400">Absent</p>
                        <h3 class="text-xl font-black text-rose-600 mt-0.5">{{ stats.absent }}</h3>
                    </div>
                </div>
                <div onclick="filterByStatus('ML')" class="stat-card bg-white p-3.5 rounded-2xl shadow-sm border border-slate-200 flex items-center justify-between border-l-4 border-l-cyan-500">
                    <div>
                        <p class="text-[10px] font-bold uppercase tracking-wider text-slate-400">Med/Leave</p>
                        <h3 class="text-xl font-black text-cyan-600 mt-0.5">{{ stats.ml }}</h3>
                    </div>
                </div>
                <div onclick="filterByStatus('Weekly Off')" class="stat-card bg-white p-3.5 rounded-2xl shadow-sm border border-slate-200 flex items-center justify-between border-l-4 border-l-slate-400">
                    <div>
                        <p class="text-[10px] font-bold uppercase tracking-wider text-slate-400">Week Off</p>
                        <h3 class="text-xl font-black text-slate-700 mt-0.5">{{ stats.off }}</h3>
                    </div>
                </div>
                <div onclick="filterByStatus('Late Arrival')" class="stat-card bg-white p-3.5 rounded-2xl shadow-sm border border-slate-200 flex items-center justify-between border-l-4 border-l-amber-500">
                    <div>
                        <p class="text-[10px] font-bold uppercase tracking-wider text-slate-400">Late Arr.</p>
                        <h3 class="text-xl font-black text-amber-600 mt-0.5">{{ stats.late_arrival }}</h3>
                    </div>
                </div>
                <div onclick="filterByStatus('Mis Punch')" class="stat-card bg-white p-3.5 rounded-2xl shadow-sm border border-slate-200 flex items-center justify-between border-l-4 border-l-orange-500">
                    <div>
                        <p class="text-[10px] font-bold uppercase tracking-wider text-slate-400">Mis-Punch</p>
                        <h3 class="text-xl font-black text-orange-600 mt-0.5">{{ stats.mispunch }}</h3>
                    </div>
                </div>
                <div onclick="filterByStatus('Shift A')" class="stat-card bg-white p-3.5 rounded-2xl shadow-sm border border-slate-200 flex items-center justify-between border-l-4 border-l-blue-500">
                    <div>
                        <p class="text-[10px] font-bold uppercase tracking-wider text-slate-400">Shift A</p>
                        <h3 class="text-xl font-black text-blue-600 mt-0.5">{{ stats.shift_a }}</h3>
                    </div>
                </div>
                <div onclick="filterByStatus('Shift B')" class="stat-card bg-white p-3.5 rounded-2xl shadow-sm border border-slate-200 flex items-center justify-between border-l-4 border-l-indigo-500">
                    <div>
                        <p class="text-[10px] font-bold uppercase tracking-wider text-slate-400">Shift B</p>
                        <h3 class="text-xl font-black text-indigo-600 mt-0.5">{{ stats.shift_b }}</h3>
                    </div>
                </div>
                <div onclick="filterByStatus('ALL')" class="stat-card bg-white p-3.5 rounded-2xl shadow-sm border border-slate-200 flex items-center justify-between border-l-4 border-l-emerald-600 col-span-2 lg:col-span-1">
                    <div>
                        <p class="text-[10px] font-bold uppercase tracking-wider text-slate-400">Tot Hrs</p>
                        <h3 class="text-xl font-black text-emerald-600 mt-0.5">{{ stats.total_hrs }}</h3>
                    </div>
                </div>
            </div>

            <!-- Filter Controls Bar -->
            <div class="bg-white rounded-2xl shadow-sm border border-slate-200 p-4 sm:p-5">
                <form id="filter-form" method="GET" action="/" class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-4 items-end">
                    <div>
                        <label class="block text-[11px] font-bold uppercase tracking-wider text-slate-500 mb-1.5">Start Date</label>
                        <input type="date" name="start_date" value="{{ start_date }}" onchange="this.form.submit()" class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3.5 py-2.5 text-sm font-medium focus:ring-2 focus:ring-emerald-500 focus:outline-none">
                    </div>
                    <div>
                        <label class="block text-[11px] font-bold uppercase tracking-wider text-slate-500 mb-1.5">End Date</label>
                        <input type="date" name="end_date" value="{{ end_date }}" onchange="this.form.submit()" class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3.5 py-2.5 text-sm font-medium focus:ring-2 focus:ring-emerald-500 focus:outline-none">
                    </div>
                    {% if role == 'admin' or role == 'developer' %}
                    <div>
                        <label class="block text-[11px] font-bold uppercase tracking-wider text-slate-500 mb-1.5">Employee Filter</label>
                        <select name="employee" onchange="this.form.submit()" class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3.5 py-2.5 text-sm font-medium focus:ring-2 focus:ring-emerald-500 focus:outline-none">
                            <option value="ALL">-- All Personnel --</option>
                            {% for emp in all_users %}
                                <option value="{{ emp.user_id }}" {% if selected_emp == emp.user_id %}selected{% endif %}>{{ emp.name }} ({{ emp.user_id }})</option>
                            {% endfor %}
                        </select>
                    </div>
                    {% else %}
                    <div>
                        <label class="block text-[11px] font-bold uppercase tracking-wider text-slate-500 mb-1.5">Logged In As</label>
                        <input type="hidden" name="employee" value="{{ selected_emp }}">
                        <input type="text" disabled value="{{ selected_emp }}" class="w-full bg-slate-100 border border-slate-200 rounded-xl px-3.5 py-2.5 text-sm font-bold text-emerald-700 cursor-not-allowed">
                    </div>
                    {% endif %}
                    <div class="flex space-x-2">
                        {% if role == 'admin' or role == 'developer' %}
                        <button type="button" onclick="toggleModal('export-modal', true)" class="flex-1 bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs uppercase tracking-wider py-3 px-3 rounded-xl text-center shadow-md transition">Export 📥</button>
                        {% endif %}
                        <button type="button" onclick="toggleModal('calendar-modal', true)" class="flex-1 bg-slate-900 hover:bg-slate-800 text-white font-bold text-xs uppercase tracking-wider py-3 px-3 rounded-xl text-center shadow-md transition">📅 Rota</button>
                    </div>
                </form>
                
                <div class="flex items-center space-x-2 mt-4 pt-4 border-t border-slate-100 text-[11px] flex-wrap gap-y-2">
                    <span class="text-slate-400 font-bold uppercase tracking-wide mr-1">Quick Range:</span>
                    <button type="button" onclick="setQuickDate('today')" class="px-3 py-1.5 bg-slate-100 hover:bg-emerald-50 hover:text-emerald-700 text-slate-700 rounded-lg font-semibold transition">Today</button>
                    <button type="button" onclick="setQuickDate('yesterday')" class="px-3 py-1.5 bg-slate-100 hover:bg-emerald-50 hover:text-emerald-700 text-slate-700 rounded-lg font-semibold transition">Yesterday</button>
                    <button type="button" onclick="setQuickDate('week')" class="px-3 py-1.5 bg-slate-100 hover:bg-emerald-50 hover:text-emerald-700 text-slate-700 rounded-lg font-semibold transition">This Week</button>
                    <button type="button" onclick="setQuickDate('month')" class="px-3 py-1.5 bg-slate-100 hover:bg-emerald-50 hover:text-emerald-700 text-slate-700 rounded-lg font-semibold transition">This Month</button>
                    <button type="button" onclick="setQuickDate('last_month')" class="px-3 py-1.5 bg-slate-100 hover:bg-emerald-50 hover:text-emerald-700 text-slate-700 rounded-lg font-semibold transition">Last Month</button>
                </div>
            </div>

            <!-- Attendance Data Table -->
            <div class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
                <div class="p-4 bg-slate-50 border-b border-slate-200 flex flex-col sm:flex-row justify-between items-center gap-3">
                    <div class="text-[10px] sm:text-xs text-slate-500 font-semibold w-full sm:w-auto text-center sm:text-left">
                        💡 Shift A = 06:00-10:00 | Shift B = >10:00 AM
                    </div>
                    <div class="w-full sm:w-auto">
                        <input type="text" id="table-search-input" onkeyup="filterTableSearch()" placeholder="🔍 Search Name, ID..." class="w-full sm:w-64 bg-white border border-slate-300 rounded-xl px-3.5 py-2 text-xs shadow-sm focus:outline-none focus:border-emerald-500 font-medium">
                    </div>
                </div>
                <!-- Overflow-x-auto ensures horizontal scrolling on mobile -->
                <div class="overflow-x-auto w-full block">
                    <table id="attendance-table" class="w-full text-left border-collapse excel-table min-w-[800px]">
                        <thead>
                            <tr class="bg-slate-100 text-slate-600 uppercase text-[11px] font-bold tracking-wider">
                                <th class="py-3 px-4">Sr.</th>
                                <th class="py-3 px-4 sortable" onclick="sortTable(1)">Date ↕</th>
                                <th class="py-3 px-4 sortable" onclick="sortTable(2)">ID ↕</th>
                                <th class="py-3 px-4 sortable" onclick="sortTable(3)">Employee Name ↕</th>
                                <th class="py-3 px-4 sortable" onclick="sortTable(4)">Dept ↕</th>
                                <th class="py-3 px-4 sortable" onclick="sortTable(5)">Store In ↕</th>
                                <th class="py-3 px-4">Lunch Out</th>
                                <th class="py-3 px-4">Lunch In</th>
                                <th class="py-3 px-4 sortable" onclick="sortTable(8)">Out Time ↕</th>
                                <th class="py-3 px-4 sortable" onclick="sortTable(9)">Total Lunch ↕</th>
                                <th class="py-3 px-4 sortable" onclick="sortTable(10)">Working Hrs ↕</th>
                                <th class="py-3 px-4 sortable" onclick="sortTable(11)">Hora Extra ↕</th>
                                <th class="py-3 px-4 text-center sortable" onclick="sortTable(12)">Status ↕</th>
                            </tr>
                        </thead>
                        <tbody class="text-sm text-slate-700 divide-y divide-slate-100">
                            {% if logs %}
                                {% for log in logs %}
                                <tr class="hover:bg-slate-50 transition-colors" data-late="{{ log.is_late }}" data-shift="{{ log.shift_type }}">
                                    <td class="py-3 px-4 nowrap-cell font-medium text-slate-400">{{ loop.index }}</td>
                                    <td class="py-3 px-4 nowrap-cell font-medium">{{ log.date }}</td>
                                    <td class="py-3 px-4 nowrap-cell text-slate-500 font-mono text-xs">{{ log.user_id }}</td>
                                    <td class="py-3 px-4 font-bold text-slate-900 nowrap-cell">{{ log.name }}</td>
                                    <td class="py-3 px-4 nowrap-cell"><span class="px-2.5 py-1 rounded-lg bg-emerald-50 text-emerald-700 text-[11px] font-bold">{{ log.dept }}</span></td>
                                    
                                    <td class="py-3 px-4 nowrap-cell font-mono text-xs group relative">
                                        {{ log.store_in }}
                                        {% if role == 'developer' %}
                                        <span onclick="inlineEdit('{{ log.date }}', '{{ log.user_id }}', 'store_in', '{{ log.store_in }}')" class="cursor-pointer ml-1 text-amber-400 hover:text-amber-600 sm:opacity-0 group-hover:opacity-100 transition-opacity" title="Edit Store In">✎</span>
                                        {% endif %}
                                        {% if log.shift_type == 'Shift A' %}
                                            <span class="text-[9px] bg-blue-100 text-blue-700 px-1 py-0.5 rounded font-sans font-bold ml-1 absolute top-1/2 -translate-y-1/2 right-2">A</span>
                                        {% elif log.shift_type == 'Shift B' %}
                                            <span class="text-[9px] bg-indigo-100 text-indigo-700 px-1 py-0.5 rounded font-sans font-bold ml-1 absolute top-1/2 -translate-y-1/2 right-2">B</span>
                                        {% endif %}
                                    </td>
                                    <td class="py-3 px-4 nowrap-cell font-mono text-xs text-slate-500 group">
                                        {{ log.lunch_out }}
                                        {% if role == 'developer' %}
                                        <span onclick="inlineEdit('{{ log.date }}', '{{ log.user_id }}', 'lunch_out', '{{ log.lunch_out }}')" class="cursor-pointer ml-1 text-amber-400 hover:text-amber-600 sm:opacity-0 group-hover:opacity-100 transition-opacity" title="Edit Lunch Out">✎</span>
                                        {% endif %}
                                    </td>
                                    <td class="py-3 px-4 nowrap-cell font-mono text-xs text-slate-500 group">
                                        {{ log.lunch_in }}
                                        {% if role == 'developer' %}
                                        <span onclick="inlineEdit('{{ log.date }}', '{{ log.user_id }}', 'lunch_in', '{{ log.lunch_in }}')" class="cursor-pointer ml-1 text-amber-400 hover:text-amber-600 sm:opacity-0 group-hover:opacity-100 transition-opacity" title="Edit Lunch In">✎</span>
                                        {% endif %}
                                    </td>
                                    <td class="py-3 px-4 nowrap-cell font-mono text-xs group">
                                        {{ log.out_time }}
                                        {% if role == 'developer' %}
                                        <span onclick="inlineEdit('{{ log.date }}', '{{ log.user_id }}', 'out_time', '{{ log.out_time }}')" class="cursor-pointer ml-1 text-amber-400 hover:text-amber-600 sm:opacity-0 group-hover:opacity-100 transition-opacity" title="Edit Out Time">✎</span>
                                        {% endif %}
                                    </td>
                                    
                                    <td class="py-3 px-4 font-bold nowrap-cell font-mono text-xs {% if log.lunch_seconds > 3600 %}text-rose-600 bg-rose-50/50{% else %}text-slate-700{% endif %}">{{ log.total_lunch }}</td>
                                    <td class="py-3 px-4 font-bold text-slate-900 nowrap-cell font-mono text-xs">{{ log.total_hours }}</td>
                                    <td class="py-3 px-4 font-bold nowrap-cell font-mono text-xs {% if log.variance_type == 'positive' %}text-emerald-600{% elif log.variance_type == 'pending' %}text-amber-600{% elif log.variance_type == 'rejected' %}text-slate-400{% elif log.variance_type == 'negative' %}text-rose-600{% else %}text-slate-600{% endif %}">{{ log.net_variance }}</td>
                                    <td class="py-3 px-4 text-center nowrap-cell">
                                        {% if log.status == 'Weekly Off' %}
                                            <span class="px-2 py-1 rounded-full text-[10px] sm:text-xs font-bold bg-slate-100 text-slate-600 border border-slate-200">Weekly Off</span>
                                        {% elif log.status == 'Absent' %}
                                            <span class="px-2 py-1 rounded-full text-[10px] sm:text-xs font-bold bg-rose-50 text-rose-600 border border-rose-200">Absent</span>
                                        {% elif 'F01;1' in log.status or 'F05;1' in log.status or 'F10;1' in log.status or 'F51;1' in log.status or 'F60;1' in log.status or 'F61;1' in log.status or 'F62;1' in log.status %}
                                            <span class="px-2 py-1 rounded-full text-[10px] sm:text-xs font-bold bg-cyan-50 text-cyan-700 border border-cyan-200">{{ log.status }}</span>
                                        {% elif log.status == 'ML' %}
                                            <span class="px-2 py-1 rounded-full text-[10px] sm:text-xs font-bold bg-cyan-50 text-cyan-700 border border-cyan-200">Medical/Leave</span>
                                        {% elif log.status == 'Present' %}
                                            <span class="px-2 py-1 rounded-full text-[10px] sm:text-xs font-bold bg-emerald-50 text-emerald-700 border border-emerald-200">Present</span>
                                        {% elif log.status == 'Mis Punch' %}
                                            <span class="px-2 py-1 rounded-full text-[10px] sm:text-xs font-bold bg-orange-50 text-orange-700 border border-orange-200">Mis Punch</span>
                                        {% endif %}
                                    </td>
                                </tr>
                                {% endfor %}
                            {% else %}
                                <tr><td colspan="13" class="text-center py-16 text-slate-400 font-medium">No attendance records found for this selection.</td></tr>
                            {% endif %}
                        </tbody>
                        <tfoot class="bg-slate-100 font-bold text-slate-900 text-sm border-t border-slate-200">
                            <tr>
                                <td colspan="9" class="py-4 px-4 text-right uppercase text-[10px] sm:text-xs tracking-wider text-slate-500">Total Summary:</td>
                                <td class="py-4 px-4 text-emerald-700 font-mono text-xs sm:text-sm">{{ grand_total_lunch_hours }}</td>
                                <td class="py-4 px-4 text-slate-900 font-mono text-xs sm:text-sm">{{ grand_total_hours }}</td>
                                <td class="py-4 px-4 font-mono text-xs sm:text-sm {% if stats.variance_type == 'positive' %}text-emerald-700{% else %}text-rose-700{% endif %}" colspan="2">{{ grand_total_variance }}</td>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>
        </main>
    </div>

    <!-- Salary Slips & Export Modal (Merged Payroll) -->
    <div id="salary-modal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[60] flex items-center justify-center hidden">
        <div class="bg-white rounded-2xl shadow-2xl border border-slate-200 p-4 sm:p-6 w-full max-w-4xl mx-4 space-y-6 max-h-[90vh] flex flex-col">
            <div class="flex justify-between items-center border-b border-slate-100 pb-4">
                <h3 class="text-sm sm:text-lg font-bold text-slate-900 flex items-center gap-2">💰 Payroll & Reports</h3>
                <button onclick="toggleModal('salary-modal', false)" class="text-slate-400 hover:text-slate-600 font-bold text-lg">✕</button>
            </div>
            
            <div class="overflow-y-auto flex-1 space-y-6">
                <!-- EXPORT SECTION (Admin/Dev) -->
                {% if role in ['admin', 'developer'] %}
                <div class="bg-slate-50 border border-slate-200 rounded-xl p-4">
                    <h4 class="text-xs sm:text-sm font-bold text-slate-900 mb-3">Download Attendance Reports</h4>
                    <div class="flex flex-col sm:flex-row gap-4">
                        <a href="/export?start_date={{ start_date }}&end_date={{ end_date }}&employee={{ selected_emp }}" class="flex-1 bg-white border border-slate-300 hover:border-emerald-500 p-3 rounded-xl text-center transition group">
                            <div class="font-bold text-slate-800 text-sm group-hover:text-emerald-700">Standard Row Export 📄</div>
                            <div class="text-[10px] text-slate-500 mt-1">Individual records per date</div>
                        </a>
                        <a href="/export_matrix?start_date={{ start_date }}&end_date={{ end_date }}&employee={{ selected_emp }}" class="flex-1 bg-white border border-slate-300 hover:border-emerald-500 p-3 rounded-xl text-center transition group">
                            <div class="font-bold text-slate-800 text-sm group-hover:text-emerald-700">Employee Matrix 📅</div>
                            <div class="text-[10px] text-slate-500 mt-1">Monthly grid with Leave Codes</div>
                        </a>
                    </div>
                </div>
                
                <!-- UPLOAD BULK SALARY SLIP (Admin/Dev) -->
                <div class="bg-indigo-50 border border-indigo-200 rounded-xl p-4">
                    <h4 class="text-xs sm:text-sm font-bold text-indigo-900 mb-3">Bulk Upload (Merged PDF)</h4>
                    <form action="/upload_bulk_salary" method="POST" enctype="multipart/form-data" class="flex flex-col sm:flex-row items-end gap-3" onsubmit="uploadLoading()">
                        <div class="w-full sm:flex-1">
                            <label class="block text-[10px] font-bold uppercase text-indigo-700 mb-1">Month & Year</label>
                            <input type="month" name="salary_month" required class="w-full bg-white border border-indigo-300 rounded-lg px-3 py-2 text-xs focus:ring-2 focus:ring-indigo-500">
                        </div>
                        <div class="w-full sm:flex-1">
                            <label class="block text-[10px] font-bold uppercase text-indigo-700 mb-1">Master PDF</label>
                            <input type="file" name="salary_pdf" accept=".pdf" required class="w-full bg-white border border-indigo-300 rounded-lg px-2 py-1.5 text-[11px] focus:ring-2 focus:ring-indigo-500">
                        </div>
                        <button id="upload-btn" type="submit" class="w-full sm:w-auto bg-indigo-600 hover:bg-indigo-700 text-white font-bold px-4 py-2 rounded-lg text-xs transition">
                            <span id="upload-btn-text">Split & Upload</span>
                        </button>
                    </form>
                </div>

                <!-- UPLOAD INDIVIDUAL SALARY SLIP (Admin/Dev) -->
                <div class="bg-emerald-50 border border-emerald-200 rounded-xl p-4">
                    <h4 class="text-xs sm:text-sm font-bold text-emerald-900 mb-3">Individual Slip Upload</h4>
                    <form action="/upload_individual_salary" method="POST" enctype="multipart/form-data" class="flex flex-col sm:flex-row items-end gap-3">
                        <div class="w-full sm:flex-1">
                            <label class="block text-[10px] font-bold uppercase text-emerald-700 mb-1">Employee</label>
                            <select name="emp_id" required class="w-full bg-white border border-emerald-300 rounded-lg px-2 py-2 text-xs focus:ring-2 focus:ring-emerald-500">
                                <option value="">-- Select Employee --</option>
                                {% for emp in all_users %}
                                    <option value="{{ emp.user_id }}">{{ emp.name }} ({{ emp.user_id }})</option>
                                {% endfor %}
                            </select>
                        </div>
                        <div class="w-full sm:flex-1">
                            <label class="block text-[10px] font-bold uppercase text-emerald-700 mb-1">Month & Year</label>
                            <input type="month" name="salary_month" required class="w-full bg-white border border-emerald-300 rounded-lg px-3 py-2 text-xs focus:ring-2 focus:ring-emerald-500">
                        </div>
                        <div class="w-full sm:flex-1">
                            <label class="block text-[10px] font-bold uppercase text-emerald-700 mb-1">PDF File</label>
                            <input type="file" name="individual_pdf" accept=".pdf" required class="w-full bg-white border border-emerald-300 rounded-lg px-2 py-1.5 text-[11px] focus:ring-2 focus:ring-emerald-500">
                        </div>
                        <button type="submit" class="w-full sm:w-auto bg-emerald-600 hover:bg-emerald-700 text-white font-bold px-4 py-2 rounded-lg text-xs transition">
                            Upload
                        </button>
                    </form>
                </div>
                {% endif %}

                <!-- VIEW SALARY SLIPS (All) -->
                <div>
                    <div class="flex justify-between items-end mb-3">
                        <h4 class="text-xs sm:text-sm font-bold text-slate-900">{% if role == 'employee' %}My Salary Slips{% else %}Uploaded Salary Slips{% endif %}</h4>
                        <div class="w-32 sm:w-48">
                            <input type="month" id="salary-month-select" onchange="filterSalarySlips()" class="w-full bg-slate-50 border border-slate-300 rounded-lg px-2 py-1 text-xs focus:outline-none">
                        </div>
                    </div>
                    <div class="border border-slate-200 rounded-xl overflow-x-auto">
                        <table class="w-full text-left min-w-[300px]">
                            <thead class="bg-slate-100 text-slate-600 uppercase text-[10px] font-bold sticky top-0">
                                <tr>
                                    <th class="py-2 px-3 border-b border-slate-200">Month</th>
                                    {% if role in ['admin', 'developer'] %}<th class="py-2 px-3 border-b border-slate-200">Employee</th>{% endif %}
                                    <th class="py-2 px-3 border-b border-slate-200 text-center">Action</th>
                                </tr>
                            </thead>
                            <tbody class="text-[11px] sm:text-xs text-slate-700 divide-y divide-slate-100">
                                {% if salary_slips %}
                                    {% for slip in salary_slips|reverse %}
                                    <tr class="hover:bg-slate-50 salary-row" data-month="{{ slip.month }}">
                                        <td class="py-2 px-3 font-mono font-bold">{{ slip.month }}</td>
                                        {% if role in ['admin', 'developer'] %}
                                        <td class="py-2 px-3 font-bold">{{ slip.emp_name }} <span class="text-[9px] text-slate-400">({{ slip.user_id }})</span></td>
                                        {% endif %}
                                        <td class="py-2 px-3 text-center">
                                            <div class="flex justify-center gap-2">
                                                <button onclick="openPdfViewer('/salary_file/{{ slip.file_id }}')" class="bg-blue-50 text-blue-600 hover:bg-blue-100 border border-blue-200 font-bold px-2 py-1 rounded text-[10px] transition">View</button>
                                                {% if role in ['admin', 'developer'] %}
                                                <a href="/delete_salary/{{ slip.file_id }}" onclick="return confirm('Are you sure you want to delete this salary slip?');" class="bg-rose-50 text-rose-600 hover:bg-rose-100 border border-rose-200 font-bold px-2 py-1 rounded text-[10px] transition">Del</a>
                                                {% endif %}
                                            </div>
                                        </td>
                                    </tr>
                                    {% endfor %}
                                {% else %}
                                    <tr><td colspan="3" class="text-center py-6 text-slate-400">No salary slips uploaded yet.</td></tr>
                                {% endif %}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- PDF Viewer Modal (Restricted Download) -->
    <div id="pdf-viewer-modal" class="fixed inset-0 bg-slate-900/90 backdrop-blur-md z-[70] flex items-center justify-center hidden p-2 sm:p-4">
        <div class="bg-slate-800 rounded-xl shadow-2xl p-2 w-full max-w-4xl h-full sm:h-[90vh] flex flex-col relative">
            <button onclick="toggleModal('pdf-viewer-modal', false); document.getElementById('pdf-viewer-frame').src='';" class="absolute -top-2 -right-2 sm:-top-4 sm:-right-4 bg-rose-600 hover:bg-rose-700 text-white rounded-full w-8 h-8 flex items-center justify-center font-bold shadow-lg z-10">✕</button>
            <div class="flex-1 rounded-lg overflow-hidden bg-white" oncontextmenu="return false;">
                <!-- PDF embedded using iframe -->
                <iframe id="pdf-viewer-frame" class="w-full h-full pointer-events-none" style="pointer-events: auto;" src=""></iframe>
            </div>
            <div class="text-center mt-2 text-[10px] text-slate-400 uppercase tracking-widest font-bold">Confidential Document</div>
        </div>
    </div>

    <!-- Calendar & Rota Modal -->
    <div id="calendar-modal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[60] flex items-center justify-center hidden p-2">
        <div class="bg-white rounded-2xl shadow-2xl border border-slate-200 p-5 w-full {% if role=='employee' %}max-w-5xl{% else %}max-w-2xl{% endif %} mx-auto space-y-4 max-h-[92vh] flex flex-col">
            <div class="flex justify-between items-center border-b border-slate-100 pb-3">
                <div>
                    <h3 class="text-sm sm:text-lg font-bold text-slate-900">📅 {% if role=='employee' %}{% if ui_language=='pt' %}Minha Escala Mensal{% else %}My Monthly Roster{% endif %}{% else %}Calendar & Rota Rotation{% endif %}</h3>
                    {% if role=='employee' %}<p class="text-[11px] text-slate-500 mt-1">{{ employee_roster_month_label }} • {{ logged_user_name }}</p>{% endif %}
                </div>
                <button onclick="toggleModal('calendar-modal', false)" class="text-slate-400 hover:text-slate-600 font-bold text-lg">✕</button>
            </div>
            {% if role == 'employee' %}
            <div class="grid grid-cols-2 sm:grid-cols-4 gap-2">
                <div class="rounded-xl bg-emerald-50 border border-emerald-200 p-3"><span class="text-[10px] text-emerald-700">{% if ui_language=='pt' %}Turno A{% else %}Shift A{% endif %}</span><b class="block text-xl text-emerald-800">{{ employee_roster_counts.shift_a }}</b></div>
                <div class="rounded-xl bg-indigo-50 border border-indigo-200 p-3"><span class="text-[10px] text-indigo-700">{% if ui_language=='pt' %}Turno B{% else %}Shift B{% endif %}</span><b class="block text-xl text-indigo-800">{{ employee_roster_counts.shift_b }}</b></div>
                <div class="rounded-xl bg-rose-50 border border-rose-200 p-3"><span class="text-[10px] text-rose-700">{% if ui_language=='pt' %}Folga Semanal{% else %}Weekly Off{% endif %}</span><b class="block text-xl text-rose-800">{{ employee_roster_counts.weekly_off }}</b></div>
                <div class="rounded-xl bg-slate-100 border border-slate-200 p-3"><span class="text-[10px] text-slate-600">{% if ui_language=='pt' %}Padrão{% else %}Default{% endif %}</span><b class="block text-xl text-slate-800">{{ employee_roster_counts.default }}</b></div>
            </div>
            <div class="overflow-auto border border-slate-200 rounded-xl flex-1">
                <table class="w-full min-w-[720px] text-xs">
                    <thead class="sticky top-0 bg-slate-900 text-white"><tr><th class="p-3 text-left">{% if ui_language=='pt' %}Data{% else %}Date{% endif %}</th><th class="p-3 text-left">{% if ui_language=='pt' %}Dia{% else %}Day{% endif %}</th><th class="p-3 text-left">{% if ui_language=='pt' %}Turno / Folga{% else %}Shift / Weekly Off{% endif %}</th><th class="p-3 text-left">{% if ui_language=='pt' %}Estado{% else %}Status{% endif %}</th></tr></thead>
                    <tbody>{% for day in employee_month_roster %}<tr class="border-b hover:bg-slate-50 {% if day.is_today %}bg-cyan-50{% endif %}"><td class="p-3 font-mono">{{ day.display_date }}</td><td class="p-3 font-bold">{% if ui_language=='pt' %}{{ day.weekday_pt }}{% else %}{{ day.weekday_en }}{% endif %}</td><td class="p-3"><span class="inline-block rounded-full px-3 py-1 font-bold {% if day.value=='Shift A' %}bg-emerald-100 text-emerald-700{% elif day.value=='Shift B' %}bg-indigo-100 text-indigo-700{% elif day.value=='Weekly Off' %}bg-rose-100 text-rose-700{% else %}bg-slate-100 text-slate-600{% endif %}">{% if ui_language=='pt' %}{{ day.value_pt }}{% else %}{{ day.value_en }}{% endif %}</span></td><td class="p-3">{% if day.is_today %}<span class="text-cyan-700 font-bold">{% if ui_language=='pt' %}Hoje{% else %}Today{% endif %}</span>{% else %}<span class="text-slate-400">{% if ui_language=='pt' %}Programado{% else %}Scheduled{% endif %}</span>{% endif %}</td></tr>{% endfor %}</tbody>
                </table>
            </div>
            <p class="text-[10px] text-slate-500">{% if ui_language=='pt' %}Esta escala mostra apenas os seus turnos e folgas do mês atual.{% else %}This roster shows only your shifts and weekly-off dates for the current month.{% endif %}</p>
            {% else %}
            <div class="text-center py-8 text-slate-500 text-sm"><p>Rota Rotation and Shift Schedules are integrated with the <strong>Roster Planner</strong>.</p><p class="mt-2 text-xs">Use Roster Planner to assign shifts and weekly-off rotations.</p><button onclick="toggleModal('calendar-modal', false); toggleModal('roster-planner-modal', true)" class="mt-6 bg-indigo-600 hover:bg-indigo-700 text-white px-5 py-2.5 rounded-xl font-bold transition shadow-md">Open Roster Planner</button></div>
            {% endif %}
        </div>
    </div>

      <!-- Employee Card Directory Modal -->
      <div id="employee-list-modal" class="fixed inset-0 bg-slate-950/75 backdrop-blur-sm z-[65] flex items-center justify-center hidden p-2 sm:p-4">
          <div class="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-7xl max-h-[94vh] flex flex-col overflow-hidden">
              <div class="px-5 py-4 bg-gradient-to-r from-slate-950 via-slate-900 to-emerald-950 text-white flex justify-between items-center">
                  <div><h3 class="text-lg font-black">Employee ID Cards</h3><p class="text-[11px] text-slate-300 mt-1">All employee identification cards in one place</p></div>
                  <div class="flex items-center gap-2">
                      <button onclick="resetEmployeeFilters()" class="h-9 w-9 rounded-lg bg-white/10 hover:bg-white/20">▦</button>
                      <button onclick="toggleModal('add-employee-modal',true)" class="bg-rose-500 hover:bg-rose-600 text-white px-4 py-2 rounded-xl text-xs font-bold">＋ Add Employee</button>
                      <button onclick="toggleModal('employee-list-modal',false)" class="h-9 w-9 rounded-lg bg-white/10 hover:bg-white/20 font-bold">✕</button>
                  </div>
              </div>
              <div class="p-4 bg-slate-50 border-b border-slate-200 grid grid-cols-1 md:grid-cols-4 gap-3">
                  <input id="emp-filter-id" oninput="filterEmployeeCards()" placeholder="Employee ID" class="border border-slate-300 rounded-xl px-3 py-2.5 text-xs">
                  <input id="emp-filter-name" oninput="filterEmployeeCards()" placeholder="Employee Name" class="border border-slate-300 rounded-xl px-3 py-2.5 text-xs">
                  <select id="emp-filter-designation" onchange="filterEmployeeCards()" class="border border-slate-300 rounded-xl px-3 py-2.5 text-xs"><option value="">Select Designation</option>{% for d in employee_designations %}<option value="{{ d|lower }}">{{ d }}</option>{% endfor %}</select>
                  <button onclick="filterEmployeeCards()" class="bg-emerald-500 hover:bg-emerald-600 text-white rounded-xl text-xs font-bold">Search</button>
              </div>
              <div class="p-4 overflow-y-auto bg-slate-700 flex-1">
                  <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-5" id="employee-card-grid">
                      {% for emp in employee_cards %}
                      <article class="employee-directory-card rounded-2xl bg-slate-200 p-2 shadow-xl" data-id="{{emp.user_id|lower}}" data-name="{{emp.name|lower}}" data-designation="{{emp.designation|lower}}">
                        <div class="relative mx-auto h-[500px] w-full max-w-[340px] overflow-hidden rounded-xl bg-white text-center shadow-lg">
                          <div class="flex h-16 items-center justify-center px-5 pt-2"><img src="{{url_for('static',filename='fresmart.png')}}" class="h-12 w-44 object-contain" alt="Company Logo"></div>
                          <div class="mx-auto mt-1 h-[138px] w-[122px] overflow-hidden rounded-lg border bg-slate-50">
                            {% if emp.profile_photo %}<img src="/uploads/{{emp.profile_photo}}" class="h-full w-full object-cover" alt="Employee Photo">{% else %}<div class="flex h-full w-full items-center justify-center bg-gradient-to-br from-amber-300 to-rose-500 text-3xl font-black text-white">{{emp.name[:2]}}</div>{% endif %}
                          </div>
                          <h4 class="mx-auto mt-2 max-w-[290px] truncate px-2 text-base font-black uppercase text-slate-800" title="{{emp.name}}">{{emp.name}}</h4>
                          <div class="mx-5 mt-3 space-y-1 text-left text-[11px] text-slate-800">
                            <div class="grid grid-cols-[110px_1fr]"><span>HRMS Code</span><b>: {{emp.user_id}}</b></div>
                            <div class="grid grid-cols-[110px_1fr]"><span>Designation</span><b class="truncate" title="{{emp.designation}}">: {{emp.designation}}</b></div>
                            <div class="grid grid-cols-[110px_1fr]"><span>Identification</span><b class="truncate">: {{emp.identificacao or '-'}}</b></div>
                            <div class="grid grid-cols-[110px_1fr]"><span>Contract Date</span><b>: {{emp.data_de_contrato or '-'}}</b></div>
                            <div class="grid grid-cols-[110px_1fr]"><span>Store</span><b>: {{emp.stores|join(', ')}}</b></div>
                          </div>
                          <div class="absolute bottom-0 left-0 w-full"><div class="h-2 bg-gradient-to-r from-lime-400 via-amber-400 to-fuchsia-500"></div><div class="bg-black px-3 py-2 text-[9px] leading-3 text-white">Employee Identification Card<br>{{emp.user_id}}</div></div>
                        </div>
                        <div class="mt-2 grid grid-cols-2 gap-2"><a href="/employee_id/{{emp.user_id}}" target="_blank" class="rounded-lg bg-slate-900 py-2 text-center text-[10px] font-bold text-white">Open / Download ID Card</a><button onclick="toggleModal('user-mgmt-modal',true)" class="rounded-lg bg-emerald-600 py-2 text-[10px] font-bold text-white">Manage</button></div>
                        <form action="/admin_employee_photo/{{emp.user_id}}" method="POST" enctype="multipart/form-data" class="mt-2 flex gap-2 rounded-xl bg-white p-2"><input type="file" name="profile_photo" accept=".jpg,.jpeg,.png,.webp" required class="min-w-0 flex-1 text-[9px]"><button class="shrink-0 rounded-lg bg-amber-400 px-2 py-1 text-[9px] font-black">Change Photo</button></form>
                      </article>
                      {% else %}<div class="col-span-full py-12 text-center text-white/70">No employees found.</div>{% endfor %}
                  </div>
              </div>
          </div>
      </div>
      <!-- Add Employee Modal -->
      <div id="add-employee-modal" class="fixed inset-0 bg-slate-950/75 backdrop-blur-sm z-[70] flex items-center justify-center hidden p-3">
          <div class="bg-white rounded-2xl shadow-2xl w-full max-w-2xl p-5">
              <div class="flex justify-between items-center mb-4"><div><h3 class="font-black text-slate-900">Add Employee</h3><p class="text-xs text-slate-500">Create login and employee profile</p></div><button onclick="toggleModal('add-employee-modal',false)" class="text-xl">✕</button></div>
              <form action="/manage_user" method="POST" class="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <input type="hidden" name="action" value="create"><input type="hidden" name="role" value="employee"><input type="hidden" name="status" value="active"><input type="hidden" name="job_role" value="EMPLOYEE">
                  <input name="uid" placeholder="Employee ID, e.g. NWC9001" required class="border rounded-xl p-3 text-xs"><input name="name" placeholder="Full name" required class="border rounded-xl p-3 text-xs">
                  <input name="designation" placeholder="Designation" required class="border rounded-xl p-3 text-xs"><input name="dept" placeholder="Department" required class="border rounded-xl p-3 text-xs">
                  <input name="email" type="email" placeholder="Email address" class="border rounded-xl p-3 text-xs"><input name="identificacao" placeholder="Identificação / NIF" class="border rounded-xl p-3 text-xs"><input name="data_de_contrato" type="date" class="border rounded-xl p-3 text-xs"><input name="password" value="123" placeholder="Initial password" class="border rounded-xl p-3 text-xs">
                  <div class="sm:col-span-2 border rounded-xl p-3"><p class="text-[10px] font-bold text-slate-500 mb-2">STORE ACCESS</p>{% for code,m in machines.items() if code!='DEV' %}<label class="mr-4 text-xs"><input type="checkbox" name="stores" value="{{code}}" {% if code==store %}checked{% endif %}> {{code}}</label>{% endfor %}</div>
                  {% for p in role_presets.EMPLOYEE %}<input type="hidden" name="permissions" value="{{p}}">{% endfor %}
                  <button class="sm:col-span-2 bg-emerald-600 hover:bg-emerald-700 text-white font-bold py-3 rounded-xl text-xs">Create Employee</button>
              </form>
          </div>
      </div>
    <div id="profile-modal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[60] flex items-center justify-center hidden p-4"><div class="bg-white rounded-2xl shadow-2xl p-5 w-full max-w-md space-y-4"><div class="flex justify-between"><h3 class="font-bold">👤 My Profile</h3><button onclick="toggleModal('profile-modal',false)">✕</button></div><div class="text-center">{% if current_user.profile_photo %}<img src="/uploads/{{current_user.profile_photo}}" class="w-24 h-24 rounded-full object-cover mx-auto">{% else %}<div class="w-24 h-24 rounded-full bg-slate-100 flex items-center justify-center mx-auto text-3xl">👤</div>{% endif %}<p class="font-bold mt-2">{{logged_user_name}}</p><p class="text-xs text-slate-500">{{current_user.designation}} · {{current_user.job_role}}</p></div><form action="/upload_profile_photo" method="POST" enctype="multipart/form-data" class="space-y-2"><input type="file" name="profile_photo" accept=".jpg,.jpeg,.png,.webp" required class="w-full border rounded-lg p-2 text-xs"><button class="w-full bg-indigo-600 text-white py-2 rounded-lg font-bold text-xs">Upload / Change Photo</button></form>{% if current_user.profile_photo %}<form action="/remove_profile_photo" method="POST"><button class="w-full bg-rose-50 text-rose-700 py-2 rounded-lg font-bold text-xs">Remove Photo</button></form>{% endif %}</div></div>

    <!-- Leave Management Modal -->
    <div id="leave-modal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[60] flex items-center justify-center hidden p-2"><div class="bg-white rounded-2xl shadow-2xl p-4 w-full max-w-5xl max-h-[90vh] flex flex-col space-y-4"><div class="flex justify-between"><h3 class="font-bold">🏖️ Leave Management</h3><button onclick="toggleModal('leave-modal',false)">✕</button></div>{% if role=='employee' %}<form action="/apply_leave" method="POST" enctype="multipart/form-data" class="grid grid-cols-1 sm:grid-cols-5 gap-2 bg-emerald-50 p-3 rounded-xl"><input type="date" name="start_date" required class="border rounded p-2 text-xs"><input type="date" name="end_date" required class="border rounded p-2 text-xs"><select name="leave_type" class="border rounded p-2 text-xs" required>
<option value="">Select Leave Code / Motivo</option>
<option value="F01;1">F01;1 — Baixa Médica</option>
<option value="F03;1">F03;1 — Falta Injustificada</option>
<option value="F05;1">F05;1 — Licença sem vencimento</option>
<option value="F10;1">F10;1 — Falta Justificada</option>
<option value="F51;1">F51;1 — Casamento</option>
<option value="F60;1">F60;1 — Nascimento</option>
<option value="F61;1">F61;1 — Obito</option>
<option value="F62;1">F62;1 — Gravidez</option>
</select><input type="file" name="supporting_doc" accept=".pdf,.jpg,.jpeg,.png,.doc,.docx" class="border rounded p-1 text-[10px] bg-white"><button class="bg-emerald-600 text-white font-bold rounded text-xs">Submit Leave</button></form>{% endif %}<div class="overflow-auto flex-1 border rounded-xl"><table class="w-full text-left min-w-[650px] text-[10px]"><thead class="bg-slate-100 font-bold"><tr><th class="p-2">Employee</th><th class="p-2">Dates</th><th class="p-2">Type</th><th class="p-2">Status</th><th class="p-2">Action</th></tr></thead><tbody>{% for req in leave_requests|reverse %}<tr class="border-b"><td class="p-2 font-bold">{{req.name}}<br><span class="text-slate-400">{{req.user_id}} · {{req.get('store','')}}</span></td><td class="p-2">{{req.start_date}} → {{req.end_date}}</td><td class="p-2"><b>{{ req.leave_type }}</b><br><span class="text-slate-500">{{ leave_code_names.get(req.leave_type, req.get('leave_reason', req.leave_type)) }}</span></td><td class="p-2">{{req.status}}</td><td class="p-2">{% if req.status=='Pending' and role in ['admin','developer'] %}<a href="/update_leave/{{req.id}}/approve" class="bg-emerald-500 text-white px-2 py-1 rounded mr-1">Approve</a><a href="/update_leave/{{req.id}}/reject" class="bg-rose-500 text-white px-2 py-1 rounded">Reject</a>{% else %}<span class="text-slate-400">Processed</span>{% endif %}</td></tr>{% else %}<tr><td colspan="5" class="p-6 text-center text-slate-400">No leave requests</td></tr>{% endfor %}</tbody></table></div></div></div>

      <div id="overtime-approval-modal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[70] flex items-center justify-center hidden p-2">
        <div class="bg-white rounded-2xl shadow-2xl p-4 w-full max-w-5xl max-h-[90vh] flex flex-col gap-3">
          <div class="flex justify-between"><div><h3 class="font-black">⏱️ Overtime Approval</h3><p class="text-xs text-slate-500">Only pending overtime requests are shown here.</p></div><button onclick="toggleModal('overtime-approval-modal',false)">✕</button></div>
          <div class="overflow-auto border rounded-xl"><table class="w-full min-w-[760px] text-xs"><thead class="bg-slate-100"><tr><th>Date</th><th>Employee</th><th>Store</th><th>Extra Time</th><th>Status</th><th>Action</th></tr></thead><tbody>
          {% for ot in overtime_rows|reverse if ot.status=='Pending' %}<tr class="border-b"><td>{{ot.date}}</td><td><b>{{ot.name}}</b><br>{{ot.user_id}}</td><td>{{ot.store}}</td><td>{{ot.minutes//60}}h {{ot.minutes%60}}m</td><td><span class="bg-amber-100 text-amber-700 px-2 py-1 rounded-full">{{ot.status}}</span></td><td><a class="bg-emerald-600 text-white px-3 py-2 rounded-lg mr-1" href="/overtime_action/{{ot.id}}/approve">Approve</a><a class="bg-rose-600 text-white px-3 py-2 rounded-lg" href="/overtime_action/{{ot.id}}/reject">Reject</a></td></tr>{% else %}<tr><td colspan="6" class="p-6 text-center text-slate-400">No pending overtime requests</td></tr>{% endfor %}
          </tbody></table></div>
        </div>
      </div>

      <div id="overtime-history-modal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[70] flex items-center justify-center hidden p-2">
        <div class="bg-white rounded-2xl shadow-2xl p-4 w-full max-w-6xl max-h-[92vh] flex flex-col gap-3">
          <div class="flex justify-between"><div><h3 class="font-black">🕘 Overtime History</h3><p class="text-xs text-slate-500">Complete employee-wise Pending, Approved and Rejected history.</p></div><button onclick="toggleModal('overtime-history-modal',false)">✕</button></div>
          <div class="grid grid-cols-1 sm:grid-cols-3 gap-2"><input id="overtime-history-search" oninput="filterOvertimeHistory()" placeholder="Search employee name or code" class="border rounded-xl px-3 py-2 text-xs"><select id="overtime-history-status" onchange="filterOvertimeHistory()" class="border rounded-xl px-3 py-2 text-xs"><option value="">All Status</option><option>Pending</option><option>Approved</option><option>Rejected</option></select><button onclick="document.getElementById('overtime-history-search').value='';document.getElementById('overtime-history-status').value='';filterOvertimeHistory()" class="border rounded-xl px-3 py-2 text-xs font-bold">Clear Filters</button></div>
          <div class="overflow-auto border rounded-xl"><table id="overtime-history-table" class="w-full min-w-[900px] text-xs"><thead class="bg-slate-100"><tr><th>Date</th><th>Employee</th><th>Store</th><th>Detected Extra</th><th>Status</th><th>Reviewed By</th><th>Reviewed At</th></tr></thead><tbody>
          {% for ot in overtime_rows|reverse %}<tr class="border-b overtime-history-row" data-search="{{ot.name|lower}} {{ot.user_id|lower}}" data-status="{{ot.status}}"><td>{{ot.date}}</td><td><b>{{ot.name}}</b><br>{{ot.user_id}}</td><td>{{ot.store}}</td><td>{{ot.minutes//60}}h {{ot.minutes%60}}m</td><td>{{ot.status}}</td><td>{{ot.reviewed_by or '-'}}</td><td>{{ot.reviewed_at or '-'}}</td></tr>{% else %}<tr><td colspan="7" class="p-6 text-center text-slate-400">No overtime history</td></tr>{% endfor %}
          </tbody></table></div>
        </div>
      </div>


    <div id="permission-control-modal" class="fixed inset-0 bg-slate-950/75 backdrop-blur-sm z-[90] hidden p-2 sm:p-4">
      <div class="mx-auto flex h-full w-full max-w-7xl flex-col overflow-hidden rounded-2xl bg-white shadow-2xl">
        <div class="flex items-center justify-between bg-gradient-to-r from-violet-950 to-slate-950 px-5 py-4 text-white"><div><h3 class="text-lg font-black">🔐 Portal Rights Control</h3><p class="text-xs text-violet-200">Developer can allow or hide every portal option for existing and new users.</p></div><button onclick="toggleModal('permission-control-modal',false)" class="rounded-lg bg-white/10 px-3 py-2">✕</button></div>
        <div class="overflow-auto p-4 space-y-4">
          {% for uid,raw in users_db.items() if uid != session.get('user_id') %}{% set u=raw %}
          <form action="/developer_user_rights/{{uid}}" method="POST" class="rounded-2xl border border-slate-200 bg-slate-50 p-4 shadow-sm">
            <div class="mb-3 flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between"><div><h4 class="font-black text-slate-900">{{u.name}} <span class="font-mono text-xs text-slate-500">{{uid}}</span></h4><p class="text-xs text-slate-500">Current role: {{u.role}} · Store: {{u.store}} · Updated: {{u.get('rights_updated_at','-')}}</p></div><div class="flex gap-2"><select name="role" class="rounded-lg border px-3 py-2 text-xs"><option value="employee" {% if u.role=='employee' %}selected{% endif %}>Employee</option><option value="admin" {% if u.role=='admin' %}selected{% endif %}>Admin</option></select><button type="button" onclick="setRights(this,true)" class="rounded-lg border px-3 py-2 text-xs font-bold">Select All</button><button type="button" onclick="setRights(this,false)" class="rounded-lg border px-3 py-2 text-xs font-bold">Clear All</button><button class="rounded-lg bg-violet-600 px-4 py-2 text-xs font-black text-white">Save Rights</button></div></div>
            <div class="mb-3 flex flex-wrap gap-3 rounded-xl bg-white p-3"><b class="text-xs">Store Access:</b>{% for code,m in machines.items() if code!='DEV' %}<label class="text-xs"><input type="checkbox" name="stores" value="{{code}}" {% if code in u.get('stores',[u.get('store')]) %}checked{% endif %}> {{code}}</label>{% endfor %}</div>
            <div class="rights-grid grid grid-cols-1 gap-2 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4">{% for p in permission_list %}<label class="flex items-center gap-2 rounded-xl border bg-white p-2 text-xs"><input type="checkbox" name="permissions" value="{{p}}" {% if p in u.get('permissions',[]) %}checked{% endif %}><span>{{permission_labels.get(p,p)}}</span></label>{% endfor %}</div>
          </form>{% endfor %}
        </div>
      </div>
    </div>
    <!-- Password Resets Modal (Admin & Dev) -->
    <div id="reset-approvals-modal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[60] flex items-center justify-center hidden p-2">
        <div class="bg-white rounded-2xl shadow-2xl border border-slate-200 p-4 w-full max-w-4xl mx-auto space-y-4 max-h-[85vh] flex flex-col">
            <div class="flex justify-between items-center border-b border-slate-100 pb-2">
                <h3 class="text-sm sm:text-lg font-bold text-slate-900">🔑 Manage Passwords</h3>
                <button onclick="toggleModal('reset-approvals-modal', false)" class="text-slate-400 hover:text-slate-600 font-bold text-lg">✕</button>
            </div>
            
            {% if role == 'developer' %}
            <div class="p-3 border border-emerald-200 bg-emerald-50 rounded-xl">
                <h4 class="text-xs font-bold text-emerald-800 mb-2">Direct Reset (Admin Override)</h4>
                <form action="/dev_force_reset" method="POST" class="flex flex-col sm:flex-row gap-2">
                    <input type="text" name="target_user_id" placeholder="User ID (e.g. NWC1234)" required class="flex-1 px-3 py-2 text-xs border border-slate-300 rounded-lg focus:outline-none">
                    <input type="text" name="target_new_password" placeholder="New Password" required class="flex-1 px-3 py-2 text-xs border border-slate-300 rounded-lg focus:outline-none">
                    <button type="submit" class="bg-emerald-600 hover:bg-emerald-700 text-white font-bold px-4 py-2 rounded-lg text-xs uppercase transition">Set</button>
                </form>
            </div>
            {% endif %}

            <div class="overflow-auto flex-1 border border-slate-200 rounded-xl">
                <table class="w-full text-left min-w-[500px]">
                    <thead class="bg-slate-100 text-slate-600 uppercase text-[9px] font-bold sticky top-0">
                        <tr>
                            <th class="py-2 px-2">Emp</th>
                            <th class="py-2 px-2">Requested Pwd</th>
                            {% if role == 'developer' %}
                            <th class="py-2 px-2">Timeline</th>
                            <th class="py-2 px-2">Approved By</th>
                            {% endif %}
                            <th class="py-2 px-2 text-center">Status</th>
                            <th class="py-2 px-2 text-center">Action</th>
                        </tr>
                    </thead>
                    <tbody class="divide-y divide-slate-100 text-[10px] sm:text-xs">
                        {% if reset_requests %}
                            {% for req in reset_requests|reverse %}
                            <tr class="hover:bg-slate-50">
                                <td class="py-2 px-2 font-bold">{{ req.name }} <br><span class="text-[9px] text-slate-400">{{ req.user_id }}</span></td>
                                <td class="py-2 px-2 font-mono text-emerald-600">{{ req.new_password }}</td>
                                
                                {% if role == 'developer' %}
                                <td class="py-2 px-2 text-[9px] text-slate-500 whitespace-nowrap">
                                    Req: {{ req.get('request_date', 'N/A') }}<br>
                                    Appr: <span class="font-bold text-slate-700">{{ req.get('approve_date', '-') }}</span>
                                </td>
                                <td class="py-2 px-2 font-bold text-slate-700 text-[10px]">{{ req.get('approved_by', '-') }}</td>
                                {% endif %}
                                
                                <td class="py-2 px-2 text-center">
                                    {% if req.status == 'Pending' %}<span class="bg-amber-100 text-amber-700 px-1.5 py-0.5 rounded">Pending</span>
                                    {% elif req.status == 'Approved' %}<span class="bg-emerald-100 text-emerald-700 px-1.5 py-0.5 rounded">Approved</span>
                                    {% else %}<span class="bg-rose-100 text-rose-700 px-1.5 py-0.5 rounded">Rejected</span>{% endif %}
                                </td>
                                <td class="py-2 px-2 text-center">
                                    {% if req.status == 'Pending' %}
                                        {% if role == 'developer' or (role == 'admin' and req.get('role') != 'admin') %}
                                        <div class="flex flex-col sm:flex-row justify-center gap-1">
                                            <a href="/update_password_req/{{ req.id }}/approve" class="bg-emerald-500 text-white px-2 py-1 rounded">Approve</a>
                                            <a href="/update_password_req/{{ req.id }}/reject" class="bg-rose-500 text-white px-2 py-1 rounded">Reject</a>
                                        </div>
                                        {% else %}
                                        <span class="text-slate-400 text-[9px] font-bold">Dev Approval Only</span>
                                        {% endif %}
                                    {% else %}
                                        <span class="text-slate-400">Processed</span>
                                    {% endif %}
                                </td>
                            </tr>
                            {% endfor %}
                        {% else %}
                            <tr><td colspan="6" class="text-center py-4 text-slate-400">No requests</td></tr>
                        {% endif %}
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- Monthly Roster Matrix Planner -->
    <div id="roster-planner-modal" class="fixed inset-0 bg-slate-950/70 backdrop-blur-sm z-[60] flex items-center justify-center hidden p-2">
        <div class="bg-white rounded-2xl shadow-2xl border border-slate-200 p-4 w-full max-w-[98vw] h-[94vh] flex flex-col gap-3">
            <div class="flex flex-col md:flex-row md:items-center justify-between gap-3 border-b pb-3">
                <div><h3 class="font-black text-lg">📅 Monthly Roster Matrix</h3><p class="text-[11px] text-slate-500">Employee-wise Shift A, Shift B and Weekly Off allocation</p><span class="inline-block mt-1 rounded-full bg-cyan-50 text-cyan-700 border border-cyan-200 px-2 py-1 text-[9px] font-bold">{% if role == 'developer' %}Developer View: All Stores{% else %}Store View: {{ store }} Only{% endif %}</span></div>
                <div class="flex flex-wrap items-end gap-2"><div><label class="block text-[9px] font-bold uppercase text-slate-500">Month</label><input type="month" id="rp-month" class="border rounded-xl px-3 py-2 text-xs"></div><button onclick="loadRosterPlanner()" class="bg-indigo-600 text-white font-bold text-xs px-5 py-2.5 rounded-xl">Load Month</button><button id="save-roster-btn" onclick="saveRoster()" class="bg-emerald-600 text-white font-bold text-xs px-5 py-2.5 rounded-xl hidden">Save Monthly Roster</button><button onclick="exportRosterExcel()" class="bg-amber-500 hover:bg-amber-600 text-white font-bold text-xs px-5 py-2.5 rounded-xl">Export Excel</button><button onclick="toggleModal('roster-planner-modal', false)" class="border px-4 py-2.5 rounded-xl text-xs font-bold">Close</button></div>
            </div>
            <div class="grid grid-cols-1 sm:grid-cols-4 gap-2 bg-slate-50 border border-slate-200 rounded-xl p-3">
                <input id="roster-search" oninput="filterRosterRows()" placeholder="Search code or employee name" class="border rounded-xl px-3 py-2 text-xs">
                <div class="relative">
                    <button type="button" onclick="toggleDesignationMenu()" id="roster-designation-button" class="w-full border bg-white rounded-xl px-3 py-2 text-xs text-left flex justify-between"><span>All Designations</span><b>⌄</b></button>
                    <div id="roster-designation-menu" class="hidden absolute z-50 mt-1 w-full max-h-56 overflow-auto bg-white border border-slate-200 rounded-xl shadow-xl p-2"></div>
                </div>
                <select id="roster-allocation-filter" onchange="filterRosterRows()" class="border rounded-xl px-3 py-2 text-xs"><option value="">All Allocations</option><option>Shift A</option><option>Shift B</option><option>Weekly Off</option></select>
                <button onclick="resetRosterFilters()" class="border border-slate-300 bg-white hover:bg-slate-100 rounded-xl px-3 py-2 text-xs font-bold">Reset Filters</button>
            </div>
            <div id="roster-unsaved-notice" class="hidden rounded-xl border border-amber-200 bg-amber-50 px-4 py-2 text-[10px] font-bold text-amber-800"></div>
            <div id="roster-matrix-table" class="overflow-auto flex-1 border border-slate-200 rounded-xl"><div class="py-12 text-center text-slate-400">Month select karke Load Month click karein</div></div>
        </div>
    </div>

    <!-- Apply Shift / Weekly Off to Week Range -->
    <div id="week-apply-modal" class="fixed inset-0 bg-slate-950/70 backdrop-blur-sm z-[80] flex items-center justify-center hidden p-3">
        <div class="bg-white rounded-2xl shadow-2xl w-full max-w-lg p-5">
            <div class="flex justify-between items-start"><div><h3 id="week-apply-title" class="font-black text-slate-900">Apply Roster</h3><p id="week-apply-subtitle" class="text-xs text-slate-500 mt-1"></p><span id="week-selection-count" class="inline-block mt-2 rounded-full bg-emerald-100 text-emerald-700 px-3 py-1 text-[10px] font-bold">1 option selected</span></div><button onclick="cancelRosterWeekApply()" class="text-xl">✕</button></div>
            <div class="grid grid-cols-1 sm:grid-cols-2 gap-2 mt-5 text-xs">
                <label class="border rounded-xl p-3 hover:bg-slate-50"><input type="checkbox" name="week-apply-option" onchange="handleWeekOptionChange(this)" value="selected_date"> Selected Date Only</label>
                <label class="border rounded-xl p-3 hover:bg-slate-50"><input type="checkbox" name="week-apply-option" onchange="handleWeekOptionChange(this)" value="week_1"> 1st Week (1-7)</label>
                <label class="border rounded-xl p-3 hover:bg-slate-50"><input type="checkbox" name="week-apply-option" onchange="handleWeekOptionChange(this)" value="week_2"> 2nd Week (8-14)</label>
                <label class="border rounded-xl p-3 hover:bg-slate-50"><input type="checkbox" name="week-apply-option" onchange="handleWeekOptionChange(this)" value="week_3"> 3rd Week (15-21)</label>
                <label class="border rounded-xl p-3 hover:bg-slate-50"><input type="checkbox" name="week-apply-option" onchange="handleWeekOptionChange(this)" value="week_4"> 4th Week (22-28)</label>
                <label class="border rounded-xl p-3 hover:bg-slate-50"><input type="checkbox" name="week-apply-option" onchange="handleWeekOptionChange(this)" value="last_week"> Last Week (29-End)</label>
                <label class="border border-emerald-200 bg-emerald-50 rounded-xl p-3"><input type="checkbox" id="all-weeks-checkbox" onchange="toggleAllRosterWeeks(this)"> All Weeks (1-End)</label>
                <label class="sm:col-span-2 border border-indigo-200 bg-indigo-50 rounded-xl p-3"><input type="checkbox" name="week-apply-option" onchange="handleWeekOptionChange(this)" value="same_weekday_all"> Same Weekday for Full Month</label>
            </div>
            <div class="flex justify-end gap-2 mt-5"><button onclick="cancelRosterWeekApply()" class="border rounded-xl px-4 py-2 text-xs font-bold">Cancel</button><button onclick="applyRosterWeekChoice()" class="bg-emerald-600 text-white rounded-xl px-5 py-2 text-xs font-bold">Apply</button></div>
        </div>
    </div>

    <!-- Employee Shift/Off Request Modal -->
    <div id="shift-request-modal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[60] flex items-center justify-center hidden p-4">
        <div class="bg-white rounded-2xl shadow-2xl border border-slate-200 p-5 w-full max-w-md mx-auto space-y-4">
            <div class="flex justify-between items-center border-b border-slate-100 pb-2">
                <h3 class="text-sm sm:text-base font-bold text-slate-900">🔄 Request Shift Change</h3>
                <button onclick="toggleModal('shift-request-modal', false)" class="text-slate-400 hover:text-slate-600 font-bold text-lg">✕</button>
            </div>
            <form action="/request_shift" method="POST" class="space-y-3">
                <div>
                    <label class="block text-xs font-bold text-slate-500 mb-1">Date</label>
                    <input type="date" name="req_date" required class="w-full border rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500">
                </div>
                <div>
                    <label class="block text-xs font-bold text-slate-500 mb-1">Requested Action</label>
                    <select name="req_status" required class="w-full border rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500">
                        <option value="Shift A">Change to Shift A (06:00)</option>
                        <option value="Shift B">Change to Shift B (10:00+)</option>
                        <option value="Weekly Off">Set as Weekly Off</option>
                    </select>
                </div>
                <div>
                    <label class="block text-xs font-bold text-slate-500 mb-1">Reason</label>
                    <input type="text" name="req_reason" placeholder="Reason..." required class="w-full border rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500">
                </div>
                <button type="submit" class="w-full bg-indigo-600 text-white font-bold text-xs uppercase py-2.5 rounded-xl">Submit</button>
            </form>
        </div>
    </div>

    <!-- Admin Shift Approvals Modal -->
    <div id="shift-approvals-modal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[60] flex items-center justify-center hidden p-2">
        <div class="bg-white rounded-2xl shadow-2xl border border-slate-200 p-4 w-full max-w-4xl mx-auto space-y-4 max-h-[85vh] flex flex-col">
            <div class="flex justify-between items-center border-b border-slate-100 pb-2">
                <h3 class="text-sm sm:text-lg font-bold text-slate-900">🔔 Shift & Off Requests</h3>
                <button onclick="toggleModal('shift-approvals-modal', false)" class="text-slate-400 hover:text-slate-600 font-bold text-lg">✕</button>
            </div>
            <div class="overflow-auto flex-1 border border-slate-200 rounded-xl">
                <table class="w-full text-left min-w-[500px]">
                    <thead class="bg-slate-100 text-slate-600 uppercase text-[9px] sm:text-[10px] font-bold">
                        <tr>
                            <th class="py-2 px-2">Emp</th>
                            <th class="py-2 px-2">Date</th>
                            <th class="py-2 px-2">Request</th>
                            <th class="py-2 px-2">Reason</th>
                            <th class="py-2 px-2 text-center">Status</th>
                            <th class="py-2 px-2 text-center">Action</th>
                        </tr>
                    </thead>
                    <tbody class="divide-y divide-slate-100 text-[10px] sm:text-xs">
                        {% if shift_requests %}
                            {% for req in shift_requests|reverse %}
                            <tr class="hover:bg-slate-50">
                                <td class="py-2 px-2 font-bold">{{ req.name }} <br><span class="text-[9px] text-slate-400">{{ req.user_id }}</span></td>
                                <td class="py-2 px-2 font-mono">{{ req.date }}</td>
                                <td class="py-2 px-2 font-bold text-indigo-700">{{ req.requested }}</td>
                                <td class="py-2 px-2 text-slate-500 italic">{{ req.reason }}</td>
                                <td class="py-2 px-2 text-center">
                                    {% if req.status == 'Pending' %}<span class="bg-amber-100 text-amber-700 px-1.5 py-0.5 rounded">Pending</span>
                                    {% elif req.status == 'Approved' %}<span class="bg-emerald-100 text-emerald-700 px-1.5 py-0.5 rounded">Approved</span>
                                    {% else %}<span class="bg-rose-100 text-rose-700 px-1.5 py-0.5 rounded">Rejected</span>{% endif %}
                                </td>
                                <td class="py-2 px-2 text-center">
                                    {% if req.status == 'Pending' %}
                                    <div class="flex gap-1 justify-center">
                                        <a href="/update_shift_req/{{ req.id }}/approve" class="bg-emerald-500 text-white px-2 py-1 rounded">Ok</a>
                                        <a href="/update_shift_req/{{ req.id }}/reject" class="bg-rose-500 text-white px-2 py-1 rounded">No</a>
                                    </div>
                                    {% else %}<span class="text-slate-400">Done</span>{% endif %}
                                </td>
                            </tr>
                            {% endfor %}
                        {% else %}
                            <tr><td colspan="6" class="text-center py-4 text-slate-400">No requests</td></tr>
                        {% endif %}
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- Developer User Management Modal -->
    <div id="user-mgmt-modal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[60] flex items-center justify-center hidden p-2"><div class="bg-white rounded-2xl shadow-2xl p-4 w-full max-w-6xl max-h-[92vh] flex flex-col space-y-3"><div class="flex justify-between"><h3 class="font-bold">🪪 User / Role / Store Rights</h3><button onclick="toggleModal('user-mgmt-modal',false)">✕</button></div><form action="/manage_user" method="POST" class="grid grid-cols-1 md:grid-cols-4 gap-2 bg-slate-50 p-3 rounded-xl"><input type="hidden" name="action" value="create"><input name="uid" placeholder="User ID" required class="border rounded p-2 text-xs"><input name="name" placeholder="Name" required class="border rounded p-2 text-xs"><input name="password" placeholder="Password" class="border rounded p-2 text-xs"><input name="designation" placeholder="Designation" class="border rounded p-2 text-xs"><input name="identificacao" placeholder="Identificação / NIF" class="border rounded p-2 text-xs"><input name="data_de_contrato" type="date" class="border rounded p-2 text-xs"><select name="role" class="border rounded p-2 text-xs"><option value="employee">Employee</option><option value="admin">Admin</option></select><input name="job_role" placeholder="HR / AREA MANAGER / OPERATION HEAD" class="border rounded p-2 text-xs"><input name="dept" placeholder="Department" class="border rounded p-2 text-xs"><select name="status" class="border rounded p-2 text-xs"><option value="active">Active</option><option value="blocked">Deactive / Blocked</option></select><div class="md:col-span-2"><b class="text-[10px]">Stores:</b>{% for code,m in machines.items() if code!='DEV' %}<label class="ml-2 text-[10px]"><input type="checkbox" name="stores" value="{{code}}">{{code}}</label>{% endfor %}<label class="ml-2 text-[10px] font-bold"><input type="checkbox" name="all_stores" value="1"> ALL</label></div><div class="md:col-span-2 grid grid-cols-2 sm:grid-cols-3 gap-1 max-h-16 overflow-auto">{% for p in permission_list %}<label class="text-[9px]"><input type="checkbox" name="permissions" value="{{p}}">{{p|replace('_',' ')|title}}</label>{% endfor %}</div><button class="md:col-span-4 bg-emerald-600 text-white font-bold py-2 rounded text-xs">Create / Update</button></form><div class="overflow-auto flex-1 border rounded-xl"><table class="w-full min-w-[950px] text-left text-[10px]"><thead class="bg-slate-100 sticky top-0"><tr><th class="p-2">ID / Name</th><th>Designation / Role</th><th>Stores</th><th>Status</th><th>Permissions</th><th>Action</th></tr></thead><tbody>{% for uid,info in users_db.items() %}<tr class="border-b"><td class="p-2 font-bold">{{uid}}<br>{{info.name}}</td><td class="p-2">{{info.designation}}<br>{{info.job_role}} / {{info.role}}</td><td class="p-2">{{info.stores|join(', ')}}</td><td class="p-2">{{info.status|upper}}</td><td class="p-2">{{info.permissions|join(', ')}}</td><td class="p-2"><div class="flex gap-1"><form action="/manage_user" method="POST"><input type="hidden" name="uid" value="{{uid}}"><input type="hidden" name="action" value="toggle_status"><button class="bg-amber-100 text-amber-700 px-2 py-1 rounded">Toggle</button></form><form action="/manage_user" method="POST" onsubmit="return confirm('Delete this user?')"><input type="hidden" name="uid" value="{{uid}}"><input type="hidden" name="action" value="delete"><button class="bg-rose-100 text-rose-700 px-2 py-1 rounded">Del</button></form></div></td></tr>{% endfor %}</tbody></table></div></div></div>

    <!-- Biometric Machine Management -->
    <div id="machine-mgmt-modal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[60] flex items-center justify-center hidden p-2"><div class="bg-white rounded-2xl shadow-2xl p-4 w-full max-w-5xl max-h-[90vh] overflow-auto"><div class="flex justify-between mb-3"><h3 class="font-bold">🖥️ Biometric Machines / Stores</h3><button onclick="toggleModal('machine-mgmt-modal',false)">✕</button></div><form action="/manage_machine" method="POST" class="grid grid-cols-1 md:grid-cols-5 gap-2 bg-slate-50 p-3 rounded-xl"><input name="code" placeholder="Store Code" required class="border rounded p-2 text-xs"><input name="name" placeholder="Portal Name" required class="border rounded p-2 text-xs"><input name="ip" placeholder="Machine IP" required class="border rounded p-2 text-xs"><input name="port" value="4370" required class="border rounded p-2 text-xs"><input name="admin" placeholder="Portal Admin User ID" class="border rounded p-2 text-xs"><button class="md:col-span-5 bg-indigo-600 text-white font-bold py-2 rounded text-xs">Save / Merge Machine</button></form><table class="w-full text-left text-xs mt-3"><thead class="bg-slate-100"><tr><th class="p-2">Store</th><th class="p-2">Portal</th><th class="p-2">Address</th><th class="p-2">Admin</th><th class="p-2">Status</th><th></th></tr></thead><tbody>{% for code,m in machines.items() %}<tr class="border-b"><td class="p-2 font-bold">{{code}}</td><td class="p-2">{{m.name}}</td><td class="p-2 font-mono">{{m.ip}}:{{m.port}}</td><td class="p-2">{{m.get('admin','-')}}</td><td class="p-2">{% if machine_status.get(code) %}<span class="text-emerald-600 font-bold">ONLINE</span>{% else %}<span class="text-rose-600 font-bold">OFFLINE</span>{% endif %}</td><td class="p-2">{% if code not in ['LM11','LF07','DEV'] %}<form action="/manage_machine" method="POST"><input type="hidden" name="code" value="{{code}}"><input type="hidden" name="action" value="delete"><button class="text-rose-600 font-bold" onclick="return confirm('Delete machine?')">Delete</button></form>{% endif %}</td></tr>{% endfor %}</tbody></table></div></div>


<script>
window.HRMS_LANGUAGE={{ ui_language|default('en')|tojson }};
window.HRMS_PT={'Correction Rejected':'Correcção Rejeitada','Correction Approved':'Correcção Aprovada','Correction Request':'Pedido de Correcção','New Document':'Novo Documento','was Rejected':'foi rejeitada','was Approved':'foi aprovada','Your correction':'A sua correcção','Your roster':'A sua escala','Month select karke Load Month click karein':'Seleccione o mês e clique em Carregar Mês','Employee-wise Shift A, Shift B and Weekly Off allocation':'Alocação por trabalhador de Turno A, Turno B e Folga Semanal','Public Holiday':'Feriado Público','Half Day':'Meio Dia','Full Day':'Dia Completo','Full Month':'Mês Completo','Clear All':'Limpar Tudo','Select All':'Seleccionar Tudo','Uploaded By':'Carregado por','Updated By':'Actualizado por','Created On':'Criado em','Created By':'Criado por','Records':'Registos','Record':'Registo','Summary':'Resumo','Details':'Detalhes','Code':'Código','Address':'Endereço','Phone':'Telefone','Email':'E-mail','Photo':'Fotografia','Store Access':'Acesso à Loja','Apply Selected':'Aplicar Seleccionados','Copy Previous Month':'Copiar Mês Anterior','Copy Previous Week':'Copiar Semana Anterior','Monthly roster saved':'Escala mensal guardada','Kam se kam ek option select karein.':'Seleccione pelo menos uma opção.','Month select karein':'Seleccione o mês','Galat User ID ya Password!':'ID do utilizador ou palavra-passe incorrectos!','Sirf PDF files allowed hain!':'Apenas ficheiros PDF são permitidos!','Sabhi fields bharna zaroori hai!':'Todos os campos são obrigatórios!','This roster shows only your shifts and weekly-off dates for the current month.':'Esta escala mostra apenas os seus turnos e folgas semanais do mês actual.','Overtime Minutes':'Minutos de Horas Extra','Submit Overtime':'Enviar Horas Extra','Email - requires SMTP':'E-mail - requer SMTP','Internal Notification':'Notificação Interna','Download Center':'Centro de Transferências','Monthly':'Mensal','Weekly':'Semanal','Daily':'Diário','Delivery':'Entrega','Frequency':'Frequência','File':'Ficheiro','Uploaded At':'Carregado em','Created':'Criado','To':'Para','From':'De','Effective':'Efectivo','Estimate':'Estimativa','Employer':'Empregador','Gross':'Bruto','Configurable estimate using 2026 settings. Validate payroll categories, taxable base and exemptions before finalisation.':'Estimativa configurável com definições de 2026. Valide as categorias salariais, a base tributável e as isenções antes da finalização.','Identification masking':'Ocultação da Identificação','Document access log':'Registo de Acesso a Documentos','Access audit':'Auditoria de Acesso','Last backups':'Últimas Cópias de Segurança','Number of Users':'Número de Utilizadores','Application Version':'Versão da Aplicação','retry required':'nova tentativa necessária','Last saved attendance remains available':'A última assiduidade guardada continua disponível','Queue State':'Estado da Fila','Level':'Nível','Late Count':'Número de Atrasos','Rolling 30-day view.':'Vista móvel de 30 dias.','roster cells copied':'células de escala copiadas','assignments saved':'alocações guardadas','Copy failed':'Falha ao copiar','Save Lifecycle Event':'Guardar Evento do Ciclo de Vida','Employee Lifecycle':'Ciclo de Vida do Trabalhador','New Records':'Novos Registos','Records Received':'Registos Recebidos','Machine Status':'Estado do Equipamento','Store Code':'Código da Loja','Payroll, holidays, lifecycle, overtime, backups and privacy controls.':'Salários, feriados, ciclo de vida, horas extra, cópias de segurança e controlos de privacidade.','Privacy-safe QR verification cards.':'Cartões QR de verificação com protecção de privacidade.','Internal leave, shift, payroll and contract notices.':'Avisos internos de ausências, turnos, salários e contratos.','Track user, store, action, IP and changes.':'Acompanhar utilizador, loja, acção, IP e alterações.','Minimum staffing rules and shortage warnings.':'Regras de dotação mínima e avisos de falta de pessoal.','Copy week or month schedules.':'Copiar escalas semanais ou mensais.','Contracts, IDs and certificates.':'Contratos, documentos de identificação e certificados.','Annual, used, pending and available balances.':'Saldos anuais, usados, pendentes e disponíveis.','Approve missing-punch correction requests.':'Aprovar pedidos de correcção de marcações em falta.','Machine status, manual checks and sync history.':'Estado dos equipamentos, verificações manuais e histórico de sincronização.','Combined attendance with store code and Excel export.':'Assiduidade combinada com código da loja e exportação Excel.','Multi-store operations, compliance and employee self-service':'Operações multi-loja, conformidade e auto-serviço do trabalhador','No contract alerts.':'Sem alertas de contrato.','Audit Events':'Eventos de Auditoria','Contracts':'Contratos','Corrections':'Correcções','Machine':'Equipamento','Store Comparison':'Comparação de Lojas','Monthly Summary':'Resumo Mensal','Default Schedule':'Horário Padrão','All Allocations':'Todas as Alocações','Designations Selected':'Funções Seleccionadas','All Designations Selected':'Todas as Funções Seleccionadas','Same Weekday for Full Month':'Mesmo Dia da Semana durante Todo o Mês','Monday-Sunday':'Segunda a Domingo','Last Week':'Última Semana','5th Week':'5.ª Semana','4th Week':'4.ª Semana','3rd Week':'3.ª Semana','2nd Week':'2.ª Semana','1st Week':'1.ª Semana','All Weeks (Monday-Sunday)':'Todas as Semanas (Segunda a Domingo)','Not in this month':'Não existe neste mês','options selected':'opções seleccionadas','option selected':'opção seleccionada','Apply to selected date':'Aplicar à data seleccionada','Month & Year':'Mês e Ano','load month':'carregar mês','Select Month':'Seleccionar Mês','No shift requests':'Sem pedidos de turno','Current Assignment':'Alocação Actual','Requested':'Solicitado','Shift Request':'Pedido de Turno','Request Shift / Off':'Pedir Turno / Folga','Select Employee':'Seleccionar Trabalhador','Bulk Upload':'Carregamento em Massa','Standard Row':'Linha Padrão','Monthly grid':'Grelha mensal','Master':'Principal','Individual Slip':'Recibo Individual','Split':'Separar','All Personnel':'Todo o Pessoal','No requests':'Sem pedidos','Minutes':'Minutos','ALL':'TODOS','Change':'Alteração','Emp':'Trab.','Remove Photo':'Remover Fotografia','STORE ACCESS':'ACESSO À LOJA','Create login and employee profile':'Criar início de sessão e perfil do trabalhador','No employees found.':'Nenhum trabalhador encontrado.','Select Designation':'Seleccionar Função','Add Employee':'Adicionar Trabalhador','Search, review and maintain workforce profiles':'Pesquisar, rever e manter perfis dos trabalhadores','Open Roster Planner':'Abrir Planeador de Escala','Use Roster Planner to assign shifts and weekly-off rotations.':'Utilize o Planeador de Escala para atribuir turnos e rotações de folga semanal.','Rota Rotation and Shift Schedules are integrated with the':'A rotação da escala e os horários de turno estão integrados com o','No salary slips uploaded yet.':'Ainda não foram carregados recibos de salário.','View':'Ver','PDF File':'Ficheiro PDF','-- Select Employee --':'-- Seleccionar Trabalhador --','Individual Slip Upload':'Carregar Recibo Individual','Split & Upload':'Separar e Carregar','Master PDF':'PDF Principal','Bulk Upload (Merged PDF)':'Carregamento em Massa (PDF Unido)','Monthly grid with Leave Codes':'Grelha mensal com códigos de ausência','Employee Matrix':'Matriz de Trabalhadores','Individual records per date':'Registos individuais por data','Standard Row Export':'Exportação em Linhas','Download Attendance Reports':'Baixar Relatórios de Assiduidade','Total Summary':'Resumo Total','No attendance records found for this selection.':'Nenhum registo de assiduidade encontrado para esta selecção.','Status ↕':'Estado ↕','Working Hrs ↕':'Horas Trabalhadas ↕','Total Lunch ↕':'Total de Almoço ↕','Out Time ↕':'Hora de Saída ↕','Store In ↕':'Entrada na Loja ↕','Dept ↕':'Departamento ↕','Employee Name ↕':'Nome do Trabalhador ↕','Date ↕':'Data ↕','Last Month':'Mês Passado','This Month':'Este Mês','This Week':'Esta Semana','Yesterday':'Ontem','Quick Range':'Intervalo Rápido','Rota':'Escala','Export':'Exportar','Logged In As':'Sessão Iniciada Como','-- All Personnel --':'-- Todo o Pessoal --','Employee Filter':'Filtro de Trabalhadores','Tot Hrs':'Total de Horas','Mis Punch':'Marcação em Falta','Mis-Punch':'Marcação em Falta','Late Arr.':'Chegada Tardia','Week Off':'Folga Semanal','Shutdown':'Encerrar','Sync':'Sincronização','Device':'Equipamento','Leaves':'Ausências','Dev':'Programador','Biometric live tracking active for Attendance Portal.':'Acompanhamento biométrico em tempo real activo no Portal de Assiduidade.','Announcements':'Comunicados','Employees Info (ID Card)':'Informações dos Trabalhadores (Cartão de Identificação)','Manage Passwords':'Gerir Palavras-passe','Biometric Machines':'Equipamentos Biométricos','Shift Approvals':'Aprovações de Turno','Reset':'Repor','Off':'Folga','Daily manpower summary':'Resumo diário de efectivos','Loading monthly roster...':'A carregar a escala mensal...','Send Request for Approval':'Enviar Pedido para Aprovação','Developed by':'Desenvolvido por','Attendance Portal':'Portal de Assiduidade','Sign in to access your dashboard':'Inicie sessão para aceder ao seu painel','User ID':'ID do Utilizador','Password':'Palavra-passe','Secure Login':'Iniciar Sessão','Forgot/Reset Password?':'Esqueceu/Redefinir Palavra-passe?','Reset Password':'Redefinir Palavra-passe','Create a new password request':'Criar um novo pedido de palavra-passe','Submit Reset Request':'Enviar Pedido de Redefinição','Back to Login':'Voltar ao Início de Sessão','Current Password':'Palavra-passe Actual','New Password':'Nova Palavra-passe','Confirm Password':'Confirmar Palavra-passe','Dashboard':'Painel','Good day':'Bom dia','Main Menu':'Menu Principal','Team Management':'Gestão da Equipa','Quick Actions':'Acções Rápidas','Search':'Pesquisar','Search code or employee name':'Pesquisar código ou nome do trabalhador','Filter':'Filtrar','Reset Filters':'Limpar Filtros','Close':'Fechar','Save':'Guardar','Cancel':'Cancelar','Apply':'Aplicar','Load':'Carregar','Load Month':'Carregar Mês','Export Excel':'Exportar Excel','English':'Inglês','Portuguese':'Português','Language':'Idioma','Employee':'Trabalhador','Employees':'Trabalhadores','Employee Code':'Código do Trabalhador','Employee Name':'Nome do Trabalhador','Employee ID':'ID do Trabalhador','Designation':'Função','Department':'Departamento','Store':'Loja','Stores':'Lojas','All Stores':'Todas as Lojas','All Designations':'Todas as Funções','All Allocations':'Todas as Alocações','All Weeks':'Todas as Semanas','Selected':'Seleccionado','Selected Date Only':'Apenas a Data Seleccionada','Same Weekday for Full Month':'Mesmo Dia da Semana em Todo o Mês','First Week':'Primeira Semana','Second Week':'Segunda Semana','Third Week':'Terceira Semana','Fourth Week':'Quarta Semana','Fifth Week':'Quinta Semana','Last Week':'Última Semana','Week calculation':'Cálculo da semana','Monday is the first day and Sunday is the last day':'Segunda-feira é o primeiro dia e Domingo é o último dia','Attendance':'Assiduidade','Attendance Correction':'Correcção de Assiduidade','Attendance Corrections':'Correcções de Assiduidade','Present':'Presente','Absent':'Ausente','Late':'Atrasado','Late Arrival':'Chegada Tardia','Early Departure':'Saída Antecipada','Weekly Off':'Folga Semanal','Shift A':'Turno A','Shift B':'Turno B','Default':'Padrão','Roster Planner':'Planeador de Escala','Monthly Roster':'Escala Mensal','Monthly Roster Matrix':'Matriz de Escala Mensal','Save Monthly Roster':'Guardar Escala Mensal','Daily manpower summary':'Resumo diário de efectivos','Shift Allocation':'Alocação de Turno','Date & Day':'Data e Dia','My Monthly Roster':'Minha Escala Mensal','Shift / Weekly Off':'Turno / Folga Semanal','Today':'Hoje','Scheduled':'Programado','Date':'Data','Day':'Dia','Month':'Mês','Year':'Ano','Hours':'Horas','Minutes':'Minutos','Time':'Hora','Store In':'Entrada na Loja','Lunch Out':'Saída para Almoço','Lunch In':'Regresso do Almoço','Out Time':'Hora de Saída','Working Hours':'Horas Trabalhadas','Total Hours':'Total de Horas','Lunch Hours':'Horas de Almoço','Variance':'Variação','Device':'Equipamento','Online':'Online','Offline':'Offline','Loading':'A carregar','No data':'Sem dados','No records found':'Nenhum registo encontrado','Leave Portal':'Portal de Ausências','Leave Management':'Gestão de Ausências','Leave Balances':'Saldos de Ausências','Leave Type':'Tipo de Ausência','Leave Code':'Código de Ausência','Submit Leave':'Enviar Pedido','Start Date':'Data Inicial','End Date':'Data Final','Supporting Document':'Documento Comprovativo','Status':'Estado','Action':'Acção','Pending':'Pendente','Approved':'Aprovado','Rejected':'Rejeitado','Processed':'Processado','Approve':'Aprovar','Reject':'Rejeitar','Reason':'Motivo','Request':'Pedido','Requests':'Pedidos','No leave requests':'Sem pedidos de ausência','Select Leave Code / Motivo':'Seleccionar Código de Ausência / Motivo','Payroll & Reports':'Salários e Relatórios','Payroll':'Salários','Salary':'Salário','Salary Slip':'Recibo de Salário','Gross Salary':'Salário Bruto','Net Salary':'Salário Líquido','Base Salary':'Salário Base','Meal Allowance':'Subsídio de Alimentação','Transport Allowance':'Subsídio de Transporte','Bonus':'Prémio','Overtime':'Horas Extra','Overtime Approval':'Aprovação de Horas Extra','Net estimate':'Estimativa Líquida','Calculate Estimate':'Calcular Estimativa','Reports':'Relatórios','Report':'Relatório','Upload Salary Slip':'Carregar Recibo de Salário','Download Salary Slip':'Baixar Recibo de Salário','Employee Documents':'Documentos do Trabalhador','My Documents':'Meus Documentos','Notifications':'Notificações','Notification Center':'Centro de Notificações','ID Card':'Cartão de Identificação','Download ID Card':'Baixar Cartão de Identificação','Baixar / Download ID Card':'Baixar Cartão de Identificação','Calendar & Rota':'Calendário e Escala','Employee List':'Lista de Trabalhadores','Employee Cards':'Cartões dos Trabalhadores','Employee Information':'Informações do Trabalhador','Profile':'Perfil','My Profile':'Meu Perfil','Nationality':'Nacionalidade','HRMS Code':'Código HRMS','Identification':'Identificação','Contract Date':'Data de Contrato','Data de Contrato':'Data de Contrato','Create / Update':'Criar / Actualizar','Create Employee':'Criar Trabalhador','Update Employee':'Actualizar Trabalhador','Manage':'Gerir','Edit':'Editar','Delete':'Eliminar','Open':'Abrir','Download':'Baixar','Upload':'Carregar','Document':'Documento','Documents':'Documentos','No documents':'Sem documentos','Settings':'Definições','Logout':'Sair','User Management':'Gestão de Utilizadores','Password Management':'Gestão de Palavras-passe','Machine Management':'Gestão de Equipamentos','Permissions':'Permissões','Role':'Perfil de Acesso','Admin':'Administrador','Developer':'Programador','Blocked':'Bloqueado','Active':'Activo','Inactive':'Inactivo','Create User':'Criar Utilizador','Update User':'Actualizar Utilizador','Store View':'Vista da Loja','Developer View':'Vista do Programador','All Stores View':'Vista de Todas as Lojas','Workforce Automation Hub':'Centro de Automação da Força de Trabalho','Multi-store operations, compliance and employee self-service':'Operações multi-loja, conformidade e auto-serviço do trabalhador','Contract Alerts':'Alertas de Contrato','All Stores Attendance':'Assiduidade de Todas as Lojas','Biometric Sync History':'Histórico de Sincronização Biométrica','Manual Machine Check':'Verificação Manual do Equipamento','Check Now':'Verificar Agora','Last Sync':'Última Sincronização','Roster Copy':'Copiar Escala','Copy Week':'Copiar Semana','Copy Month':'Copiar Mês','Source Date':'Data de Origem','Target Date':'Data de Destino','Staffing Rules':'Regras de Dotação','Minimum Staffing Rules':'Regras de Dotação Mínima','Audit Log':'Registo de Auditoria','Audit Events':'Eventos de Auditoria','Employee QR Directory':'Directório QR dos Trabalhadores','Holiday Calendar':'Calendário de Feriados','National Holiday':'Feriado Nacional','Company Holiday':'Feriado da Empresa','Bridge Day':'Ponte','Paid':'Remunerado','Unpaid':'Não Remunerado','Late Escalation':'Escalonamento de Atrasos','Late Arrival Escalation':'Escalonamento de Chegadas Tardias','Onboarding Checklist':'Lista de Integração','Offboarding':'Desvinculação','Store Transfer':'Transferência de Loja','Employee Timeline':'Linha do Tempo do Trabalhador','Performance & Training':'Desempenho e Formação','Scheduled Reports':'Relatórios Programados','Offline Queue':'Fila Offline','Offline Biometric Queue':'Fila Biométrica Offline','Backup & Restore':'Cópia de Segurança e Restauro','Create Full Backup':'Criar Cópia de Segurança Completa','Backup History':'Histórico de Cópias de Segurança','System Health':'Estado do Sistema','Application Status':'Estado da Aplicação','Disk Used':'Disco Utilizado','Disk Free':'Disco Livre','Privacy Controls':'Controlos de Privacidade','Data Protection':'Protecção de Dados','Privacy':'Privacidade','Open module':'Abrir módulo','Back':'Voltar','Back to Dashboard':'Voltar ao Painel','Type':'Tipo','Name':'Nome','Title':'Título','Available':'Disponível','Used':'Usado','Annual':'Anual','Total Employees':'Total de Trabalhadores','No notifications':'Sem notificações','Manual Sync':'Sincronização Manual','Training':'Formação','Training Name':'Nome da Formação','Training Date':'Data da Formação','Expiry':'Validade','Expiry Date':'Data de Validade','Remarks':'Observações','Manager Remarks':'Observações do Gestor','Score':'Pontuação','Privacy Notice':'Aviso de Privacidade','Data Subject Rights':'Direitos do Titular dos Dados','Purpose Limitation':'Limitação da Finalidade','Retention':'Conservação','Security':'Segurança','Confidential':'Confidencial','Confidential Document':'Documento Confidencial','Submit':'Enviar','Submit Request':'Enviar Pedido','Save Rule':'Guardar Regra','Save Holiday':'Guardar Feriado','Save Schedule':'Guardar Programação','Apply Filters':'Aplicar Filtros','Clear':'Limpar','Select':'Seleccionar','Choose':'Escolher','Required':'Obrigatório','Optional':'Opcional','Success':'Sucesso','Error':'Erro','Warning':'Aviso','Attention':'Atenção','Critical':'Crítico','Normal':'Normal','System':'Sistema','Application':'Aplicação','User':'Utilizador','Created At':'Criado em','Updated At':'Actualizado em','Effective Date':'Data de Efeito','From Store':'Loja de Origem','To Store':'Loja de Destino','Exit Reason':'Motivo de Saída','Final Settlement':'Acerto Final','Documents Returned':'Documentos Devolvidos','Biometric Registration':'Registo Biométrico','Employee Photo':'Fotografia do Trabalhador','Contract Document':'Documento do Contrato','Store Assignment':'Atribuição de Loja','Default Shift':'Turno Padrão','ID Card Generated':'Cartão de Identificação Gerado','Apply Roster':'Aplicar Escala','All Weeks (Monday-Sunday)':'Todas as Semanas (Segunda a Domingo)','Monday':'Segunda-feira','Tuesday':'Terça-feira','Wednesday':'Quarta-feira','Thursday':'Quinta-feira','Friday':'Sexta-feira','Saturday':'Sábado','Sunday':'Domingo'};
function hrmsTranslate(root=document.body){
 if(window.HRMS_LANGUAGE!=='pt'||!root)return;
 const keys=Object.keys(window.HRMS_PT).sort((a,b)=>b.length-a.length);
 const cv=(value)=>{let t=value||'';keys.forEach(k=>{t=t.split(k).join(window.HRMS_PT[k])});return t};
 const walker=document.createTreeWalker(root,NodeFilter.SHOW_TEXT);const nodes=[];while(walker.nextNode())nodes.push(walker.currentNode);
 nodes.forEach(n=>{if(!n.parentElement||['SCRIPT','STYLE','TEXTAREA'].includes(n.parentElement.tagName))return;n.nodeValue=cv(n.nodeValue)});
 root.querySelectorAll('[placeholder],[title],[aria-label],input[type="button"],input[type="submit"]').forEach(el=>{
   ['placeholder','title','aria-label','value'].forEach(a=>{const v=el.getAttribute(a);if(v)el.setAttribute(a,cv(v))});
 });
 document.documentElement.lang='pt';document.title=cv(document.title);
}
document.addEventListener('DOMContentLoaded',()=>hrmsTranslate());
new MutationObserver(ms=>{if(window.HRMS_LANGUAGE==='pt')ms.forEach(m=>m.addedNodes.forEach(n=>{if(n.nodeType===1)hrmsTranslate(n)}))}).observe(document.documentElement,{childList:true,subtree:true});
</script>
<script id="universal-table-sort">
(function(){
  function value(cell,type){const text=(cell?.innerText||'').trim();if(type==='number'){const n=parseFloat(text.replace(/[^0-9.-]/g,''));return Number.isNaN(n)?-Infinity:n;}return text.toLocaleLowerCase();}
  function enhance(root=document){root.querySelectorAll('table').forEach((table,ti)=>{if(table.dataset.sortReady)return;table.dataset.sortReady='1';const heads=table.querySelectorAll('thead th');heads.forEach((th,ci)=>{if(th.dataset.noSort==='1')return;th.style.cursor='pointer';th.style.userSelect='none';if(!/[↕↑↓]/.test(th.textContent))th.insertAdjacentText('beforeend',' ↕');th.title='Click to sort';th.addEventListener('click',()=>{const tbody=table.tBodies[0];if(!tbody)return;const rows=Array.from(tbody.rows).filter(r=>r.cells.length>ci&&!r.querySelector('[colspan]'));const asc=th.dataset.direction!=='asc';heads.forEach(h=>{h.dataset.direction='';h.textContent=h.textContent.replace(/ [↑↓]$/,' ↕');});th.dataset.direction=asc?'asc':'desc';th.textContent=th.textContent.replace(/ ↕$/,'')+(asc?' ↑':' ↓');const type=th.dataset.sortType||((rows.every(r=>/^[-+]?\d[\d.,]*$/.test((r.cells[ci]?.innerText||'').trim())))?'number':'text');rows.sort((a,b)=>{const av=value(a.cells[ci],type),bv=value(b.cells[ci],type);return (av>bv?1:av<bv?-1:0)*(asc?1:-1);});rows.forEach(r=>tbody.appendChild(r));});});});}
  window.applyKviFilters=function(){const q=(document.getElementById('kvi-search')?.value||'').trim().toLowerCase();const type=document.getElementById('kvi-type-filter')?.value||'ALL';let visible=0;document.querySelectorAll('.kvi-row').forEach(r=>{let rt=(r.dataset.type||'OTHER').trim().toUpperCase();if(rt.includes('KVI')&&rt.includes('POWER'))rt='KVI+POWER';else if(rt.includes('KVI'))rt='KVI';else if(rt.includes('POWER'))rt='POWER SKU';const typeOk=type==='ALL'||(type==='OTHER'&&!['KVI','KVI+POWER','POWER SKU'].includes(rt))||rt===type;const searchOk=!q||(r.dataset.search||'').includes(q);const show=typeOk&&searchOk;r.style.display=show?'':'none';if(show)visible++;});const box=document.getElementById('kvi-filter-count');if(box)box.textContent=visible+' article(s) shown';};
  window.filterKviType=function(type){const sel=document.getElementById('kvi-type-filter');if(sel)sel.value=type;window.applyKviFilters();document.getElementById('kvi-article-table')?.scrollIntoView({behavior:'smooth',block:'start'});};
  document.addEventListener('DOMContentLoaded',()=>{enhance(document);if(window.applyKviFilters)window.applyKviFilters();});window.enhanceSortableTables=enhance;
})();
</script>
<style id="hrms-toast-style">
#hrms-toast-host{position:fixed;top:18px;right:18px;z-index:99999;width:min(420px,calc(100vw - 36px));display:flex;flex-direction:column;gap:10px;pointer-events:none}
.hrms-toast{pointer-events:auto;display:flex;align-items:flex-start;gap:12px;padding:14px 16px;border-radius:14px;color:#fff;font:700 12px/1.45 Inter,system-ui,sans-serif;box-shadow:0 18px 45px rgba(15,23,42,.24);transform:translateX(0);opacity:1;transition:opacity .3s ease,transform .3s ease;overflow-wrap:anywhere}
.hrms-toast-success{background:#059669}.hrms-toast-danger,.hrms-toast-error{background:#e11d48}.hrms-toast-warning{background:#d97706}.hrms-toast-info{background:#2563eb}
.hrms-toast-hide{opacity:0;transform:translateX(30px)}.hrms-toast-message{flex:1}.hrms-toast-close{border:0;background:transparent;color:#fff;font-size:21px;line-height:1;cursor:pointer;min-height:0;padding:0 0 0 6px}
@media(max-width:640px){ #hrms-toast-host{top:10px;right:10px;width:calc(100vw - 20px)}}
</style>
<script id="hrms-toast-script">
(function(){
  function host(){let h=document.getElementById('hrms-toast-host');if(!h){h=document.createElement('div');h.id='hrms-toast-host';h.setAttribute('aria-live','polite');h.setAttribute('aria-atomic','false');document.body.appendChild(h)}return h}
  window.showToast=function(message,type='success'){
    if(message===undefined||message===null||String(message).trim()==='')return;
    const normalized=(type==='error'?'danger':type)||'info';
    const t=document.createElement('div');t.className='hrms-toast hrms-toast-'+normalized;t.setAttribute('role',normalized==='danger'?'alert':'status');
    const m=document.createElement('div');m.className='hrms-toast-message';m.textContent=String(message);
    const x=document.createElement('button');x.type='button';x.className='hrms-toast-close';x.innerHTML='&times;';x.setAttribute('aria-label','Close notification');
    const remove=()=>{if(t.dataset.closing)return;t.dataset.closing='1';t.classList.add('hrms-toast-hide');setTimeout(()=>t.remove(),300)};
    x.addEventListener('click',remove);t.append(m,x);host().appendChild(t);setTimeout(remove,5000);
  };
  document.addEventListener('DOMContentLoaded',()=>document.querySelectorAll('.hrms-flash').forEach(el=>{showToast(el.dataset.message,el.dataset.category);el.remove()}));
})();
</script>
</body>
</html>
"""

ID_CARD_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Employee ID Card - {{ employee.name }}</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
    <style>
        body { display: flex; justify-content: center; align-items: center; min-height: 100vh; background-color: #f1f5f9; margin: 0; font-family: Arial, sans-serif; flex-direction: column; }
        .id-card {
            width: 340px;
            height: 485px;
            border: 1px solid #e0e0e0;
            border-radius: 12px;
            padding: 14px 22px 0 22px;
            text-align: center;
            background-color: #fff;
            box-shadow: 0 10px 25px rgba(0,0,0,0.1);
            position: relative;
            overflow: hidden;
            box-sizing: border-box;
        }
        .logo-wrap { width: 100%; display: flex; justify-content: center; align-items: center; margin: 0 auto 8px; min-height: 42px; }
        .logo { display: block; width: 175px; max-width: 82%; height: 46px; object-fit: contain; object-position: center; border-radius: 4px; }
        .photo-placeholder {
            width: 122px; height: 138px; background-color: #f9f9f9; margin: 0 auto 8px auto;
            display: flex; align-items: center; justify-content: center; overflow: hidden; border-radius: 8px; border: 1px solid #eee;
        }
        .photo-placeholder img { width: 100%; height: 100%; object-fit: cover; }
        .emp-name { font-size: 19px; font-weight: bold; margin: 0 auto 12px; line-height: 1.18; text-transform: uppercase; color: #333; max-width: 285px; }
        .details { text-align: left; font-size: 12px; line-height: 1.42; color: #000; padding: 0 5px 64px; width: 100%; box-sizing: border-box; }
        .details div { display: grid; grid-template-columns: 126px minmax(0, 1fr); align-items: center; column-gap: 4px; min-height: 20px; white-space: nowrap; }
        .details div span:first-child { width: auto; min-width: 0; font-weight: normal; white-space: nowrap; }
        .details div span:last-child { min-width: 0; font-weight: bold; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; font-size: 11px; }
        .details .designation-value { font-size: 10.5px; letter-spacing: -0.15px; }
        .footer-container { position: absolute; bottom: 0; left: 0; width: 100%; }
        .color-bar { height: 8px; background: linear-gradient(to right, #cddc39, #ffc107, #ff9800, #e91e63, #9c27b0, #00bcd4); }
        .footer-address { background-color: #000; color: #fff; font-size: 10px; padding: 8px 10px; line-height: 1.25; text-align: center; }
        .download-btn {
            margin-top: 14px; padding: 12px 24px; background-color: #059669; color: white; border: none;
            border-radius: 8px; font-weight: bold; cursor: pointer; box-shadow: 0 4px 6px rgba(0,0,0,0.1); font-size: 14px;
        }
        .download-btn:hover { background-color: #047857; }
    </style>
<style id="id-card-mobile-fix">@media(max-width:640px){body{padding:10px;overflow-x:hidden}.id-card{width:min(340px,94vw)!important}.download-btn{width:min(340px,94vw);margin-left:auto;margin-right:auto}}</style></head>
<body>
    <div class="id-card" id="id-card-element">
        <div class="logo-wrap"><img src="{{ url_for('static', filename='fresmart.png') }}" alt="Fresmart Logo" class="logo"></div>
        
        <div class="photo-placeholder">
            {% if employee.profile_photo %}
                <img src="/uploads/{{ employee.profile_photo }}" alt="Employee Photo">
            {% endif %}
        </div>

        <div class="emp-name">{{ employee.name }}</div>

        <div class="details">
            
            <div><span>HRMS Code</span> <span>: {{ employee.emp_code }}</span></div>
            <div><span>Designation</span> <span class="designation-value">: {{ employee.designation or employee.dept or 'Employee' }}</span></div>
            <div><span>Identificação</span> <span>: {{ employee.identificacao or '-' }}</span></div>
            
            <div><span>Data de Contrato</span> <span>: {{ employee.data_de_contrato or '-' }}</span></div>
        </div>

        <div class="footer-container">
            <div class="color-bar"></div>
            <div class="footer-address">
                Kwame Nkrumah,<br>
                Edifício Torre Imporáfrica B, 9º Andar,<br>
                Município: Ingombota, Luanda<br>
                www.newacogrupo.com
            </div>
        </div>
    </div>

    <button class="download-btn" onclick="downloadCard()">📥 Baixar / Download ID Card</button>

    <script>
        function downloadCard() {
            const card = document.getElementById('id-card-element');
            html2canvas(card, {scale: 3, useCORS: true}).then(canvas => {
                let link = document.createElement('a');
                link.download = '{{ employee.emp_code }}_ID_Card.png';
                link.href = canvas.toDataURL('image/png');
                link.click();
            });
        }
    </script>

<style id="hrms-toast-style">
#hrms-toast-host{position:fixed;top:18px;right:18px;z-index:99999;width:min(420px,calc(100vw - 36px));display:flex;flex-direction:column;gap:10px;pointer-events:none}
.hrms-toast{pointer-events:auto;display:flex;align-items:flex-start;gap:12px;padding:14px 16px;border-radius:14px;color:#fff;font:700 12px/1.45 Inter,system-ui,sans-serif;box-shadow:0 18px 45px rgba(15,23,42,.24);transform:translateX(0);opacity:1;transition:opacity .3s ease,transform .3s ease;overflow-wrap:anywhere}
.hrms-toast-success{background:#059669}.hrms-toast-danger,.hrms-toast-error{background:#e11d48}.hrms-toast-warning{background:#d97706}.hrms-toast-info{background:#2563eb}
.hrms-toast-hide{opacity:0;transform:translateX(30px)}.hrms-toast-message{flex:1}.hrms-toast-close{border:0;background:transparent;color:#fff;font-size:21px;line-height:1;cursor:pointer;min-height:0;padding:0 0 0 6px}
@media(max-width:640px){ #hrms-toast-host{top:10px;right:10px;width:calc(100vw - 20px)}}
</style>
<script id="hrms-toast-script">
(function(){
  function host(){let h=document.getElementById('hrms-toast-host');if(!h){h=document.createElement('div');h.id='hrms-toast-host';h.setAttribute('aria-live','polite');h.setAttribute('aria-atomic','false');document.body.appendChild(h)}return h}
  window.showToast=function(message,type='success'){
    if(message===undefined||message===null||String(message).trim()==='')return;
    const normalized=(type==='error'?'danger':type)||'info';
    const t=document.createElement('div');t.className='hrms-toast hrms-toast-'+normalized;t.setAttribute('role',normalized==='danger'?'alert':'status');
    const m=document.createElement('div');m.className='hrms-toast-message';m.textContent=String(message);
    const x=document.createElement('button');x.type='button';x.className='hrms-toast-close';x.innerHTML='&times;';x.setAttribute('aria-label','Close notification');
    const remove=()=>{if(t.dataset.closing)return;t.dataset.closing='1';t.classList.add('hrms-toast-hide');setTimeout(()=>t.remove(),300)};
    x.addEventListener('click',remove);t.append(m,x);host().appendChild(t);setTimeout(remove,5000);
  };
  document.addEventListener('DOMContentLoaded',()=>document.querySelectorAll('.hrms-flash').forEach(el=>{showToast(el.dataset.message,el.dataset.category);el.remove()}));
})();
</script>
</body>
</html>
"""

@app.errorhandler(413)
def upload_too_large(error):
    flash(f"Selected files exceed the {int(app.config['MAX_CONTENT_LENGTH']/1024/1024)} MB total upload limit.",'danger')
    return redirect(request.referrer or url_for('index'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    failed_attempts = session.get('failed_attempts', 0)
    
    if request.method == 'POST':
        uid = request.form.get('user_id').strip().upper()
        pwd = request.form.get('password').strip()
        
        dev_pass = os.getenv('DEV_PWD', 'Shama@8577')
        db = load_users_db()
        
        if uid == 'NCSA0608' and pwd == dev_pass:
            session.update({'logged_in': True, 'role': 'developer', 'user_id': uid, 'user_name': 'Sonu Kumar (Dev)', 'store': 'DEV', 'stores': list(MACHINES.keys()), 'designation':'Developer', 'job_role':'DEVELOPER', 'permissions':PERMISSIONS[:]})
            return redirect(url_for('index'))
            
        emp_key = f"NWC{uid}" if not uid.startswith('NWC') and uid not in ['LM11', 'LF07'] else uid
        
        if emp_key in db:
            user_data = db[emp_key]
            if user_data.get('status') == 'blocked':
                return render_template_string(LOGIN_TEMPLATE, error="Account is Suspended/Blocked. Contact Developer.")
            
            if pwd == user_data.get('password'):
                primary_store=user_data.get('store','LM11')
                session.update({'logged_in': True, 'role': user_data['role'], 'user_id': emp_key, 'user_name': user_data['name'], 'store': primary_store, 'stores': [primary_store], 'designation':user_data.get('designation',''), 'job_role':user_data.get('job_role',user_data.get('role','employee')), 'permissions':user_data.get('permissions',[])})
                session['failed_attempts'] = 0
                return redirect(url_for('index'))
                
        failed_attempts += 1
        session['failed_attempts'] = failed_attempts
        if failed_attempts >= 5:
            session['failed_attempts'] = 0
            flash("5 baar galat password dala gaya. Kripya apna password reset karein.", "danger")
            return redirect(url_for('reset_password'))
        return render_template_string(LOGIN_TEMPLATE, error=f"Galat User ID ya Password! (Attempt {failed_attempts}/5)")
        
    return render_template_string(LOGIN_TEMPLATE, error=None)

@app.route('/reset_password', methods=['GET', 'POST'])
def reset_password():
    if request.method == 'POST':
        uid = request.form.get('user_id').strip().upper()
        new_pwd = request.form.get('new_password').strip()
        
        if uid == 'NCSA0608':
            flash("Developer ka password change nahi kiya ja sakta!", "danger")
            return redirect(url_for('reset_password'))
            
        emp_key = f"NWC{uid}" if not uid.startswith('NWC') and uid not in ['LM11', 'LF07'] else uid
        
        db = load_users_db()
        if emp_key not in db:
            flash("User ID system me nahi mila!", "danger")
            return redirect(url_for('reset_password'))
            
        name = db[emp_key]['name']
        role_req = db[emp_key].get('role', 'employee')
        store_req = db[emp_key].get('store', 'LM11')
        
        resets = load_json_file(PASSWORD_RESETS_FILE)
        if not isinstance(resets, list): resets = []
        
        resets.append({
            'id': len(resets) + 1,
            'user_id': emp_key,
            'name': name,
            'new_password': new_pwd,
            'status': 'Pending',
            'role': role_req,
            'store': store_req,
            'request_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'approve_date': '-',
            'approved_by': '-'
        })
        save_json_file(PASSWORD_RESETS_FILE, resets)
        
        flash("Your request sent successfully for approval 🚀", "success")
        return redirect(url_for('login'))
        
    return render_template_string(RESET_TEMPLATE)

@app.route('/update_password_req/<int:req_id>/<action>')
def update_password_req(req_id, action):
    if session.get('role') not in ['admin', 'developer']: return redirect(url_for('login'))
    
    resets = load_json_file(PASSWORD_RESETS_FILE)
    for req in resets:
        if req['id'] == req_id:
            # Authorization Check for Admin
            if session.get('role') == 'admin':
                if req.get('role') == 'admin' or req.get('store') != session.get('store'):
                    flash("Aap is user ki request ko approve nahi kar sakte.", "danger")
                    return redirect(url_for('index'))

            if action == 'approve':
                req['status'] = 'Approved'
                req['approve_date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                req['approved_by'] = session.get('user_name')
                
                db = load_users_db()
                if req['user_id'] in db:
                    db[req['user_id']]['password'] = req['new_password']
                    save_json_file(USERS_DB_FILE, db)
                flash(f"{req['name']} ka naya password approve ho gaya hai. Purana ab block ho gaya.", "success")
                
            elif action == 'reject':
                req['status'] = 'Rejected'
                req['approve_date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                req['approved_by'] = session.get('user_name')
                flash(f"{req['name']} ki password request reject ki gayi.", "success")
            break
            
    save_json_file(PASSWORD_RESETS_FILE, resets)
    return redirect(url_for('index'))

@app.route('/dev_force_reset', methods=['POST'])
def dev_force_reset():
    if session.get('role') != 'developer': return redirect(url_for('login'))
        
    target_uid = request.form.get('target_user_id').strip().upper()
    new_pwd = request.form.get('target_new_password').strip()
    
    if target_uid == 'NCSA0608':
        flash("Developer ka password yahan se change nahi kiya ja sakta!", "danger")
        return redirect(url_for('index'))
        
    emp_key = f"NWC{target_uid}" if not target_uid.startswith('NWC') and target_uid not in ['LM11', 'LF07'] else target_uid
    
    db = load_users_db()
    if emp_key not in db:
        flash(f"User ID {target_uid} system me nahi mila!", "danger")
        return redirect(url_for('index'))
        
    db[emp_key]['password'] = new_pwd
    save_json_file(USERS_DB_FILE, db)
    
    flash(f"{emp_key} ka password successfully update ho gaya hai.", "success")
    return redirect(url_for('index'))

@app.route('/manage_user', methods=['POST'])
def manage_user():
    if session.get('role') not in ['admin','developer']: return redirect(url_for('login'))
    db=load_users_db(); action=request.form.get('action','create'); raw=(request.form.get('uid') or '').strip().upper()
    uid=raw if raw in ['LM11','LF07'] or raw.startswith('NWC') else f'NWC{raw}'
    if action=='create':
        old=normalize_user_record(db.get(uid,{})); role=request.form.get('role','employee')
        job=(request.form.get('job_role') or ('ADMIN' if role=='admin' else 'EMPLOYEE')).strip().upper()
        stores=[x.upper() for x in request.form.getlist('stores') if x.upper() in MACHINES and x.upper()!='DEV']
        if request.form.get('all_stores')=='1': stores=[x for x in MACHINES if x!='DEV']
        if session.get('role') == 'admin':
            role='employee'; job='EMPLOYEE'
            allowed_admin_stores=get_user_stores()
            stores=[x for x in stores if x in allowed_admin_stores] or allowed_admin_stores[:1]

        if not stores: stores=old.get('stores') or [request.form.get('store','LM11').upper()]
        perms=[x for x in request.form.getlist('permissions') if x in PERMISSIONS]
        if not perms: perms=ROLE_PRESETS.get(job,ROLE_PRESETS['EMPLOYEE'])[:]
        db[uid]=normalize_user_record({**old,'name':(request.form.get('name') or old.get('name') or uid).strip(),'password':(request.form.get('password') or old.get('password') or '123').strip(),'role':role if role in ['admin','employee'] else 'employee','designation':(request.form.get('designation') or old.get('designation') or 'Employee').strip(),'email':(request.form.get('email') or old.get('email') or '').strip(),'identificacao':(request.form.get('identificacao') or old.get('identificacao') or '').strip(),'data_de_contrato':(request.form.get('data_de_contrato') or old.get('data_de_contrato') or '').strip(),'job_role':job,'stores':stores,'store':stores[0],'permissions':perms,'status':request.form.get('status','active'),'dept':(request.form.get('dept') or old.get('dept') or 'General').strip()})
        flash(f'User {uid} created/updated successfully.','success')
    elif action=='toggle_status' and uid in db:
        db[uid]=normalize_user_record(db[uid]); db[uid]['status']='blocked' if db[uid].get('status')=='active' else 'active'; flash(f'User {uid} status changed to {db[uid]["status"]}.','success')
    elif action=='delete' and uid in db: del db[uid]; flash(f'User {uid} deleted permanently.','success')
    save_json_file(USERS_DB_FILE,db); return redirect(url_for('index'))

@app.route('/developer_user_rights/<uid>',methods=['POST'])
def developer_user_rights(uid):
    if session.get('role')!='developer':return 'Developer access required',403
    uid=uid.upper();db=load_users_db()
    if uid not in db:return 'User not found',404
    info=normalize_user_record(db[uid]);selected=[p for p in request.form.getlist('permissions') if p in PERMISSIONS]
    info['permissions']=selected
    if request.form.get('role') in ['admin','employee']:info['role']=request.form.get('role')
    stores=[x for x in request.form.getlist('stores') if x in MACHINES and x!='DEV']
    if stores:info['stores']=stores;info['store']=stores[0]
    info['rights_updated_at']=datetime.now().strftime('%Y-%m-%d %H:%M:%S');info['rights_updated_by']=session.get('user_name')
    db[uid]=info;save_json_file(USERS_DB_FILE,db)
    flash(f'Portal rights updated for {info.get("name",uid)}.','success');return redirect(url_for('index'))

@app.route('/admin_employee_photo/<uid>',methods=['POST'])
def admin_employee_photo(uid):
    if session.get('role') not in ['admin','developer']:return redirect(url_for('login'))
    uid=uid.upper();db=load_users_db();info=normalize_user_record(db.get(uid,{}))
    if not info or info.get('role')!='employee':return 'Employee not found',404
    if session.get('role')!='developer' and info.get('store') not in get_user_stores():return 'Unauthorized',403
    f=request.files.get('profile_photo')
    if not f or not f.filename:flash('Select an employee photo.','danger');return redirect(url_for('index'))
    ext=Path(f.filename).suffix.lower()
    if ext not in ['.png','.jpg','.jpeg','.webp']:flash('Only JPG, JPEG, PNG or WEBP photos are allowed.','danger');return redirect(url_for('index'))
    old=info.get('profile_photo','');filename=secure_filename(f'profile_{uid}_{uuid.uuid4().hex}{ext}');path=os.path.join(app.config['UPLOAD_FOLDER'],filename);f.save(path)
    if old and os.path.isfile(os.path.join(app.config['UPLOAD_FOLDER'],old)):
        try:os.remove(os.path.join(app.config['UPLOAD_FOLDER'],old))
        except OSError:pass
    info['profile_photo']=filename;info['photo_updated_at']=datetime.now().strftime('%Y-%m-%d %H:%M:%S');db[uid]=info;save_json_file(USERS_DB_FILE,db)
    flash(f'Photo updated for {info.get("name",uid)}.','success');return redirect(url_for('index'))

@app.route('/upload_profile_photo', methods=['POST'])
def upload_profile_photo():
    if not session.get('logged_in'): return redirect(url_for('login'))
    uid=session.get('user_id'); db=load_users_db(); f=request.files.get('profile_photo')
    if not f or not f.filename: flash('Photo select karein.','danger'); return redirect(url_for('index'))
    ext=os.path.splitext(f.filename)[1].lower()
    if ext not in ['.png','.jpg','.jpeg','.webp']: flash('Sirf JPG, JPEG, PNG ya WEBP photo allowed hai.','danger'); return redirect(url_for('index'))
    info=normalize_user_record(db.get(uid,{})); old=info.get('profile_photo',''); filename=secure_filename(f'profile_{uid}_{uuid.uuid4().hex}{ext}'); f.save(os.path.join(app.config['UPLOAD_FOLDER'],filename))
    if old and os.path.isfile(os.path.join(app.config['UPLOAD_FOLDER'],old)):
        try: os.remove(os.path.join(app.config['UPLOAD_FOLDER'],old))
        except OSError: pass
    info['profile_photo']=filename; db[uid]=info; save_json_file(USERS_DB_FILE,db); flash('Profile photo update ho gaya.','success'); return redirect(url_for('index'))

@app.route('/remove_profile_photo', methods=['POST'])
def remove_profile_photo():
    if not session.get('logged_in'): return redirect(url_for('login'))
    uid=session.get('user_id'); db=load_users_db(); info=normalize_user_record(db.get(uid,{})); old=info.get('profile_photo','')
    if old and os.path.isfile(os.path.join(app.config['UPLOAD_FOLDER'],old)):
        try: os.remove(os.path.join(app.config['UPLOAD_FOLDER'],old))
        except OSError: pass
    info['profile_photo']=''; db[uid]=info; save_json_file(USERS_DB_FILE,db); flash('Profile photo remove ho gaya.','success'); return redirect(url_for('index'))

@app.route('/manage_machine', methods=['POST'])
def manage_machine():
    if session.get('role')!='developer': return redirect(url_for('login'))
    code=(request.form.get('code') or '').strip().upper(); action=request.form.get('action','save'); machines=load_machines_db()
    if action=='delete':
        if code in ['LM11','LF07','DEV']: flash('Default machine delete nahi ki ja sakti.','danger')
        else: machines.pop(code,None); save_json_file(MACHINES_DB_FILE,machines); load_machines_db(); flash(f'{code} machine remove ho gayi.','success')
        return redirect(url_for('index'))
    if not re.match(r'^[A-Z0-9_-]{2,30}$',code): flash('Store/Machine code invalid hai.','danger'); return redirect(url_for('index'))
    try: port=int(request.form.get('port','4370'))
    except ValueError: flash('Port invalid hai.','danger'); return redirect(url_for('index'))
    machines[code]={'ip':(request.form.get('ip') or '').strip(),'port':port,'name':(request.form.get('name') or code).strip(),'admin':(request.form.get('admin') or '').strip()}; save_json_file(MACHINES_DB_FILE,machines); load_machines_db(); flash(f'{code} biometric portal save ho gaya.','success'); return redirect(url_for('index'))

@app.route('/set_language/<lang>')
def set_language(lang):
    session['ui_language'] = 'pt' if lang == 'pt' else 'en'
    return redirect(request.referrer or url_for('index'))

@app.route('/logout')
def logout():
    role = session.get('role')
    user_name = session.get('user_name', '')
    msg_text = "Thank you admin" if role == 'admin' else f"Thank you {user_name}"
    session.clear()
    flash(msg_text, 'success')
    return redirect(url_for('login'))

@app.route('/')
def index():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
        
    role=session.get('role'); logged_user_id=session.get('user_id'); db_all=load_users_db()
    ui_language=session.get('ui_language','pt' if role=='employee' else 'en')
    if ui_language not in ['en','pt']: ui_language='en'
    session['ui_language']=ui_language
    current_user=normalize_user_record(db_all.get(logged_user_id,{'name':session.get('user_name','')})) if role!='developer' else {'name':'Sonu Kumar (Dev)','profile_photo':'','designation':'Developer','job_role':'DEVELOPER','stores':list(MACHINES.keys()),'permissions':PERMISSIONS}
    accessible_stores=get_user_stores(logged_user_id)
    requested_store=(request.args.get('store') or session.get('store') or (accessible_stores[0] if accessible_stores else 'LM11')).upper()
    store=requested_store if requested_store in accessible_stores else (accessible_stores[0] if accessible_stores else 'LM11'); session['store']=store; session['stores']=accessible_stores
    portal_name=MACHINES.get(store,MACHINES.get('LM11'))['name']; logged_user_name=session.get('user_name'); selected_emp=logged_user_id if role=='employee' else request.args.get('employee','ALL')
        
    today_str = datetime.now().strftime('%Y-%m-%d')
    start_date = request.args.get('start_date', today_str)
    end_date = request.args.get('end_date', today_str)
    
    logs, all_users, g_hrs, g_l_hrs, g_var, raw_punches, stats = fetch_attendance_data(start_date, end_date, selected_emp, store)
    
    pending_leaves_count = sum(1 for req in LEAVE_REQUESTS if req['status'] == 'Pending')
    current_user_leave_requests = [req for req in LEAVE_REQUESTS if req['user_id'] == logged_user_id] if role == 'employee' else LEAVE_REQUESTS
    roster_first=datetime.now().replace(day=1)
    roster_next=(roster_first.replace(day=28)+timedelta(days=4)).replace(day=1)
    roster_days=(roster_next-roster_first).days
    employee_saved_roster=load_roster().get(logged_user_id,{}) if role=='employee' else {}
    employee_default_shift=current_user.get('shift','morning') if role=='employee' else 'morning'
    employee_default_off=str(current_user.get('off','')).upper() if role=='employee' else ''
    pt_weekdays=['Segunda-feira','Terça-feira','Quarta-feira','Quinta-feira','Sexta-feira','Sábado','Domingo']
    employee_month_roster=[]
    employee_roster_counts={'shift_a':0,'shift_b':0,'weekly_off':0,'default':0}
    for roster_i in range(roster_days):
        roster_date=roster_first+timedelta(days=roster_i)
        roster_date_str=roster_date.strftime('%Y-%m-%d')
        roster_value=employee_saved_roster.get(roster_date_str,'')
        if not roster_value:
            if roster_date.strftime('%A').upper()==employee_default_off:
                roster_value='Weekly Off'
            elif employee_default_shift=='second':
                roster_value='Shift B'
            else:
                roster_value='Shift A'
        if roster_value=='Shift A': employee_roster_counts['shift_a']+=1
        elif roster_value=='Shift B': employee_roster_counts['shift_b']+=1
        elif roster_value=='Weekly Off': employee_roster_counts['weekly_off']+=1
        else: employee_roster_counts['default']+=1
        employee_month_roster.append({'date':roster_date_str,'display_date':roster_date.strftime('%d/%m/%Y'),'weekday_en':roster_date.strftime('%A'),'weekday_pt':pt_weekdays[roster_date.weekday()],'value':roster_value,'value_en':roster_value or 'Default','value_pt':{'Shift A':'Turno A','Shift B':'Turno B','Weekly Off':'Folga Semanal'}.get(roster_value,'Padrão'),'is_today':roster_date.date()==datetime.now().date()})
    pt_months={1:'Janeiro',2:'Fevereiro',3:'Março',4:'Abril',5:'Maio',6:'Junho',7:'Julho',8:'Agosto',9:'Setembro',10:'Outubro',11:'Novembro',12:'Dezembro'}
    employee_roster_month_label=(f"{pt_months[roster_first.month]} {roster_first.year}" if ui_language=='pt' else roster_first.strftime('%B %Y'))
        
    shift_reqs = load_shift_requests()
    pending_shifts_count = sum(1 for req in shift_reqs if req.get('status') == 'Pending')
    my_shift_reqs = [r for r in shift_reqs if r.get('user_id') == logged_user_id] if role == 'employee' else shift_reqs
    
    # Password Reset Visibility Logic
    if role == 'developer':
        reset_requests = load_json_file(PASSWORD_RESETS_FILE)
        pending_resets_count = sum(1 for req in reset_requests if req.get('status') == 'Pending')
    elif role == 'admin':
        all_resets = load_json_file(PASSWORD_RESETS_FILE)
        reset_requests = [r for r in all_resets if r.get('store') == store and r.get('role') != 'admin']
        pending_resets_count = sum(1 for req in reset_requests if req.get('status') == 'Pending')
    else:
        reset_requests = []
        pending_resets_count = 0
    
    overtime_all=load_overtime_approvals()
    overtime_visible=[x for x in overtime_all if role=='developer' or (role=='admin' and x.get('store') in accessible_stores) or x.get('user_id')==logged_user_id]
    pending_overtime_count=sum(1 for x in overtime_visible if x.get('status')=='Pending')
    # Flash new workflow notifications once in the intended portal.
    notif_rows=suite_load(NOTIFICATIONS_FILE,[]);notif_changed=False
    for n in notif_rows:
        intended=((role=='employee' and n.get('user_id')==logged_user_id and 'approved' in n.get('title','').lower()) or (role=='admin' and n.get('user_id') in accessible_stores) or role=='developer')
        if intended and not n.get('flash_seen',False) and ('Overtime' in n.get('title','') or 'overtime' in n.get('title','')):
            flash(f"{n.get('title')}: {n.get('message')}",'success' if 'approved' in n.get('title','').lower() else 'danger' if 'rejected' in n.get('title','').lower() else 'success')
            n['flash_seen']=True;notif_changed=True
    if notif_changed:suite_save(NOTIFICATIONS_FILE,notif_rows)
    all_salary_slips = load_salary_slips()
    
    # FIX: List comprehension applied here to show only the logged-in employee's slips
    my_salary_slips = [s for s in all_salary_slips if s['user_id'] == logged_user_id] if role == 'employee' else all_salary_slips
    
    users_db=load_users_db() if role=='developer' else {}
    employee_cards=[]
    for emp_id, emp_info_raw in db_all.items():
        emp_info=normalize_user_record(emp_info_raw)
        if emp_info.get('role')!='employee': continue
        if role!='developer' and not any(st in accessible_stores for st in emp_info.get('stores',[])): continue
        employee_cards.append({'user_id':emp_id,'name':emp_info.get('name',emp_id),'designation':emp_info.get('designation') or emp_info.get('dept','Employee'),'dept':emp_info.get('dept','General'),'email':emp_info.get('email',''),'profile_photo':emp_info.get('profile_photo',''),'photo_updated_at':emp_info.get('photo_updated_at',''),'identificacao':emp_info.get('identificacao',''),'data_de_contrato':emp_info.get('data_de_contrato',''),'stores':emp_info.get('stores',[])})
    employee_cards.sort(key=lambda x:x['name'])
    employee_designations=sorted({x['designation'] for x in employee_cards if x['designation']})

    machine_status={code:check_device_connectivity(code) for code in accessible_stores if code in MACHINES}
    current_machine=MACHINES.get(store,{})
    
    return render_template_string(
        HTML_TEMPLATE,
        logs=logs, all_users=all_users, start_date=start_date, end_date=end_date, selected_emp=selected_emp,
        grand_total_hours=g_hrs, grand_total_lunch_hours=g_l_hrs, grand_total_variance=g_var, stats=stats,
        raw_punches=raw_punches, role=role, logged_user_name=logged_user_name, portal_name=portal_name,
        leave_requests=current_user_leave_requests, pending_leaves_count=pending_leaves_count,
        shift_requests=my_shift_reqs, pending_shifts_count=pending_shifts_count,
        reset_requests=reset_requests, pending_resets_count=pending_resets_count,
        salary_slips=my_salary_slips, users_db=users_db, store=store, accessible_stores=accessible_stores, current_user=current_user, machine_status=machine_status, current_machine=current_machine, machines=MACHINES, permission_list=PERMISSIONS, permission_labels=PERMISSION_LABELS, role_presets=ROLE_PRESETS, employee_cards=employee_cards, employee_designations=employee_designations, employee_month_roster=employee_month_roster, employee_roster_counts=employee_roster_counts, employee_roster_month_label=employee_roster_month_label, leave_code_names=LEAVE_CODE_NAMES, ui_language=ui_language, overtime_rows=overtime_visible, pending_overtime_count=pending_overtime_count
    )

# --- BULK SALARY SLIP APIs ---
@app.route('/upload_bulk_salary', methods=['POST'])
def upload_bulk_salary():
    if not session_has_permission('payroll'):
        return redirect(url_for('index'))

    salary_month = (request.form.get('salary_month') or '').strip()
    file = request.files.get('salary_pdf')
    if not salary_month or not file or file.filename == '':
        flash('Sabhi fields bharna zaroori hai!', 'danger')
        return redirect(url_for('index'))
    if not file.filename.lower().endswith('.pdf'):
        flash('Sirf PDF files allowed hain!', 'danger')
        return redirect(url_for('index'))

    try:
        reader = PdfReader(file)
        total_pages = len(reader.pages)
        if total_pages == 0:
            flash('Uploaded PDF mein koi page nahi hai!', 'danger')
            return redirect(url_for('index'))

        page_groups = _build_salary_pdf_groups(reader)
        slips = load_salary_slips()

        uploaded_employee_codes = set(page_groups.keys())
        kept_slips = []
        for old in slips:
            if old.get('month') == salary_month and old.get('user_id') in uploaded_employee_codes:
                old_path = os.path.join(app.config['UPLOAD_FOLDER'], old.get('filename', ''))
                if os.path.isfile(old_path):
                    try:
                        os.remove(old_path)
                    except OSError:
                        pass
            else:
                kept_slips.append(old)
        slips = kept_slips

        pages_processed = 0
        employees_processed = 0
        for emp_code, page_indexes in page_groups.items():
            final_emp_code = emp_code if emp_code.startswith('NWC') else f'NWC{emp_code}'
            emp_info = get_emp_info(emp_code)
            writer = PdfWriter()
            for page_index in sorted(page_indexes):
                writer.add_page(reader.pages[page_index])
                pages_processed += 1

            file_id = str(uuid.uuid4())
            secure_name = f'salary_{final_emp_code}_{salary_month}_{file_id}.pdf'
            local_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_name)
            with open(local_path, 'wb') as output_pdf:
                writer.write(output_pdf)

            slips.append({
                'file_id': file_id,
                'user_id': final_emp_code,
                'emp_name': emp_info['name'],
                'month': salary_month,
                'filename': secure_name,
                'page_count': len(page_indexes),
                'upload_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            })
            employees_processed += 1

        save_salary_slips(slips)
        missing_employees = [
            get_emp_info(code)['name'] for code in MASTER_EMPLOYEES
            if code not in page_groups
        ]

        if missing_employees:
            preview = ', '.join(missing_employees[:5])
            suffix = ' ...' if len(missing_employees) > 5 else ''
            flash(
                f'Success! {pages_processed}/{total_pages} pages save hue aur {employees_processed} employees ki slips bani. '
                f'{len(missing_employees)} employees ke liye PDF mein identifiable page nahi mila: {preview}{suffix}',
                'danger'
            )
        else:
            flash(
                f'Success! Merged PDF ke saare {total_pages} pages save ho gaye aur {employees_processed} employees ke salary slips ban gaye. '
                f'Multi-page slips ke saare pages ek hi employee slip mein rakhe gaye.',
                'success'
            )
    except Exception as e:
        flash(f'PDF Split Error: {str(e)}', 'danger')
    return redirect(url_for('index'))

@app.route('/upload_individual_salary', methods=['POST'])
def upload_individual_salary():
    if not session_has_permission('payroll'):
        return redirect(url_for('index'))

    emp_id = request.form.get('emp_id')
    salary_month = request.form.get('salary_month')
    file = request.files.get('individual_pdf')

    if not emp_id or not salary_month or not file or file.filename == '':
        flash('Sabhi fields bharna zaroori hai!', 'danger')
        return redirect(url_for('index'))

    if not file.filename.lower().endswith('.pdf'):
        flash('Sirf PDF files allowed hain!', 'danger')
        return redirect(url_for('index'))

    emp_info = get_emp_info(emp_id)
    final_emp_code = emp_id if emp_id.startswith('NWC') else f'NWC{emp_id}'

    file_id = str(uuid.uuid4())
    secure_name = f'salary_{final_emp_code}_{salary_month}_{file_id}.pdf'
    local_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_name)
    file.save(local_path)

    slips = load_salary_slips()
    # Remove old slip for same month and same user if it already exists
    kept_slips = []
    for old in slips:
        if old.get('month') == salary_month and old.get('user_id') == final_emp_code:
            old_path = os.path.join(app.config['UPLOAD_FOLDER'], old.get('filename', ''))
            if os.path.isfile(old_path):
                try: os.remove(old_path)
                except OSError: pass
        else:
            kept_slips.append(old)
    slips = kept_slips

    slips.append({
        'file_id': file_id,
        'user_id': final_emp_code,
        'emp_name': emp_info['name'],
        'month': salary_month,
        'filename': secure_name,
        'page_count': len(PdfReader(local_path).pages) if os.path.exists(local_path) else 1,
        'upload_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    })

    save_salary_slips(slips)
    flash(f'{emp_info["name"]} ki individual salary slip successfully upload ho gayi!', 'success')
    return redirect(url_for('index'))

@app.route('/salary_file/<file_id>')
def salary_file(file_id):
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    slips = load_salary_slips()
    target_slip = next((s for s in slips if s.get('file_id') == file_id), None)
    if not target_slip:
        return 'File not found', 404

    if session.get('role') == 'employee':
        if target_slip.get('user_id') != session.get('user_id'):
            return 'Unauthorized Access', 403
    elif not session_has_permission('payroll'):
        return 'Unauthorized Access', 403

    file_path = os.path.join(app.config['UPLOAD_FOLDER'], target_slip.get('filename', ''))
    if not os.path.isfile(file_path):
        return 'Salary slip file missing on server', 404
    return send_file(file_path, mimetype='application/pdf')

@app.route('/delete_salary/<file_id>')
def delete_salary(file_id):
    if not session_has_permission('payroll'):
        return redirect(url_for('index'))
    slips = load_salary_slips()
    target_slip = next((s for s in slips if s.get('file_id') == file_id), None)
    if target_slip:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], target_slip.get('filename', ''))
        if os.path.exists(file_path):
            os.remove(file_path)
        slips = [s for s in slips if s.get('file_id') != file_id]
        save_salary_slips(slips)
        flash('Salary slip successfully deleted.', 'success')
    return redirect(url_for('index'))

# --- MONTHLY ROSTER MATRIX APIs ---
@app.route('/api/roster_matrix')
def api_roster_matrix():
    if not session_has_permission('roster'): return jsonify({'error':'Unauthorized'}), 403
    month=request.args.get('month',datetime.now().strftime('%Y-%m'))
    try: first=datetime.strptime(month+'-01','%Y-%m-%d')
    except ValueError: return jsonify({'error':'Invalid month'}),400
    next_month=(first.replace(day=28)+timedelta(days=4)).replace(day=1)
    days=(next_month-first).days
    weekday_pt=['Seg','Ter','Qua','Qui','Sex','Sáb','Dom']; dates=[{'date':(first+timedelta(days=i)).strftime('%Y-%m-%d'),'day':i+1,'weekday':(weekday_pt[(first+timedelta(days=i)).weekday()] if session.get('ui_language')=='pt' else (first+timedelta(days=i)).strftime('%a'))} for i in range(days)]
    all_roster=load_roster(); db=load_users_db(); stores=get_user_stores(); role=session.get('role')
    employees=[]; summary={d['date']:{'shift_a':0,'shift_b':0,'weekly_off':0} for d in dates}
    for uid,raw in db.items():
        info=normalize_user_record(raw)
        if info.get('role')!='employee' or info.get('status')=='blocked': continue
        if role!='developer' and (not stores or stores[0] not in info.get('stores',[])): continue
        r={d['date']:all_roster.get(uid,{}).get(d['date'],'') for d in dates}
        for dt,val in r.items():
            if val=='Shift A': summary[dt]['shift_a']+=1
            elif val=='Shift B': summary[dt]['shift_b']+=1
            elif val=='Weekly Off': summary[dt]['weekly_off']+=1
        employees.append({'user_id':uid,'name':info.get('name',uid),'designation':info.get('designation') or info.get('dept','Employee'),'dept':info.get('dept','General'),'roster':r})
    employees.sort(key=lambda x:x['name'])
    return jsonify({'month':month,'dates':dates,'employees':employees,'summary':summary})

@app.route('/api/save_roster_matrix',methods=['POST'])
def api_save_roster_matrix():
    if not session_has_permission('roster'): return jsonify({'error':'Unauthorized'}),403
    payload=request.get_json(silent=True) or {}; updates=payload.get('updates',[]); roster=load_roster(); db=load_users_db(); stores=get_user_stores(); role=session.get('role'); saved=0
    for item in updates:
        uid=str(item.get('emp_id','')).upper(); dt=str(item.get('date','')); val=str(item.get('value',''))
        info=normalize_user_record(db.get(uid,{}))
        if not info or (role!='developer' and not any(x in stores for x in info.get('stores',[]))): continue
        if val not in ['','Shift A','Shift B','Weekly Off']: continue
        roster.setdefault(uid,{})
        if val: roster[uid][dt]=val
        else: roster[uid].pop(dt,None)
        saved+=1
    save_roster(roster)
    return jsonify({'status':'success','message':f'{saved} roster cells saved'})

@app.route('/export_roster_excel')
def export_roster_excel():
    if not session_has_permission('roster'): return redirect(url_for('login'))
    month=request.args.get('month',datetime.now().strftime('%Y-%m'))
    try: first=datetime.strptime(month+'-01','%Y-%m-%d')
    except ValueError: return 'Invalid month',400
    next_month=(first.replace(day=28)+timedelta(days=4)).replace(day=1); days=(next_month-first).days
    dates=[first+timedelta(days=i) for i in range(days)]
    roster=load_roster(); db=load_users_db(); stores=get_user_stores(); role=session.get('role')
    employees=[]
    for uid,raw in db.items():
        info=normalize_user_record(raw)
        if info.get('role')!='employee' or info.get('status')=='blocked': continue
        if role!='developer' and (not stores or stores[0] not in info.get('stores',[])): continue
        employees.append((uid,info))
    employees.sort(key=lambda x:x[1].get('name',x[0]))
    wb=openpyxl.Workbook(); ws=wb.active; ws.title=f'Roster {month}'
    ws.sheet_view.showGridLines=False; ws.freeze_panes='E4'
    last_col=4+len(dates)
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=last_col)
    ws.cell(1,1,f'Monthly Roster Planner - {first.strftime("%B %Y")}')
    ws.cell(1,1).font=Font(size=16,bold=True,color='FFFFFF'); ws.cell(1,1).fill=PatternFill('solid',fgColor='17365D'); ws.cell(1,1).alignment=Alignment(horizontal='center')
    headers=['Sr. Number','Employee Code','Employee Name','Designation']+[f'{d.day}\n{d.strftime("%a")}' for d in dates]
    for col,val in enumerate(headers,1):
        c=ws.cell(3,col,val); c.font=Font(bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor='1F4E78'); c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    counts={d.strftime('%Y-%m-%d'):{'Shift A':0,'Shift B':0,'Weekly Off':0} for d in dates}
    for row_idx,(uid,info) in enumerate(employees,4):
        ws.cell(row_idx,1,row_idx-3); ws.cell(row_idx,2,uid); ws.cell(row_idx,3,info.get('name',uid)); ws.cell(row_idx,4,info.get('designation') or info.get('dept','Employee'))
        for j,d in enumerate(dates,5):
            dt=d.strftime('%Y-%m-%d'); val=roster.get(uid,{}).get(dt,'Default') or 'Default'; ws.cell(row_idx,j,val); ws.cell(row_idx,j).alignment=Alignment(horizontal='center')
            if val in counts[dt]: counts[dt][val]+=1
            color={'Shift A':'C6EFCE','Shift B':'D9EAF7','Weekly Off':'FFC7CE'}.get(val,'FFFFFF'); ws.cell(row_idx,j).fill=PatternFill('solid',fgColor=color)
    summary_row=4+len(employees)+1
    ws.merge_cells(start_row=summary_row,start_column=1,end_row=summary_row,end_column=4); ws.cell(summary_row,1,'Daily Summary: A / B / Off'); ws.cell(summary_row,1).font=Font(bold=True)
    for j,d in enumerate(dates,5):
        c=counts[d.strftime('%Y-%m-%d')]; ws.cell(summary_row,j,f"{c['Shift A']} / {c['Shift B']} / {c['Weekly Off']}"); ws.cell(summary_row,j).font=Font(bold=True); ws.cell(summary_row,j).alignment=Alignment(horizontal='center')
    ws.column_dimensions['A'].width=12; ws.column_dimensions['B'].width=16; ws.column_dimensions['C'].width=30; ws.column_dimensions['D'].width=22
    for j in range(5,last_col+1): ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width=13
    thin=Side(style='thin',color='D9E2F3')
    for row in ws.iter_rows(min_row=3,max_row=summary_row,min_col=1,max_col=last_col):
        for c in row: c.border=Border(left=thin,right=thin,top=thin,bottom=thin); c.alignment=Alignment(vertical='center',wrap_text=True,horizontal=c.alignment.horizontal or 'left')
    output=io.BytesIO(); wb.save(output); output.seek(0)
    return send_file(output,as_attachment=True,download_name=f'Roster_{month}.xlsx',mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# --- ROSTER PLANNER APIs ---
@app.route('/api/get_roster', methods=['GET'])
def api_get_roster():
    if not session_has_permission('roster'): return {}, 403
    emp_id = request.args.get('emp_id')
    roster = load_roster()
    return roster.get(emp_id, {})

@app.route('/api/save_roster', methods=['POST'])
def api_save_roster():
    if not session_has_permission('roster'): return {}, 403
    data = request.get_json()
    emp_id = data.get('emp_id')
    updates = data.get('updates', {})
    
    roster = load_roster()
    if emp_id not in roster: roster[emp_id] = {}
    
    for dt, val in updates.items():
        roster[emp_id][dt] = val
        
    save_roster(roster)
    return {'status': 'success'}

# --- EMPLOYEE SHIFT REQUEST APIs ---
@app.route('/request_shift', methods=['POST'])
def request_shift():
    if session.get('role') != 'employee': return redirect(url_for('index'))
    req_date = request.form.get('req_date')
    req_status = request.form.get('req_status')
    req_reason = request.form.get('req_reason')
    
    reqs = load_shift_requests()
    req_id = len(reqs) + 1
    reqs.append({
        'id': req_id, 'user_id': session.get('user_id'), 'name': session.get('user_name'),
        'date': req_date, 'requested': req_status, 'reason': req_reason, 'status': 'Pending'
    })
    save_shift_requests(reqs)
    flash("Shift/Off request submit ho gayi hai! Admin approval ka wait karein.", "success")
    return redirect(url_for('index'))

@app.route('/update_shift_req/<int:req_id>/<action>')
def update_shift_req(req_id, action):
    if not session_has_permission('shift_approvals'): return redirect(url_for('login'))
    reqs = load_shift_requests()
    for req in reqs:
        if req['id'] == req_id:
            if action == 'approve':
                req['status'] = 'Approved'
                # Update main roster
                roster = load_roster()
                emp = req['user_id']
                if emp not in roster: roster[emp] = {}
                roster[emp][req['date']] = req['requested']
                save_roster(roster)
                flash(f"{req['name']} ki shift request approve kar di gayi hai.", "success")
            elif action == 'reject':
                req['status'] = 'Rejected'
                flash(f"{req['name']} ki shift request reject ki gayi.", "success")
            break
    save_shift_requests(reqs)
    return redirect(url_for('index'))

@app.route('/overtime_action/<rid>/<action>')
def overtime_action(rid,action):
    if session.get('role') not in ['admin','developer']:return redirect(url_for('login'))
    rows=load_overtime_approvals();rec=next((x for x in rows if x.get('id')==rid),None)
    if not rec:return redirect(url_for('index'))
    if session.get('role')!='developer' and rec.get('store') not in get_user_stores():return 'Unauthorized',403
    if action not in ['approve','reject']:return 'Invalid action',400
    rec['status']='Approved' if action=='approve' else 'Rejected';rec['reviewed_by']=session.get('user_name');rec['reviewed_at']=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    save_overtime_approvals(rows)
    mins=int(rec.get('minutes',0))
    if rec['status']=='Approved':
        suite_notify(rec['user_id'],'Overtime approved',f"{mins} extra minutes on {rec['date']} were approved by {rec['reviewed_by']}.",rec.get('store',''))
        flash(f"Overtime approved for {rec['name']}: {mins} minutes.",'success')
    else:
        # Rejection remains in manager history only. No employee notification is generated.
        flash(f"Overtime rejected for {rec['name']}. Credited working time adjusted to 8h.",'success')
    return redirect(url_for('index'))

# --- QUICK EDIT API ---
@app.route('/quick_edit')
def quick_edit():
    if not session_has_permission('attendance_edit'): return redirect(url_for('index'))
    date_str = request.args.get('date')
    emp_id = request.args.get('emp_id')
    field = request.args.get('field')
    new_time = request.args.get('time', '').strip()

    overrides = load_overrides()
    if date_str not in overrides: overrides[date_str] = {}
    if emp_id not in overrides[date_str]: overrides[date_str][emp_id] = {}

    if new_time == '' or new_time == '-':
        if field in overrides[date_str][emp_id]: del overrides[date_str][emp_id][field]
    else:
        if len(new_time.split(':')) == 2: new_time += ':00'
        try:
            datetime.strptime(new_time,'%H:%M:%S')
        except ValueError:
            flash('Invalid time. Use HH:MM or HH:MM:SS.','danger');return redirect(request.referrer or url_for('index'))
        overrides[date_str][emp_id][field] = new_time

    save_overrides(overrides)
    flash(f"Time successfully updated for {emp_id}!", "success")
    ref = request.referrer
    return redirect(ref) if ref else redirect(url_for('index'))

# --- LEAVE APIs ---
@app.route('/apply_leave', methods=['POST'])
def apply_leave():
    if not session.get('logged_in') or session.get('role') != 'employee': return redirect(url_for('login'))
    user_id = session.get('user_id')
    name = session.get('user_name')
    start_date = request.form.get('start_date')
    end_date = request.form.get('end_date')
    leave_type = request.form.get('leave_type', 'F10;1')
    if leave_type not in LEAVE_CODE_NAMES:
        flash('Invalid leave code selected.', 'danger')
        return redirect(url_for('index'))
    leave_reason = LEAVE_CODE_NAMES[leave_type]
    filename = None
    file = request.files.get('supporting_doc')
    if file and file.filename != '':
        if not allowed_file(file.filename):
            flash('Invalid file format! Sirf PDF, JPG, PNG ya DOC files allowed hain.', 'danger')
            return redirect(url_for('index'))
        filename = secure_filename(file.filename)
        local_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(local_path)
        upload_file_to_github(local_path, f"leave_documents/{user_id}_{filename}")
    save_leave_to_excel(user_id, name, start_date, end_date, leave_type, filename if filename else "No Document")
    LEAVE_REQUESTS.append({
        'id': len(LEAVE_REQUESTS) + 1, 'user_id': user_id, 'name': name, 'start_date': start_date,
        'end_date': end_date, 'leave_type': leave_type, 'leave_reason': leave_reason, 'filename': filename, 'status': 'Pending', 'store': session.get('store','LM11')
    })
    save_leave_requests(LEAVE_REQUESTS)
    flash('Aapki leave request successfully submit ho gayi hai!', 'success')
    return redirect(url_for('index'))

@app.route('/update_leave/<int:req_id>/<action>')
def update_leave(req_id, action):
    if not session_has_permission('leave_management'): return redirect(url_for('login'))
    for req in LEAVE_REQUESTS:
        if req['id'] == req_id:
            if session.get('role')!='developer' and req.get('store') not in get_user_stores(): flash('Aap is store ki leave approve nahi kar sakte.','danger'); return redirect(url_for('index'))
            if action == 'approve':
                req['status'] = 'Approved'
                flash(f"Leave request for {req['name']} approved successfully!", 'success')
            elif action == 'reject':
                req['status'] = 'Rejected'
                flash(f"Leave request for {req['name']} rejected.", 'success')
            break
    save_leave_requests(LEAVE_REQUESTS)
    return redirect(url_for('index'))

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    if not session.get('logged_in'): return redirect(url_for('login'))
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

@app.route('/export_last_six_months')
def export_last_six_months():
    if session.get('role') not in ['admin','developer'] or not session_has_permission('payroll'):
        return redirect(url_for('index'))
    today=datetime.now().date();start_month=today.month-5;start_year=today.year
    while start_month<=0:start_month+=12;start_year-=1
    start_date=f'{start_year:04d}-{start_month:02d}-01';end_date=today.strftime('%Y-%m-%d')
    wb=openpyxl.Workbook();ws=wb.active;ws.title='Six Month Attendance'
    headers=['Store','Date','Employee ID','Employee Name','Department','Store In','Lunch Out','Lunch In','Out Time','Lunch Duration','Credited Working Hours','Overtime Status','Credited Overtime','Attendance Status']
    ws.append(headers)
    for c in ws[1]:c.font=openpyxl.styles.Font(bold=True,color='FFFFFF');c.fill=openpyxl.styles.PatternFill('solid',fgColor='0F172A')
    stores=get_user_stores() if session.get('role')=='admin' else [x for x in MACHINES if x!='DEV']
    for st in stores:
        try:logs,*_=fetch_attendance_data(start_date,end_date,'ALL',st)
        except Exception as e:
            ws.append([st,'ERROR','','',str(e)]);continue
        for x in sorted(logs,key=lambda r:(r.get('date',''),r.get('name',''))):
            ws.append([st,x.get('date'),x.get('user_id'),x.get('name'),x.get('dept'),x.get('store_in'),x.get('lunch_out'),x.get('lunch_in'),x.get('out_time'),x.get('total_lunch'),x.get('total_hours'),x.get('overtime_status','None'),(f"{int(x.get('overtime_minutes',0))//60}h" if x.get('overtime_status')=='Approved' else '0h'),x.get('status')])
    ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
    widths=[12,13,16,30,22,12,12,12,12,16,22,18,18,18]
    for i,w in enumerate(widths,1):ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width=w
    summary=wb.create_sheet('Rules & Summary');summary.append(['Export Period',start_date,end_date]);summary.append(['Normal shift with valid lunch','8 hours']);summary.append(['Normal shift without lunch','7 hours']);summary.append(['Overtime qualification','Minimum 45 extra minutes = 1 credited hour']);summary.append(['Access','Admin: assigned stores; Developer: all stores'])
    out=io.BytesIO();wb.save(out);out.seek(0)
    return send_file(out,as_attachment=True,download_name=f'Attendance_Last_6_Months_{start_date}_to_{end_date}.xlsx',mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/export')
def export_excel():
    if not session_has_permission('payroll'): return redirect(url_for('login'))
    start_date = request.args.get('start_date', datetime.now().strftime('%Y-%m-%d'))
    end_date = request.args.get('end_date', datetime.now().strftime('%Y-%m-%d'))
    selected_emp = request.args.get('employee', 'ALL')
    store = session.get('store', 'LM11')
    
    logs, _, g_hrs, g_l_hrs, g_var, _, _ = fetch_attendance_data(start_date, end_date, selected_emp, store)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Attendance Report"
    ws.append(["Attendance Portal - Attendance Report"])
    ws.append([f"Period: {start_date} to {end_date}"])
    ws.append([])
    ws.append(["Sr. No.", "Date", "ID", "Employee Name", "Department", "Store In", "Lunch Out", "Lunch In", "Out Time", "Total Lunch", "Working Hours", "Total Hora Extra", "Status"])
    for idx, log in enumerate(logs, 1):
        ws.append([idx, log['date'], log['user_id'], log['name'], log['dept'], log['store_in'], log['lunch_out'], log['lunch_in'], log['out_time'], log['total_lunch'], log['total_hours'], log['net_variance'], log['status']])
    ws.append([])
    ws.append(["", "", "", "", "", "", "", "", "Total Summary:", g_l_hrs, g_hrs, g_var])
    excel_io = io.BytesIO()
    wb.save(excel_io)
    excel_io.seek(0)
    return send_file(excel_io, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name=f"Attendance_Report_{start_date}_to_{end_date}.xlsx")

@app.route('/export_matrix')
def export_matrix():
    if not session_has_permission('payroll'): return redirect(url_for('login'))
    start_date = request.args.get('start_date', datetime.now().strftime('%Y-%m-%d'))
    end_date = request.args.get('end_date', datetime.now().strftime('%Y-%m-%d'))
    store = session.get('store', 'LM11')
    
    start_dt, end_dt = datetime.strptime(start_date, '%Y-%m-%d'), datetime.strptime(end_date, '%Y-%m-%d')
    date_list = []
    while start_dt <= end_dt:
        date_list.append(start_dt.strftime('%Y-%m-%d'))
        start_dt += timedelta(days=1)
        
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Employee Matrix"
    ws.append(["Attendance Portal - Employee Matrix Attendance Report"])
    ws.append([f"Period: {start_date} to {end_date}"])
    ws.append([])
    ws.append(["ID", "Employee Name", "Department"] + date_list)
    
    db = load_users_db()
    
    for emp_code, emp_data in sorted(db.items(), key=lambda x: x[1]['name']):
        if emp_data.get('role') != 'employee': continue
        if store != 'DEV' and emp_data.get('store') != store: continue
        
        row = [emp_code, emp_data['name'], emp_data.get('dept', '')]
        for d_str in date_list:
            logs_d, _, _, _, _, _, _ = fetch_attendance_data(d_str, d_str, emp_code, store)
            if logs_d:
                st = logs_d[0]
                status = st['status']
                if any(code in status for code in ['F01;1', 'F03;1', 'F05;1', 'F10;1', 'F51;1', 'F60;1', 'F61;1', 'F62;1']):
                    row.append(next((code for code in ['F01;1', 'F03;1', 'F05;1', 'F10;1', 'F51;1', 'F60;1', 'F61;1', 'F62;1'] if code in status), 'F10;1'))
                elif status == 'Weekly Off': row.append('Off')
                elif status == 'Present': row.append(st['net_variance'] if 'H06;' in st['net_variance'] or 'H07;' in st['net_variance'] else 'P')
                elif status == 'Mis Punch': row.append('Mis Punch')
                else: row.append('F03;1')
            else:
                curr_dt_obj = datetime.strptime(d_str, '%Y-%m-%d')
                row.append('Off' if curr_dt_obj.strftime('%A').upper() == emp_data.get('off', '').upper() else 'F03;1')
        ws.append(row)
        
    excel_io = io.BytesIO()
    wb.save(excel_io)
    excel_io.seek(0)
    return send_file(excel_io, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name=f"Employee_Matrix_{start_date}_to_{end_date}.xlsx")

@app.route('/shutdown')
def shutdown():
    dev_pass = os.getenv('DEV_PWD', 'Shama@8577')
    if session.get('role') == 'developer' and dev_pass and request.args.get('pwd') == dev_pass:
        func = request.environ.get('werkzeug.server.shutdown')
        if func: func()
        else: sys.exit(0)
        return "Server successfully shutdown ho gaya hai."
    return "Unauthorized access!", 403

@app.route('/api/attendance/sync', methods=['POST'])
def sync_attendance():
    global SYNCED_ATTENDANCE_LOGS, LAST_DEVICE_SYNC_TIME
    try:
        data = request.get_json()
        if not data or 'logs' not in data: return {'status': 'error', 'message': 'No logs'}, 400
        logs = data['logs']
        store_code = data.get('store_code', 'LM11')
        
        existing_keys = {(str(item.get('user_id')), str(item.get('timestamp'))) for item in SYNCED_ATTENDANCE_LOGS}
        added_count = 0
        for log in logs:
            key = (str(log.get('user_id')), str(log.get('timestamp')))
            if key not in existing_keys:
                SYNCED_ATTENDANCE_LOGS.append(log)
                existing_keys.add(key)
                added_count += 1
                
        # Dictionary Update
        LAST_DEVICE_SYNC_TIME[store_code] = datetime.now()
        
        return {'status': 'success', 'message': f'{len(logs)} records synced ({added_count} new)'}, 200
    except Exception as e: return {'status': 'error', 'message': str(e)}, 500

# === NAYA ID CARD ROUTE ===
@app.route('/employee_id/<emp_code>')
def view_employee_id(emp_code):
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    
    # Security: Employee can only see their own ID, Admin/Developer can see anyone's
    if session.get('role') == 'employee' and session.get('user_id') != emp_code:
        return "Unauthorized Access", 403

    db = load_users_db()
    emp_data = db.get(emp_code)
    
    if not emp_data:
        # Fallback agar user database me available nahi hai par session mein hai
        if emp_code == session.get('user_id'):
            emp_data = {'name': session.get('user_name', 'Employee'), 'profile_photo': ''}
        else:
            return "Employee not found", 404
            
    emp_data = normalize_user_record(dict(emp_data))
    emp_data['emp_code'] = emp_code
    emp_data.update(EMPLOYEE_IDENTITY_DATA.get(emp_code.upper(), {}))
    return render_template_string(ID_CARD_TEMPLATE, employee=emp_data)



BUSINESS_PAGE_TEMPLATE='''<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{{title}}</title><script src="https://cdn.tailwindcss.com"></script><style>th,td{padding:.65rem;text-align:left;vertical-align:top}img{image-orientation:from-image}input,button{min-height:40px}@media(max-width:640px){main{padding:.65rem!important}.wide{min-width:720px}}</style></head><body class="bg-slate-100 text-slate-800"><header class="bg-slate-950 text-white p-4 flex flex-col sm:flex-row gap-3 sm:items-center sm:justify-between"><div><h1 class="text-xl font-black">{{icon}} {{title}}</h1><p class="text-xs text-slate-300">{{subtitle}}</p></div><div class="flex gap-2"><a href="/" class="rounded-xl bg-white/10 px-4 py-2 text-sm font-bold">← Dashboard</a></div></header><main class="p-5 space-y-4">{{body|safe}}</main><script id="universal-table-sort">
(function(){
  function value(cell,type){const text=(cell?.innerText||'').trim();if(type==='number'){const n=parseFloat(text.replace(/[^0-9.-]/g,''));return Number.isNaN(n)?-Infinity:n;}return text.toLocaleLowerCase();}
  function enhance(root=document){root.querySelectorAll('table').forEach((table,ti)=>{if(table.dataset.sortReady)return;table.dataset.sortReady='1';const heads=table.querySelectorAll('thead th');heads.forEach((th,ci)=>{if(th.dataset.noSort==='1')return;th.style.cursor='pointer';th.style.userSelect='none';if(!/[↕↑↓]/.test(th.textContent))th.insertAdjacentText('beforeend',' ↕');th.title='Click to sort';th.addEventListener('click',()=>{const tbody=table.tBodies[0];if(!tbody)return;const rows=Array.from(tbody.rows).filter(r=>r.cells.length>ci&&!r.querySelector('[colspan]'));const asc=th.dataset.direction!=='asc';heads.forEach(h=>{h.dataset.direction='';h.textContent=h.textContent.replace(/ [↑↓]$/,' ↕');});th.dataset.direction=asc?'asc':'desc';th.textContent=th.textContent.replace(/ ↕$/,'')+(asc?' ↑':' ↓');const type=th.dataset.sortType||((rows.every(r=>/^[-+]?\d[\d.,]*$/.test((r.cells[ci]?.innerText||'').trim())))?'number':'text');rows.sort((a,b)=>{const av=value(a.cells[ci],type),bv=value(b.cells[ci],type);return (av>bv?1:av<bv?-1:0)*(asc?1:-1);});rows.forEach(r=>tbody.appendChild(r));});});});}
  window.applyKviFilters=function(){const q=(document.getElementById('kvi-search')?.value||'').trim().toLowerCase();const type=document.getElementById('kvi-type-filter')?.value||'ALL';document.querySelectorAll('.kvi-row').forEach(r=>{let rt=(r.dataset.type||'OTHER').trim().toUpperCase();if(rt.includes('KVI')&&rt.includes('POWER'))rt='KVI+POWER';else if(rt.includes('KVI'))rt='KVI';else if(rt.includes('POWER'))rt='POWER SKU';const typeOk=type==='ALL'||(type==='OTHER'&&!['KVI','KVI+POWER','POWER SKU'].includes(rt))||rt===type;const searchOk=!q||(r.dataset.search||'').includes(q);r.style.display=typeOk&&searchOk?'':'none';});};
  window.filterKviType=function(type){const sel=document.getElementById('kvi-type-filter');if(sel)sel.value=type;window.applyKviFilters();document.getElementById('kvi-article-table')?.scrollIntoView({behavior:'smooth',block:'start'});};
  document.addEventListener('DOMContentLoaded',()=>enhance(document));window.enhanceSortableTables=enhance;
})();
</script><script>function toggleAllKvi(master){document.querySelectorAll('.kvi-row').forEach(r=>{const cb=r.querySelector('.kvi-select');if(cb&&r.style.display!=='none')cb.checked=master.checked;});}
function deleteSelectedKvi(){const count=document.querySelectorAll('.kvi-select:checked').length;if(!count){showToast('Select at least one article.','warning');return;}if(confirm('Delete '+count+' selected article(s)? This cannot be undone.')){document.getElementById('kvi-delete-action').value='selected';document.getElementById('kvi-delete-form').submit();}}
function deleteAllKvi(){if(confirm('Delete ALL existing KVI, KVI+Power and Power SKU article data? This cannot be undone.')){document.getElementById('kvi-delete-action').value='all';document.getElementById('kvi-delete-form').submit();}}
function showSelectedFiles(input,targetId){const box=document.getElementById(targetId);const files=Array.from(input.files||[]);if(!box)return;box.textContent=files.length?files.length+' file(s) selected: '+files.map(f=>f.name).join(', '):'You can select multiple files at once.';}</script>
<style id="hrms-toast-style">
#hrms-toast-host{position:fixed;top:18px;right:18px;z-index:99999;width:min(420px,calc(100vw - 36px));display:flex;flex-direction:column;gap:10px;pointer-events:none}
.hrms-toast{pointer-events:auto;display:flex;align-items:flex-start;gap:12px;padding:14px 16px;border-radius:14px;color:#fff;font:700 12px/1.45 Inter,system-ui,sans-serif;box-shadow:0 18px 45px rgba(15,23,42,.24);transform:translateX(0);opacity:1;transition:opacity .3s ease,transform .3s ease;overflow-wrap:anywhere}
.hrms-toast-success{background:#059669}.hrms-toast-danger,.hrms-toast-error{background:#e11d48}.hrms-toast-warning{background:#d97706}.hrms-toast-info{background:#2563eb}
.hrms-toast-hide{opacity:0;transform:translateX(30px)}.hrms-toast-message{flex:1}.hrms-toast-close{border:0;background:transparent;color:#fff;font-size:21px;line-height:1;cursor:pointer;min-height:0;padding:0 0 0 6px}
@media(max-width:640px){ #hrms-toast-host{top:10px;right:10px;width:calc(100vw - 20px)}}
</style>
<script id="hrms-toast-script">
(function(){
  function host(){let h=document.getElementById('hrms-toast-host');if(!h){h=document.createElement('div');h.id='hrms-toast-host';h.setAttribute('aria-live','polite');h.setAttribute('aria-atomic','false');document.body.appendChild(h)}return h}
  window.showToast=function(message,type='success'){
    if(message===undefined||message===null||String(message).trim()==='')return;
    const normalized=(type==='error'?'danger':type)||'info';
    const t=document.createElement('div');t.className='hrms-toast hrms-toast-'+normalized;t.setAttribute('role',normalized==='danger'?'alert':'status');
    const m=document.createElement('div');m.className='hrms-toast-message';m.textContent=String(message);
    const x=document.createElement('button');x.type='button';x.className='hrms-toast-close';x.innerHTML='&times;';x.setAttribute('aria-label','Close notification');
    const remove=()=>{if(t.dataset.closing)return;t.dataset.closing='1';t.classList.add('hrms-toast-hide');setTimeout(()=>t.remove(),300)};
    x.addEventListener('click',remove);t.append(m,x);host().appendChild(t);setTimeout(remove,5000);
  };
  document.addEventListener('DOMContentLoaded',()=>document.querySelectorAll('.hrms-flash').forEach(el=>{showToast(el.dataset.message,el.dataset.category);el.remove()}));
})();
</script>
</body></html>'''

def business_page(title,icon,subtitle,body):
    return render_template_string(BUSINESS_PAGE_TEMPLATE,title=title,icon=icon,subtitle=subtitle,body=body)

def _business_docs(path):
    data=load_json_file(path);return data if isinstance(data,list) else []

def _save_business_docs(path,data):save_json_file(path,data[-1000:])

def _business_upload(meta_file,category):
    if session.get('role') not in ['admin','developer']:return [],'Only Admin or Developer can upload documents.'
    files=[f for f in request.files.getlist('document') if f and f.filename]
    if not files:return [],'Select one or more documents first.'
    allowed={'.pdf','.xlsx','.xls','.doc','.docx','.png','.jpg','.jpeg','.webp'}
    bad=[f.filename for f in files if Path(f.filename).suffix.lower() not in allowed]
    if bad:return [],'Unsupported files: '+', '.join(bad)
    rows=_business_docs(meta_file);saved=[];common_title=(request.form.get('title') or '').strip();now=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    for position,f in enumerate(files,1):
        filename=secure_filename(f'{category}_{datetime.now().strftime("%Y%m%d%H%M%S")}_{uuid.uuid4().hex[:8]}_{f.filename}')
        f.save(os.path.join(BUSINESS_DOC_FOLDER,filename))
        file_title=(f'{common_title} - {position}' if common_title and len(files)>1 else (common_title or Path(f.filename).stem))
        item={'id':uuid.uuid4().hex,'category':category,'filename':filename,'original':f.filename,'title':file_title,'period_id':(request.form.get('period_id') or '').strip(),'uploaded_by':session.get('user_name'),'uploaded_at':now,'batch_size':len(files),'batch_position':position};rows.append(item);saved.append(item)
    _save_business_docs(meta_file,rows)
    return saved,''

def _doc_cards(rows):
    if not rows:
        return '<div class="rounded-2xl border border-dashed bg-white p-8 text-center text-slate-400">No documents uploaded yet.</div>'
    image_ext={'.png','.jpg','.jpeg','.webp','.gif'}
    cards=[]
    for x in rows:
        ext=Path(x.get('original','')).suffix.lower()
        meta='<div class="border-t bg-white px-4 py-3"><b class="block text-sm">{}</b><span class="block text-xs text-slate-500">{}</span><span class="mt-1 block text-[10px] text-slate-400">{} · {}</span></div>'.format(x.get('title','Document'),x.get('original',''),x.get('uploaded_by',''),x.get('uploaded_at',''))
        if ext in image_ext:
            cards.append('<article class="overflow-hidden rounded-2xl border bg-white shadow-sm"><a href="/business_file/{}" target="_blank" title="Open full image"><img src="/business_file/{}" loading="lazy" alt="{}" class="block h-auto w-full bg-white object-contain"></a>{}</article>'.format(x['id'],x['id'],x.get('title','Promotion image'),meta))
        else:
            cards.append('<a href="/business_file/{}" target="_blank" class="block rounded-2xl border bg-white p-4 shadow-sm hover:border-indigo-400"><b class="block text-sm">Document: {}</b><span class="mt-2 block text-xs text-slate-500">{}</span><span class="mt-2 block text-[10px] text-slate-400">{} · {}</span><span class="mt-3 inline-block rounded-lg bg-slate-900 px-3 py-2 text-xs font-bold text-white">Open Document</span></a>'.format(x['id'],x.get('title','Document'),x.get('original',''),x.get('uploaded_by',''),x.get('uploaded_at','')))
    return '<div class="mx-auto flex w-full max-w-5xl flex-col gap-5">'+''.join(cards)+'</div>'

@app.route('/duplicate_punches')
def duplicate_punches_page():
    if session.get('role') not in ['admin','developer']:return redirect(url_for('index'))
    rows=load_json_file(DUPLICATE_PUNCHES_FILE)
    if not isinstance(rows,list):rows=[]
    if session.get('role')=='admin':rows=[x for x in rows if x.get('store') in get_user_stores()]
    uid=(request.args.get('uid') or '').strip().upper();date=(request.args.get('date') or '').strip()
    if uid:rows=[x for x in rows if uid in str(x.get('user_id','')).upper() or uid in str(x.get('name','')).upper()]
    if date:rows=[x for x in rows if x.get('date')==date]
    trs=''.join(f"<tr class='border-b'><td>{x.get('date','')}</td><td><b>{x.get('name','')}</b><br>{x.get('user_id','')}</td><td>{x.get('store','')}</td><td>{x.get('kept_time','')}</td><td class='font-bold text-rose-600'>{x.get('duplicate_time','')}</td><td>{x.get('gap_seconds',0)} sec</td><td>{x.get('reason','')}</td></tr>" for x in reversed(rows))
    body=f'''<form class="grid grid-cols-1 gap-2 rounded-2xl border bg-white p-4 sm:grid-cols-[1fr_1fr_auto]"><input name="uid" value="{uid}" placeholder="Employee name or code" class="rounded-xl border px-3 py-2"><input type="date" name="date" value="{date}" class="rounded-xl border px-3 py-2"><button class="rounded-xl bg-orange-600 px-5 py-2 font-black text-white">Filter</button></form><div class="rounded-xl bg-orange-50 p-3 text-sm text-orange-900">The first punch is retained. Any later punch by the same employee within 60 seconds is stored here and excluded from attendance calculation.</div><div class="overflow-auto rounded-2xl border bg-white"><table class="wide w-full text-sm"><thead class="bg-slate-900 text-white"><tr><th>Date</th><th>Employee</th><th>Store</th><th>Kept Punch</th><th>Duplicate Punch</th><th>Gap</th><th>Reason</th></tr></thead><tbody>{trs or '<tr><td colspan="7" class="p-8 text-center text-slate-400">No duplicate punches found.</td></tr>'}</tbody></table></div>'''
    return business_page('Duplicate Punches','👆','Automatic one-minute duplicate filtering',body)

@app.route('/app_version')
def app_version():
    return {'build':APP_BUILD,'offer_periods':True,'promotion_period_route':'/promotion/period/<period_id>'}

@app.route('/promotion',methods=['GET','POST'])
def promotion_page():
    if not session.get('logged_in'):return redirect(url_for('login'))
    periods=_business_docs(PROMOTION_PERIODS_FILE)
    rows=_business_docs(PROMOTION_DOCUMENTS_FILE)
    can_manage=session.get('role') in ['admin','developer']
    if request.method=='POST':
        if not can_manage:return 'Only Admin or Developer can manage offers.',403
        action=request.form.get('action','upload')
        if action=='create_period':
            title=(request.form.get('period_title') or '').strip();start_date=request.form.get('start_date','');end_date=request.form.get('end_date','')
            if not title:flash('Offer period title is required.','danger')
            else:
                periods.append({'id':uuid.uuid4().hex,'title':title,'start_date':start_date,'end_date':end_date,'created_by':session.get('user_name'),'created_at':datetime.now().strftime('%Y-%m-%d %H:%M:%S')});_save_business_docs(PROMOTION_PERIODS_FILE,periods);flash('Offer period created.','success')
            return redirect(url_for('promotion_page'))
        saved,msg=_business_upload(PROMOTION_DOCUMENTS_FILE,'promotion')
        if not msg:flash(f'{len(saved)} promotion picture(s) uploaded.','success');return redirect(url_for('promotion_page'))
        flash(msg,'danger');return redirect(url_for('promotion_page'))
    # Existing pictures without a period remain accessible under a generated link.
    unassigned=sum(1 for x in rows if not x.get('period_id'))
    cards=[]
    for period in reversed(periods):
        count=sum(1 for x in rows if x.get('period_id')==period.get('id'))
        dates=''
        if period.get('start_date') or period.get('end_date'):dates=f"<span class='block text-xs text-slate-500'>{period.get('start_date','')} → {period.get('end_date','')}</span>"
        admin=''
        if can_manage:
            admin=f'''<div class="mt-3 flex gap-2"><a href="/promotion/period/{period['id']}/edit" class="rounded-lg bg-indigo-600 px-3 py-2 text-xs font-bold text-white">Edit / Upload</a><form method="post" action="/promotion/period/{period['id']}/delete" onsubmit="return confirm('Delete this period and all its pictures?')"><button class="rounded-lg bg-rose-600 px-3 py-2 text-xs font-bold text-white">Delete</button></form></div>'''
        cards.append(f'''<article class="rounded-2xl border bg-white p-5 shadow-sm"><a href="/promotion/period/{period['id']}" class="block"><span class="text-2xl">🗓️</span><h3 class="mt-2 text-lg font-black text-fuchsia-800">{period.get('title','Offer Period')}</h3>{dates}<span class="mt-3 inline-block rounded-full bg-fuchsia-100 px-3 py-1 text-xs font-bold text-fuchsia-700">{count} picture(s)</span></a>{admin}</article>''')
    if unassigned:
        cards.append(f'''<article class="rounded-2xl border bg-white p-5 shadow-sm"><a href="/promotion/period/unassigned" class="block"><span class="text-2xl">🖼️</span><h3 class="mt-2 text-lg font-black text-slate-800">Previous / Unassigned Offers</h3><span class="mt-3 inline-block rounded-full bg-slate-100 px-3 py-1 text-xs font-bold">{unassigned} picture(s)</span></a></article>''')
    admin_panel=''
    if can_manage:
        options=''.join(f"<option value='{x['id']}'>{x.get('title','Offer Period')}</option>" for x in periods)
        admin_panel=f'''<div class="grid grid-cols-1 gap-4 xl:grid-cols-2"><form method="post" class="rounded-2xl border bg-white p-4"><input type="hidden" name="action" value="create_period"><h2 class="mb-3 font-black">Create Offer Period Link</h2><input name="period_title" placeholder="Example: Weekend Offer 10–15 October" required class="mb-2 w-full rounded-xl border px-3 py-2"><div class="grid grid-cols-2 gap-2"><input type="date" name="start_date" class="rounded-xl border px-3 py-2"><input type="date" name="end_date" class="rounded-xl border px-3 py-2"></div><button class="mt-3 w-full rounded-xl bg-fuchsia-600 py-2 font-black text-white">Create Offer Period</button></form><form method="post" enctype="multipart/form-data" class="rounded-2xl border bg-white p-4"><input type="hidden" name="action" value="upload"><h2 class="mb-3 font-black">Upload Pictures to Offer Period</h2><select name="period_id" required class="mb-2 w-full rounded-xl border px-3 py-2"><option value="">Select offer period</option>{options}</select><input name="title" placeholder="Picture title (optional)" class="mb-2 w-full rounded-xl border px-3 py-2"><input type="file" name="document" multiple accept=".jpg,.jpeg,.png,.webp,.gif" required class="w-full rounded-xl border p-2 text-xs"><button class="mt-3 w-full rounded-xl bg-indigo-600 py-2 font-black text-white">Upload Pictures</button></form></div>'''
    directory='<div class="grid grid-cols-1 gap-4 sm:grid-cols-2 xl:grid-cols-3">'+(''.join(cards) if cards else '<div class="col-span-full rounded-2xl border border-dashed bg-white p-8 text-center text-slate-400">No offer period created yet.</div>')+'</div>'
    body='<div class="flex flex-wrap items-center justify-between gap-2 rounded-xl bg-fuchsia-50 p-3 text-sm text-fuchsia-900"><span>Employees can open an offer period link and view pictures only. Admin and Developer can create, edit, upload and delete.</span><span class="rounded-full bg-fuchsia-200 px-3 py-1 text-[10px] font-black">OFFER PERIODS V2</span></div>'+admin_panel+'<section><h2 class="mb-3 font-black">Offer Period Links</h2>'+directory+'</section>'
    return business_page('Promotion','📣','Offer periods and promotion pictures',body)

@app.route('/promotion/period/<period_id>')
def promotion_period_view(period_id):
    if not session.get('logged_in'):return redirect(url_for('login'))
    periods=_business_docs(PROMOTION_PERIODS_FILE);rows=_business_docs(PROMOTION_DOCUMENTS_FILE)
    if period_id=='unassigned':title='Previous / Unassigned Offers';pictures=[x for x in rows if not x.get('period_id')]
    else:
        period=next((x for x in periods if x.get('id')==period_id),None)
        if not period:return 'Offer period not found',404
        title=period.get('title','Offer Period');pictures=[x for x in rows if x.get('period_id')==period_id]
    image_ext={'.jpg','.jpeg','.png','.webp','.gif'}
    images=[]
    for x in pictures:
        if Path(x.get('original','')).suffix.lower() in image_ext:
            images.append(f'''<a href="/business_file/{x['id']}" target="_blank" class="block overflow-hidden rounded-2xl bg-white shadow"><img src="/business_file/{x['id']}" class="h-auto w-full object-contain" loading="lazy" alt="Offer picture"></a>''')
    admin_link=''
    if session.get('role') in ['admin','developer'] and period_id!='unassigned':admin_link=f'<a href="/promotion/period/{period_id}/edit" class="rounded-xl bg-indigo-600 px-4 py-2 text-sm font-bold text-white">Edit Offer Period</a>'
    body=f'''<div class="flex justify-end">{admin_link}</div><div class="mx-auto flex w-full max-w-5xl flex-col gap-5">{''.join(images) if images else '<div class="rounded-2xl border border-dashed bg-white p-10 text-center text-slate-400">No pictures available in this offer period.</div>'}</div>'''
    return business_page(title,'🖼️','Promotion pictures',body)

@app.route('/promotion/period/<period_id>/edit',methods=['GET','POST'])
def promotion_period_edit(period_id):
    if session.get('role') not in ['admin','developer']:return redirect(url_for('promotion_period_view',period_id=period_id))
    periods=_business_docs(PROMOTION_PERIODS_FILE);period=next((x for x in periods if x.get('id')==period_id),None)
    if not period:return 'Offer period not found',404
    if request.method=='POST':
        action=request.form.get('action','update')
        if action=='update':
            period['title']=(request.form.get('period_title') or period.get('title')).strip();period['start_date']=request.form.get('start_date','');period['end_date']=request.form.get('end_date','');_save_business_docs(PROMOTION_PERIODS_FILE,periods);flash('Offer period updated.','success')
        elif action=='upload':
            saved,msg=_business_upload(PROMOTION_DOCUMENTS_FILE,'promotion')
            if msg:flash(msg,'danger')
            else:flash(f'{len(saved)} picture(s) uploaded.','success')
        return redirect(url_for('promotion_period_edit',period_id=period_id))
    pictures=[x for x in _business_docs(PROMOTION_DOCUMENTS_FILE) if x.get('period_id')==period_id]
    pic_cards=''.join(f'''<article class="overflow-hidden rounded-xl border bg-white"><img src="/business_file/{x['id']}" class="h-52 w-full object-contain"><div class="p-3"><b class="text-xs">{x.get('title','Picture')}</b><form method="post" action="/promotion/picture/{x['id']}/delete" onsubmit="return confirm('Delete this picture?')"><input type="hidden" name="period_id" value="{period_id}"><button class="mt-2 rounded-lg bg-rose-600 px-3 py-2 text-xs font-bold text-white">Delete Picture</button></form></div></article>''' for x in pictures)
    body=f'''<div class="grid grid-cols-1 gap-4 lg:grid-cols-2"><form method="post" class="rounded-2xl border bg-white p-4"><input type="hidden" name="action" value="update"><h2 class="mb-3 font-black">Edit Offer Period</h2><input name="period_title" value="{period.get('title','')}" required class="mb-2 w-full rounded-xl border px-3 py-2"><div class="grid grid-cols-2 gap-2"><input type="date" name="start_date" value="{period.get('start_date','')}" class="rounded-xl border px-3 py-2"><input type="date" name="end_date" value="{period.get('end_date','')}" class="rounded-xl border px-3 py-2"></div><button class="mt-3 w-full rounded-xl bg-indigo-600 py-2 font-black text-white">Save Changes</button></form><form method="post" enctype="multipart/form-data" class="rounded-2xl border bg-white p-4"><input type="hidden" name="action" value="upload"><input type="hidden" name="period_id" value="{period_id}"><h2 class="mb-3 font-black">Upload More Pictures</h2><input name="title" placeholder="Picture title (optional)" class="mb-2 w-full rounded-xl border px-3 py-2"><input type="file" name="document" multiple accept=".jpg,.jpeg,.png,.webp,.gif" required class="w-full rounded-xl border p-2 text-xs"><button class="mt-3 w-full rounded-xl bg-fuchsia-600 py-2 font-black text-white">Upload Pictures</button></form></div><section><h2 class="mb-3 font-black">Manage Pictures</h2><div class="grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-3">{pic_cards or '<div class="col-span-full rounded-xl border border-dashed bg-white p-8 text-center text-slate-400">No pictures uploaded.</div>'}</div></section>'''
    return business_page('Edit: '+period.get('title','Offer Period'),'✏️','Admin promotion management',body)

@app.route('/promotion/picture/<picture_id>/delete',methods=['POST'])
def promotion_picture_delete(picture_id):
    if session.get('role') not in ['admin','developer']:return 'Unauthorized',403
    rows=_business_docs(PROMOTION_DOCUMENTS_FILE);target=next((x for x in rows if x.get('id')==picture_id),None);period_id=request.form.get('period_id','')
    if target:
        path=os.path.join(BUSINESS_DOC_FOLDER,target.get('filename',''))
        if os.path.isfile(path):
            try:os.remove(path)
            except OSError:pass
        rows=[x for x in rows if x.get('id')!=picture_id];_save_business_docs(PROMOTION_DOCUMENTS_FILE,rows);flash('Picture deleted.','success')
    return redirect(url_for('promotion_period_edit',period_id=period_id))

@app.route('/promotion/period/<period_id>/delete',methods=['POST'])
def promotion_period_delete(period_id):
    if session.get('role') not in ['admin','developer']:return 'Unauthorized',403
    periods=[x for x in _business_docs(PROMOTION_PERIODS_FILE) if x.get('id')!=period_id];rows=_business_docs(PROMOTION_DOCUMENTS_FILE)
    for x in [a for a in rows if a.get('period_id')==period_id]:
        path=os.path.join(BUSINESS_DOC_FOLDER,x.get('filename',''))
        if os.path.isfile(path):
            try:os.remove(path)
            except OSError:pass
    rows=[x for x in rows if x.get('period_id')!=period_id];_save_business_docs(PROMOTION_PERIODS_FILE,periods);_save_business_docs(PROMOTION_DOCUMENTS_FILE,rows);flash('Offer period and pictures deleted.','success')
    return redirect(url_for('promotion_page'))


def _normalize_article_type(value):
    raw=str(value or '').strip().upper()
    t=' '.join(raw.replace('_',' ').replace('-',' ').replace('+',' ').split())
    has_kvi='KVI' in t
    has_power='POWER' in t
    if has_kvi and has_power:return 'KVI+POWER'
    if has_kvi:return 'KVI'
    if has_power:return 'POWER SKU'
    return t or 'OTHER'

def _read_kvi_excel(path):
    wb=openpyxl.load_workbook(path,data_only=True,read_only=True);ws=wb.active
    rows=list(ws.iter_rows(values_only=True))
    if not rows:return []
    header_index=0
    for i,row in enumerate(rows[:15]):
        text=' '.join(str(x or '').lower() for x in row)
        if 'article' in text and ('code' in text or 'name' in text):header_index=i;break
    header=[str(x or '').strip().lower() for x in rows[header_index]]
    def col(words,fallback):
        for i,h in enumerate(header):
            if any(w in h for w in words):return i
        return fallback
    type_i=col(['article type','type','tipo'],0);code_i=col(['article code','code','codigo','código'],1);name_i=col(['article name','name','description','nome'],2)
    data=[]
    for row in rows[header_index+1:]:
        vals=list(row)
        if not any(v not in [None,''] for v in vals):continue
        typ=_normalize_article_type(vals[type_i] if type_i<len(vals) else '')
        code=str(vals[code_i] if code_i<len(vals) and vals[code_i] is not None else '').strip()
        name=str(vals[name_i] if name_i<len(vals) and vals[name_i] is not None else '').strip()
        if not code and not name:continue
        data.append({'id':uuid.uuid4().hex,'article_type':typ,'article_code':code,'article_name':name})
    return data

@app.route('/kvi_power_sku/delete',methods=['POST'])
def delete_kvi_articles():
    if session.get('role') not in ['admin','developer']:
        return 'Only Admin or Developer can delete KVI data.',403
    payload=load_json_file(KVI_POWER_DATA_FILE)
    if not isinstance(payload,dict):payload={'items':[]}
    items=payload.get('items',[]) if isinstance(payload.get('items',[]),list) else []
    action=request.form.get('action','selected')
    if action=='all':
        deleted=len(items);payload['items']=[]
    else:
        selected=set(request.form.getlist('article_ids'))
        if not selected:
            flash('Select at least one article to delete.','danger');return redirect(url_for('kvi_power_page'))
        before=len(items);payload['items']=[x for x in items if x.get('id') not in selected];deleted=before-len(payload['items'])
    payload['updated_at']=datetime.now().strftime('%Y-%m-%d %H:%M:%S');payload['updated_by']=session.get('user_name')
    save_json_file(KVI_POWER_DATA_FILE,payload)
    flash(f'{deleted} article(s) deleted successfully.','success')
    return redirect(url_for('kvi_power_page'))

@app.route('/kvi_power_sku',methods=['GET','POST'])
def kvi_power_page():
    if not session.get('logged_in'):return redirect(url_for('login'))
    msg=''
    if request.method=='POST':
        if session.get('role') not in ['admin','developer']:return 'Only Admin or Developer can upload.',403
        action=request.form.get('action','document')
        if action=='excel':
            f=request.files.get('excel')
            if not f or not f.filename.lower().endswith('.xlsx'):msg='Upload an XLSX file.'
            else:
                fn=secure_filename(f'kvi_power_{datetime.now().strftime("%Y%m%d%H%M%S")}_{f.filename}');path=os.path.join(BUSINESS_DOC_FOLDER,fn);f.save(path)
                try:
                    data=_read_kvi_excel(path);save_json_file(KVI_POWER_DATA_FILE,{'uploaded_at':datetime.now().strftime('%Y-%m-%d %H:%M:%S'),'uploaded_by':session.get('user_name'),'source_file':f.filename,'items':data});flash(f'{len(data)} articles imported from Excel.','success');return redirect(url_for('kvi_power_page'))
                except Exception as e:msg=f'Excel import failed: {e}'
        else:
            saved,msg=_business_upload(KVI_POWER_DOCUMENTS_FILE,'kvi_power')
            if not msg:flash(f'{len(saved)} KVI + Power SKU document(s) uploaded.','success');return redirect(url_for('kvi_power_page'))
    payload=load_json_file(KVI_POWER_DATA_FILE);items=payload.get('items',[]) if isinstance(payload,dict) else []
    # Upgrade old saved values (for example KVI-ART) in memory and persist canonical categories.
    changed=False
    for item in items:
        if not item.get('id'):
            item['id']=uuid.uuid4().hex;changed=True
        canonical=_normalize_article_type(item.get('article_type'))
        if item.get('article_type')!=canonical:
            item['article_type']=canonical;changed=True
    if changed and isinstance(payload,dict):
        payload['items']=items;save_json_file(KVI_POWER_DATA_FILE,payload)
    total=len(items)
    kvi_only=sum(1 for x in items if x.get('article_type')=='KVI')
    kvi_power=sum(1 for x in items if x.get('article_type')=='KVI+POWER')
    power=sum(1 for x in items if x.get('article_type')=='POWER SKU')
    summary=f'''<div class="grid grid-cols-2 gap-3 lg:grid-cols-4"><button type="button" onclick="filterKviType('ALL')" class="rounded-2xl bg-slate-900 p-5 text-left text-white"><span class="text-xs">Total Articles</span><b class="block text-3xl">{total}</b></button><button type="button" onclick="filterKviType('KVI')" class="rounded-2xl bg-emerald-600 p-5 text-left text-white"><span class="text-xs">KVI</span><b class="block text-3xl">{kvi_only}</b></button><button type="button" onclick="filterKviType('KVI+POWER')" class="rounded-2xl bg-violet-600 p-5 text-left text-white"><span class="text-xs">KVI + Power</span><b class="block text-3xl">{kvi_power}</b></button><button type="button" onclick="filterKviType('POWER SKU')" class="rounded-2xl bg-amber-500 p-5 text-left text-white"><span class="text-xs">Power SKU</span><b class="block text-3xl">{power}</b></button></div>'''
    upload=''
    if session.get('role') in ['admin','developer']:
        upload='''<div class="grid grid-cols-1 gap-3 lg:grid-cols-2"><form method="post" enctype="multipart/form-data" class="rounded-2xl border bg-white p-4"><input type="hidden" name="action" value="excel"><h3 class="mb-2 font-black">Import Article Excel</h3><p class="mb-3 text-xs text-slate-500">Expected columns: Article Type, Article Code, Article Name</p><input type="file" name="excel" accept=".xlsx" required class="w-full rounded-xl border p-2 text-xs"><button class="mt-3 w-full rounded-xl bg-amber-500 py-2 font-black">Upload & Import Excel</button></form><form method="post" enctype="multipart/form-data" class="rounded-2xl border bg-white p-4"><input type="hidden" name="action" value="document"><h3 class="mb-2 font-black">Upload Supporting Document</h3><input name="title" placeholder="Document title" class="mb-2 w-full rounded-xl border px-3 py-2"><input type="file" name="document" multiple required onchange="showSelectedFiles(this,'kvi-support-file-list')" class="w-full rounded-xl border p-2 text-xs"><p id="kvi-support-file-list" class="mt-1 text-[10px] text-slate-500">You can select multiple files at once.</p><button class="mt-3 w-full rounded-xl bg-indigo-600 py-2 font-black text-white">Upload Document</button></form></div>'''
    can_delete=session.get('role') in ['admin','developer']
    row_parts=[]
    for x in items:
        checkbox=''
        if can_delete:
            checkbox='<td><input class="kvi-select" type="checkbox" name="article_ids" value="{}"></td>'.format(x.get('id',''))
        row_parts.append("<tr class='kvi-row border-b' data-type='{}' data-search='{} {} {}'>{}<td>{}</td><td class='font-mono'>{}</td><td>{}</td></tr>".format(x.get('article_type',''),str(x.get('article_type','')).lower(),str(x.get('article_code','')).lower(),str(x.get('article_name','')).lower(),checkbox,x.get('article_type',''),x.get('article_code',''),x.get('article_name','')))
    trs=''.join(row_parts)
    controls='''<div class="grid grid-cols-1 gap-2 rounded-2xl border bg-white p-3 sm:grid-cols-2"><input id="kvi-search" oninput="applyKviFilters()" placeholder="Search article code or name" class="rounded-xl border px-3 py-2 text-sm"><select id="kvi-type-filter" onchange="applyKviFilters()" class="rounded-xl border px-3 py-2 text-sm"><option value="ALL">All Article Types</option><option value="KVI">KVI</option><option value="KVI+POWER">KVI + Power</option><option value="POWER SKU">Power SKU</option><option value="OTHER">Other</option></select><div id="kvi-filter-count" class="sm:col-span-2 text-xs font-bold text-slate-500"></div></div>'''
    delete_toolbar=''
    select_header=''
    empty_colspan=3
    if can_delete:
        empty_colspan=4
        select_header='<th data-no-sort="1"><input type="checkbox" onchange="toggleAllKvi(this)" title="Select all visible rows"></th>'
        delete_toolbar='''<div class="flex flex-wrap gap-2 rounded-2xl border border-rose-200 bg-rose-50 p-3"><button type="button" onclick="deleteSelectedKvi()" class="rounded-xl bg-rose-600 px-4 py-2 text-xs font-black text-white">Delete Selected</button><button type="button" onclick="deleteAllKvi()" class="rounded-xl border border-rose-300 bg-white px-4 py-2 text-xs font-black text-rose-700">Delete All Existing Data</button><span class="self-center text-xs text-rose-700">Deletion is permanent.</span></div>'''
    table=f'''{controls}<form id="kvi-delete-form" method="post" action="/kvi_power_sku/delete">{delete_toolbar}<input type="hidden" id="kvi-delete-action" name="action" value="selected"><div class="overflow-auto rounded-2xl border bg-white"><table id="kvi-article-table" class="wide sortable-table w-full text-sm"><thead class="bg-slate-900 text-white"><tr>{select_header}<th data-sort-type="text">Article Type ↕</th><th data-sort-type="number">Article Code ↕</th><th data-sort-type="text">Article Name ↕</th></tr></thead><tbody>{trs or f'<tr><td colspan="{empty_colspan}" class="p-8 text-center text-slate-400">No Excel data imported yet.</td></tr>'}</tbody></table></div></form>'''
    docs=_doc_cards(_business_docs(KVI_POWER_DOCUMENTS_FILE))
    meta=f"<p class='text-xs text-slate-500'>Last Excel: {payload.get('source_file','-') if isinstance(payload,dict) else '-'} · {payload.get('uploaded_at','-') if isinstance(payload,dict) else '-'}</p>"
    body=(summary+meta+(f'<p class="text-rose-600">{msg}</p>' if msg else '')+upload+'<section><h2 class="mb-3 font-black">Article Data</h2>'+table+'</section><section><h2 class="mb-3 font-black">Supporting Documents</h2>'+docs+'</section>')
    return business_page('KVI + Power SKU','⚡','Live article summary shared with every employee',body)

@app.route('/business_file/<did>')
def business_file(did):
    if not session.get('logged_in'):return redirect(url_for('login'))
    rows=_business_docs(PROMOTION_DOCUMENTS_FILE)+_business_docs(KVI_POWER_DOCUMENTS_FILE);x=next((a for a in rows if a.get('id')==did),None)
    if not x:return 'Document not found',404
    return send_from_directory(BUSINESS_DOC_FOLDER,x['filename'],as_attachment=False,download_name=x.get('original',x['filename']))

# === WORKFORCE AUTOMATION SUITE ===
def suite_load(path, default):
    data=load_json_file(path)
    return data if isinstance(data,type(default)) else default

def suite_save(path, data):
    save_json_file(path,data)

def suite_allowed_stores():
    return [x for x in get_user_stores() if x!='DEV']

def suite_store_allowed(store):
    return session.get('role')=='developer' or store in suite_allowed_stores()

def suite_audit(action, target='', old='', new='', store=''):
    rows=suite_load(AUDIT_LOG_FILE,[])
    rows.append({'id':uuid.uuid4().hex,'time':datetime.now().strftime('%Y-%m-%d %H:%M:%S'),'user_id':session.get('user_id','SYSTEM'),'user_name':session.get('user_name','SYSTEM'),'role':session.get('role','system'),'store':store or session.get('store',''),'action':action,'target':target,'old':str(old),'new':str(new),'ip':request.remote_addr or ''})
    suite_save(AUDIT_LOG_FILE,rows[-5000:])

def suite_notify(user_id, title, message, store=''):
    rows=suite_load(NOTIFICATIONS_FILE,[])
    rows.append({'id':uuid.uuid4().hex,'user_id':user_id,'title':title,'message':message,'store':store,'created_at':datetime.now().strftime('%Y-%m-%d %H:%M:%S'),'read':False})
    suite_save(NOTIFICATIONS_FILE,rows[-5000:])

def suite_employee_store(info):
    return (info.get('store') or (info.get('stores') or [''])[0]).upper()

def suite_leave_balance(uid):
    balances=suite_load(LEAVE_BALANCES_FILE,{})
    base=balances.get(uid,{'annual':18,'used':0,'pending':0})
    base['available']=max(0,int(base.get('annual',18))-int(base.get('used',0))-int(base.get('pending',0)))
    return base

WORKFORCE_HUB_TEMPLATE='''
<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Workforce Automation Hub</title><script src="https://cdn.tailwindcss.com"></script></head>
<body class="bg-slate-100 text-slate-800" data-language="{{ ui_language }}"><header class="bg-slate-950 text-white px-3 sm:px-6 py-4 flex flex-col sm:flex-row gap-3 sm:items-center sm:justify-between"><div><h1 class="font-black text-xl">⚙️ Workforce Automation Hub</h1><p class="text-xs text-slate-300">Multi-store operations, compliance and employee self-service</p></div><div class="flex gap-2 flex-wrap"><select onchange="location.href='/set_language/'+this.value" class="text-slate-900 text-xs p-2 rounded"><option value="en" {% if ui_language=='en' %}selected{% endif %}>English</option><option value="pt" {% if ui_language=='pt' %}selected{% endif %}>Português</option></select><a href="/" class="bg-white/10 px-4 py-2 rounded-xl text-sm font-bold">← Dashboard</a></div></header>
<main class="p-5 space-y-5">
<div class="grid grid-cols-2 md:grid-cols-4 xl:grid-cols-8 gap-3">{% for x in cards %}<div class="bg-white rounded-2xl p-4 border shadow-sm"><div class="text-[10px] uppercase text-slate-500 font-bold">{{x.label}}</div><div class="text-2xl font-black mt-1 {{x.color}}">{{x.value}}</div></div>{% endfor %}</div>
<section class="grid grid-cols-1 lg:grid-cols-2 gap-4">
<div class="bg-white rounded-2xl p-4 border"><div class="flex justify-between mb-3"><h2 class="font-black">🏬 Store Comparison</h2><span class="text-xs text-slate-500">{{today}}</span></div><div class="overflow-auto"><table class="w-full text-xs"><thead class="bg-slate-100"><tr><th class="p-2 text-left">Store</th><th>Machine</th><th>Present</th><th>Absent</th><th>Off</th><th>Late</th></tr></thead><tbody>{% for x in store_rows %}<tr class="border-b"><td class="p-2 font-black">{{x.store}}</td><td class="text-center">{{'Online' if x.online else 'Offline'}}</td><td class="text-center text-emerald-600 font-bold">{{x.present}}</td><td class="text-center text-rose-600 font-bold">{{x.absent}}</td><td class="text-center">{{x.off}}</td><td class="text-center text-amber-600">{{x.late}}</td></tr>{% endfor %}</tbody></table></div></div>
<div class="bg-white rounded-2xl p-4 border"><h2 class="font-black mb-3">🚨 Contract Alerts</h2><div class="max-h-60 overflow-auto space-y-2">{% for x in contract_alerts %}<div class="border rounded-xl p-3 text-xs flex justify-between"><div><b>{{x.name}}</b><br><span class="text-slate-500">{{x.uid}} · {{x.store}}</span></div><div class="text-right"><b>{{x.date}}</b><br><span class="{{x.color}}">{{x.label}}</span></div></div>{% else %}<p class="text-sm text-slate-400">No contract alerts.</p>{% endfor %}</div></div>
</section>
<section class="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-4">
{% for tool in tools %}<a href="{{tool.url}}" class="bg-white border rounded-2xl p-5 hover:border-cyan-400 hover:shadow-md transition"><div class="text-2xl">{{tool.icon}}</div><h3 class="font-black mt-2">{{tool.title}}</h3><p class="text-xs text-slate-500 mt-1">{{tool.desc}}</p></a>{% endfor %}
</section></main>
<script id="universal-table-sort">
(function(){
  function value(cell,type){const text=(cell?.innerText||'').trim();if(type==='number'){const n=parseFloat(text.replace(/[^0-9.-]/g,''));return Number.isNaN(n)?-Infinity:n;}return text.toLocaleLowerCase();}
  function enhance(root=document){root.querySelectorAll('table').forEach((table,ti)=>{if(table.dataset.sortReady)return;table.dataset.sortReady='1';const heads=table.querySelectorAll('thead th');heads.forEach((th,ci)=>{if(th.dataset.noSort==='1')return;th.style.cursor='pointer';th.style.userSelect='none';if(!/[↕↑↓]/.test(th.textContent))th.insertAdjacentText('beforeend',' ↕');th.title='Click to sort';th.addEventListener('click',()=>{const tbody=table.tBodies[0];if(!tbody)return;const rows=Array.from(tbody.rows).filter(r=>r.cells.length>ci&&!r.querySelector('[colspan]'));const asc=th.dataset.direction!=='asc';heads.forEach(h=>{h.dataset.direction='';h.textContent=h.textContent.replace(/ [↑↓]$/,' ↕');});th.dataset.direction=asc?'asc':'desc';th.textContent=th.textContent.replace(/ ↕$/,'')+(asc?' ↑':' ↓');const type=th.dataset.sortType||((rows.every(r=>/^[-+]?\d[\d.,]*$/.test((r.cells[ci]?.innerText||'').trim())))?'number':'text');rows.sort((a,b)=>{const av=value(a.cells[ci],type),bv=value(b.cells[ci],type);return (av>bv?1:av<bv?-1:0)*(asc?1:-1);});rows.forEach(r=>tbody.appendChild(r));});});});}
  window.applyKviFilters=function(){const q=(document.getElementById('kvi-search')?.value||'').trim().toLowerCase();const type=document.getElementById('kvi-type-filter')?.value||'ALL';document.querySelectorAll('.kvi-row').forEach(r=>{let rt=(r.dataset.type||'OTHER').trim().toUpperCase();if(rt.includes('KVI')&&rt.includes('POWER'))rt='KVI+POWER';else if(rt.includes('KVI'))rt='KVI';else if(rt.includes('POWER'))rt='POWER SKU';const typeOk=type==='ALL'||(type==='OTHER'&&!['KVI','KVI+POWER','POWER SKU'].includes(rt))||rt===type;const searchOk=!q||(r.dataset.search||'').includes(q);r.style.display=typeOk&&searchOk?'':'none';});};
  window.filterKviType=function(type){const sel=document.getElementById('kvi-type-filter');if(sel)sel.value=type;window.applyKviFilters();document.getElementById('kvi-article-table')?.scrollIntoView({behavior:'smooth',block:'start'});};
  document.addEventListener('DOMContentLoaded',()=>enhance(document));window.enhanceSortableTables=enhance;
})();
</script>
<script>
window.HRMS_LANGUAGE={{ ui_language|tojson }};
window.HRMS_PT={'Correction Rejected':'Correcção Rejeitada','Correction Approved':'Correcção Aprovada','Correction Request':'Pedido de Correcção','New Document':'Novo Documento','was Rejected':'foi rejeitada','was Approved':'foi aprovada','Your correction':'A sua correcção','Your roster':'A sua escala','Month select karke Load Month click karein':'Seleccione o mês e clique em Carregar Mês','Employee-wise Shift A, Shift B and Weekly Off allocation':'Alocação por trabalhador de Turno A, Turno B e Folga Semanal','Public Holiday':'Feriado Público','Half Day':'Meio Dia','Full Day':'Dia Completo','Full Month':'Mês Completo','Clear All':'Limpar Tudo','Select All':'Seleccionar Tudo','Uploaded By':'Carregado por','Updated By':'Actualizado por','Created On':'Criado em','Created By':'Criado por','Records':'Registos','Record':'Registo','Summary':'Resumo','Details':'Detalhes','Code':'Código','Address':'Endereço','Phone':'Telefone','Email':'E-mail','Photo':'Fotografia','Store Access':'Acesso à Loja','Apply Selected':'Aplicar Seleccionados','Copy Previous Month':'Copiar Mês Anterior','Copy Previous Week':'Copiar Semana Anterior','Monthly roster saved':'Escala mensal guardada','Kam se kam ek option select karein.':'Seleccione pelo menos uma opção.','Month select karein':'Seleccione o mês','Galat User ID ya Password!':'ID do utilizador ou palavra-passe incorrectos!','Sirf PDF files allowed hain!':'Apenas ficheiros PDF são permitidos!','Sabhi fields bharna zaroori hai!':'Todos os campos são obrigatórios!','This roster shows only your shifts and weekly-off dates for the current month.':'Esta escala mostra apenas os seus turnos e folgas semanais do mês actual.','Overtime Minutes':'Minutos de Horas Extra','Submit Overtime':'Enviar Horas Extra','Email - requires SMTP':'E-mail - requer SMTP','Internal Notification':'Notificação Interna','Download Center':'Centro de Transferências','Monthly':'Mensal','Weekly':'Semanal','Daily':'Diário','Delivery':'Entrega','Frequency':'Frequência','File':'Ficheiro','Uploaded At':'Carregado em','Created':'Criado','To':'Para','From':'De','Effective':'Efectivo','Estimate':'Estimativa','Employer':'Empregador','Gross':'Bruto','Configurable estimate using 2026 settings. Validate payroll categories, taxable base and exemptions before finalisation.':'Estimativa configurável com definições de 2026. Valide as categorias salariais, a base tributável e as isenções antes da finalização.','Identification masking':'Ocultação da Identificação','Document access log':'Registo de Acesso a Documentos','Access audit':'Auditoria de Acesso','Last backups':'Últimas Cópias de Segurança','Number of Users':'Número de Utilizadores','Application Version':'Versão da Aplicação','retry required':'nova tentativa necessária','Last saved attendance remains available':'A última assiduidade guardada continua disponível','Queue State':'Estado da Fila','Level':'Nível','Late Count':'Número de Atrasos','Rolling 30-day view.':'Vista móvel de 30 dias.','roster cells copied':'células de escala copiadas','assignments saved':'alocações guardadas','Copy failed':'Falha ao copiar','Save Lifecycle Event':'Guardar Evento do Ciclo de Vida','Employee Lifecycle':'Ciclo de Vida do Trabalhador','New Records':'Novos Registos','Records Received':'Registos Recebidos','Machine Status':'Estado do Equipamento','Store Code':'Código da Loja','Payroll, holidays, lifecycle, overtime, backups and privacy controls.':'Salários, feriados, ciclo de vida, horas extra, cópias de segurança e controlos de privacidade.','Privacy-safe QR verification cards.':'Cartões QR de verificação com protecção de privacidade.','Internal leave, shift, payroll and contract notices.':'Avisos internos de ausências, turnos, salários e contratos.','Track user, store, action, IP and changes.':'Acompanhar utilizador, loja, acção, IP e alterações.','Minimum staffing rules and shortage warnings.':'Regras de dotação mínima e avisos de falta de pessoal.','Copy week or month schedules.':'Copiar escalas semanais ou mensais.','Contracts, IDs and certificates.':'Contratos, documentos de identificação e certificados.','Annual, used, pending and available balances.':'Saldos anuais, usados, pendentes e disponíveis.','Approve missing-punch correction requests.':'Aprovar pedidos de correcção de marcações em falta.','Machine status, manual checks and sync history.':'Estado dos equipamentos, verificações manuais e histórico de sincronização.','Combined attendance with store code and Excel export.':'Assiduidade combinada com código da loja e exportação Excel.','Multi-store operations, compliance and employee self-service':'Operações multi-loja, conformidade e auto-serviço do trabalhador','No contract alerts.':'Sem alertas de contrato.','Audit Events':'Eventos de Auditoria','Contracts':'Contratos','Corrections':'Correcções','Machine':'Equipamento','Store Comparison':'Comparação de Lojas','Monthly Summary':'Resumo Mensal','Default Schedule':'Horário Padrão','All Allocations':'Todas as Alocações','Designations Selected':'Funções Seleccionadas','All Designations Selected':'Todas as Funções Seleccionadas','Same Weekday for Full Month':'Mesmo Dia da Semana durante Todo o Mês','Monday-Sunday':'Segunda a Domingo','Last Week':'Última Semana','5th Week':'5.ª Semana','4th Week':'4.ª Semana','3rd Week':'3.ª Semana','2nd Week':'2.ª Semana','1st Week':'1.ª Semana','All Weeks (Monday-Sunday)':'Todas as Semanas (Segunda a Domingo)','Not in this month':'Não existe neste mês','options selected':'opções seleccionadas','option selected':'opção seleccionada','Apply to selected date':'Aplicar à data seleccionada','Month & Year':'Mês e Ano','load month':'carregar mês','Select Month':'Seleccionar Mês','No shift requests':'Sem pedidos de turno','Current Assignment':'Alocação Actual','Requested':'Solicitado','Shift Request':'Pedido de Turno','Request Shift / Off':'Pedir Turno / Folga','Select Employee':'Seleccionar Trabalhador','Bulk Upload':'Carregamento em Massa','Standard Row':'Linha Padrão','Monthly grid':'Grelha mensal','Master':'Principal','Individual Slip':'Recibo Individual','Split':'Separar','All Personnel':'Todo o Pessoal','No requests':'Sem pedidos','Minutes':'Minutos','ALL':'TODOS','Change':'Alteração','Emp':'Trab.','Remove Photo':'Remover Fotografia','STORE ACCESS':'ACESSO À LOJA','Create login and employee profile':'Criar início de sessão e perfil do trabalhador','No employees found.':'Nenhum trabalhador encontrado.','Select Designation':'Seleccionar Função','Add Employee':'Adicionar Trabalhador','Search, review and maintain workforce profiles':'Pesquisar, rever e manter perfis dos trabalhadores','Open Roster Planner':'Abrir Planeador de Escala','Use Roster Planner to assign shifts and weekly-off rotations.':'Utilize o Planeador de Escala para atribuir turnos e rotações de folga semanal.','Rota Rotation and Shift Schedules are integrated with the':'A rotação da escala e os horários de turno estão integrados com o','No salary slips uploaded yet.':'Ainda não foram carregados recibos de salário.','View':'Ver','PDF File':'Ficheiro PDF','-- Select Employee --':'-- Seleccionar Trabalhador --','Individual Slip Upload':'Carregar Recibo Individual','Split & Upload':'Separar e Carregar','Master PDF':'PDF Principal','Bulk Upload (Merged PDF)':'Carregamento em Massa (PDF Unido)','Monthly grid with Leave Codes':'Grelha mensal com códigos de ausência','Employee Matrix':'Matriz de Trabalhadores','Individual records per date':'Registos individuais por data','Standard Row Export':'Exportação em Linhas','Download Attendance Reports':'Baixar Relatórios de Assiduidade','Total Summary':'Resumo Total','No attendance records found for this selection.':'Nenhum registo de assiduidade encontrado para esta selecção.','Status ↕':'Estado ↕','Working Hrs ↕':'Horas Trabalhadas ↕','Total Lunch ↕':'Total de Almoço ↕','Out Time ↕':'Hora de Saída ↕','Store In ↕':'Entrada na Loja ↕','Dept ↕':'Departamento ↕','Employee Name ↕':'Nome do Trabalhador ↕','Date ↕':'Data ↕','Last Month':'Mês Passado','This Month':'Este Mês','This Week':'Esta Semana','Yesterday':'Ontem','Quick Range':'Intervalo Rápido','Rota':'Escala','Export':'Exportar','Logged In As':'Sessão Iniciada Como','-- All Personnel --':'-- Todo o Pessoal --','Employee Filter':'Filtro de Trabalhadores','Tot Hrs':'Total de Horas','Mis Punch':'Marcação em Falta','Mis-Punch':'Marcação em Falta','Late Arr.':'Chegada Tardia','Week Off':'Folga Semanal','Shutdown':'Encerrar','Sync':'Sincronização','Device':'Equipamento','Leaves':'Ausências','Dev':'Programador','Biometric live tracking active for Attendance Portal.':'Acompanhamento biométrico em tempo real activo no Portal de Assiduidade.','Announcements':'Comunicados','Employees Info (ID Card)':'Informações dos Trabalhadores (Cartão de Identificação)','Manage Passwords':'Gerir Palavras-passe','Biometric Machines':'Equipamentos Biométricos','Shift Approvals':'Aprovações de Turno','Reset':'Repor','Off':'Folga','Daily manpower summary':'Resumo diário de efectivos','Loading monthly roster...':'A carregar a escala mensal...','Send Request for Approval':'Enviar Pedido para Aprovação','Developed by':'Desenvolvido por','Attendance Portal':'Portal de Assiduidade','Sign in to access your dashboard':'Inicie sessão para aceder ao seu painel','User ID':'ID do Utilizador','Password':'Palavra-passe','Secure Login':'Iniciar Sessão','Forgot/Reset Password?':'Esqueceu/Redefinir Palavra-passe?','Reset Password':'Redefinir Palavra-passe','Create a new password request':'Criar um novo pedido de palavra-passe','Submit Reset Request':'Enviar Pedido de Redefinição','Back to Login':'Voltar ao Início de Sessão','Current Password':'Palavra-passe Actual','New Password':'Nova Palavra-passe','Confirm Password':'Confirmar Palavra-passe','Dashboard':'Painel','Good day':'Bom dia','Main Menu':'Menu Principal','Team Management':'Gestão da Equipa','Quick Actions':'Acções Rápidas','Search':'Pesquisar','Search code or employee name':'Pesquisar código ou nome do trabalhador','Filter':'Filtrar','Reset Filters':'Limpar Filtros','Close':'Fechar','Save':'Guardar','Cancel':'Cancelar','Apply':'Aplicar','Load':'Carregar','Load Month':'Carregar Mês','Export Excel':'Exportar Excel','English':'Inglês','Portuguese':'Português','Language':'Idioma','Employee':'Trabalhador','Employees':'Trabalhadores','Employee Code':'Código do Trabalhador','Employee Name':'Nome do Trabalhador','Employee ID':'ID do Trabalhador','Designation':'Função','Department':'Departamento','Store':'Loja','Stores':'Lojas','All Stores':'Todas as Lojas','All Designations':'Todas as Funções','All Allocations':'Todas as Alocações','All Weeks':'Todas as Semanas','Selected':'Seleccionado','Selected Date Only':'Apenas a Data Seleccionada','Same Weekday for Full Month':'Mesmo Dia da Semana em Todo o Mês','First Week':'Primeira Semana','Second Week':'Segunda Semana','Third Week':'Terceira Semana','Fourth Week':'Quarta Semana','Fifth Week':'Quinta Semana','Last Week':'Última Semana','Week calculation':'Cálculo da semana','Monday is the first day and Sunday is the last day':'Segunda-feira é o primeiro dia e Domingo é o último dia','Attendance':'Assiduidade','Attendance Correction':'Correcção de Assiduidade','Attendance Corrections':'Correcções de Assiduidade','Present':'Presente','Absent':'Ausente','Late':'Atrasado','Late Arrival':'Chegada Tardia','Early Departure':'Saída Antecipada','Weekly Off':'Folga Semanal','Shift A':'Turno A','Shift B':'Turno B','Default':'Padrão','Roster Planner':'Planeador de Escala','Monthly Roster':'Escala Mensal','Monthly Roster Matrix':'Matriz de Escala Mensal','Save Monthly Roster':'Guardar Escala Mensal','Daily manpower summary':'Resumo diário de efectivos','Shift Allocation':'Alocação de Turno','Date & Day':'Data e Dia','My Monthly Roster':'Minha Escala Mensal','Shift / Weekly Off':'Turno / Folga Semanal','Today':'Hoje','Scheduled':'Programado','Date':'Data','Day':'Dia','Month':'Mês','Year':'Ano','Hours':'Horas','Minutes':'Minutos','Time':'Hora','Store In':'Entrada na Loja','Lunch Out':'Saída para Almoço','Lunch In':'Regresso do Almoço','Out Time':'Hora de Saída','Working Hours':'Horas Trabalhadas','Total Hours':'Total de Horas','Lunch Hours':'Horas de Almoço','Variance':'Variação','Device':'Equipamento','Online':'Online','Offline':'Offline','Loading':'A carregar','No data':'Sem dados','No records found':'Nenhum registo encontrado','Leave Portal':'Portal de Ausências','Leave Management':'Gestão de Ausências','Leave Balances':'Saldos de Ausências','Leave Type':'Tipo de Ausência','Leave Code':'Código de Ausência','Submit Leave':'Enviar Pedido','Start Date':'Data Inicial','End Date':'Data Final','Supporting Document':'Documento Comprovativo','Status':'Estado','Action':'Acção','Pending':'Pendente','Approved':'Aprovado','Rejected':'Rejeitado','Processed':'Processado','Approve':'Aprovar','Reject':'Rejeitar','Reason':'Motivo','Request':'Pedido','Requests':'Pedidos','No leave requests':'Sem pedidos de ausência','Select Leave Code / Motivo':'Seleccionar Código de Ausência / Motivo','Payroll & Reports':'Salários e Relatórios','Payroll':'Salários','Salary':'Salário','Salary Slip':'Recibo de Salário','Gross Salary':'Salário Bruto','Net Salary':'Salário Líquido','Base Salary':'Salário Base','Meal Allowance':'Subsídio de Alimentação','Transport Allowance':'Subsídio de Transporte','Bonus':'Prémio','Overtime':'Horas Extra','Overtime Approval':'Aprovação de Horas Extra','Net estimate':'Estimativa Líquida','Calculate Estimate':'Calcular Estimativa','Reports':'Relatórios','Report':'Relatório','Upload Salary Slip':'Carregar Recibo de Salário','Download Salary Slip':'Baixar Recibo de Salário','Employee Documents':'Documentos do Trabalhador','My Documents':'Meus Documentos','Notifications':'Notificações','Notification Center':'Centro de Notificações','ID Card':'Cartão de Identificação','Download ID Card':'Baixar Cartão de Identificação','Baixar / Download ID Card':'Baixar Cartão de Identificação','Calendar & Rota':'Calendário e Escala','Employee List':'Lista de Trabalhadores','Employee Cards':'Cartões dos Trabalhadores','Employee Information':'Informações do Trabalhador','Profile':'Perfil','My Profile':'Meu Perfil','Nationality':'Nacionalidade','HRMS Code':'Código HRMS','Identification':'Identificação','Contract Date':'Data de Contrato','Data de Contrato':'Data de Contrato','Create / Update':'Criar / Actualizar','Create Employee':'Criar Trabalhador','Update Employee':'Actualizar Trabalhador','Manage':'Gerir','Edit':'Editar','Delete':'Eliminar','Open':'Abrir','Download':'Baixar','Upload':'Carregar','Document':'Documento','Documents':'Documentos','No documents':'Sem documentos','Settings':'Definições','Logout':'Sair','User Management':'Gestão de Utilizadores','Password Management':'Gestão de Palavras-passe','Machine Management':'Gestão de Equipamentos','Permissions':'Permissões','Role':'Perfil de Acesso','Admin':'Administrador','Developer':'Programador','Blocked':'Bloqueado','Active':'Activo','Inactive':'Inactivo','Create User':'Criar Utilizador','Update User':'Actualizar Utilizador','Store View':'Vista da Loja','Developer View':'Vista do Programador','All Stores View':'Vista de Todas as Lojas','Workforce Automation Hub':'Centro de Automação da Força de Trabalho','Multi-store operations, compliance and employee self-service':'Operações multi-loja, conformidade e auto-serviço do trabalhador','Contract Alerts':'Alertas de Contrato','All Stores Attendance':'Assiduidade de Todas as Lojas','Biometric Sync History':'Histórico de Sincronização Biométrica','Manual Machine Check':'Verificação Manual do Equipamento','Check Now':'Verificar Agora','Last Sync':'Última Sincronização','Roster Copy':'Copiar Escala','Copy Week':'Copiar Semana','Copy Month':'Copiar Mês','Source Date':'Data de Origem','Target Date':'Data de Destino','Staffing Rules':'Regras de Dotação','Minimum Staffing Rules':'Regras de Dotação Mínima','Audit Log':'Registo de Auditoria','Audit Events':'Eventos de Auditoria','Employee QR Directory':'Directório QR dos Trabalhadores','Holiday Calendar':'Calendário de Feriados','National Holiday':'Feriado Nacional','Company Holiday':'Feriado da Empresa','Bridge Day':'Ponte','Paid':'Remunerado','Unpaid':'Não Remunerado','Late Escalation':'Escalonamento de Atrasos','Late Arrival Escalation':'Escalonamento de Chegadas Tardias','Onboarding Checklist':'Lista de Integração','Offboarding':'Desvinculação','Store Transfer':'Transferência de Loja','Employee Timeline':'Linha do Tempo do Trabalhador','Performance & Training':'Desempenho e Formação','Scheduled Reports':'Relatórios Programados','Offline Queue':'Fila Offline','Offline Biometric Queue':'Fila Biométrica Offline','Backup & Restore':'Cópia de Segurança e Restauro','Create Full Backup':'Criar Cópia de Segurança Completa','Backup History':'Histórico de Cópias de Segurança','System Health':'Estado do Sistema','Application Status':'Estado da Aplicação','Disk Used':'Disco Utilizado','Disk Free':'Disco Livre','Privacy Controls':'Controlos de Privacidade','Data Protection':'Protecção de Dados','Privacy':'Privacidade','Open module':'Abrir módulo','Back':'Voltar','Back to Dashboard':'Voltar ao Painel','Type':'Tipo','Name':'Nome','Title':'Título','Available':'Disponível','Used':'Usado','Annual':'Anual','Total Employees':'Total de Trabalhadores','No notifications':'Sem notificações','Manual Sync':'Sincronização Manual','Training':'Formação','Training Name':'Nome da Formação','Training Date':'Data da Formação','Expiry':'Validade','Expiry Date':'Data de Validade','Remarks':'Observações','Manager Remarks':'Observações do Gestor','Score':'Pontuação','Privacy Notice':'Aviso de Privacidade','Data Subject Rights':'Direitos do Titular dos Dados','Purpose Limitation':'Limitação da Finalidade','Retention':'Conservação','Security':'Segurança','Confidential':'Confidencial','Confidential Document':'Documento Confidencial','Submit':'Enviar','Submit Request':'Enviar Pedido','Save Rule':'Guardar Regra','Save Holiday':'Guardar Feriado','Save Schedule':'Guardar Programação','Apply Filters':'Aplicar Filtros','Clear':'Limpar','Select':'Seleccionar','Choose':'Escolher','Required':'Obrigatório','Optional':'Opcional','Success':'Sucesso','Error':'Erro','Warning':'Aviso','Attention':'Atenção','Critical':'Crítico','Normal':'Normal','System':'Sistema','Application':'Aplicação','User':'Utilizador','Created At':'Criado em','Updated At':'Actualizado em','Effective Date':'Data de Efeito','From Store':'Loja de Origem','To Store':'Loja de Destino','Exit Reason':'Motivo de Saída','Final Settlement':'Acerto Final','Documents Returned':'Documentos Devolvidos','Biometric Registration':'Registo Biométrico','Employee Photo':'Fotografia do Trabalhador','Contract Document':'Documento do Contrato','Store Assignment':'Atribuição de Loja','Default Shift':'Turno Padrão','ID Card Generated':'Cartão de Identificação Gerado','Apply Roster':'Aplicar Escala','All Weeks (Monday-Sunday)':'Todas as Semanas (Segunda a Domingo)','Monday':'Segunda-feira','Tuesday':'Terça-feira','Wednesday':'Quarta-feira','Thursday':'Quinta-feira','Friday':'Sexta-feira','Saturday':'Sábado','Sunday':'Domingo'};
function hrmsTranslate(root=document.body){
 if(window.HRMS_LANGUAGE!=='pt'||!root)return;
 const keys=Object.keys(window.HRMS_PT).sort((a,b)=>b.length-a.length);
 const cv=(value)=>{let t=value||'';keys.forEach(k=>{t=t.split(k).join(window.HRMS_PT[k])});return t};
 const walker=document.createTreeWalker(root,NodeFilter.SHOW_TEXT);const nodes=[];while(walker.nextNode())nodes.push(walker.currentNode);
 nodes.forEach(n=>{if(!n.parentElement||['SCRIPT','STYLE','TEXTAREA'].includes(n.parentElement.tagName))return;n.nodeValue=cv(n.nodeValue)});
 root.querySelectorAll('[placeholder],[title],[aria-label],input[type="button"],input[type="submit"]').forEach(el=>{
   ['placeholder','title','aria-label','value'].forEach(a=>{const v=el.getAttribute(a);if(v)el.setAttribute(a,cv(v))});
 });
 document.documentElement.lang='pt';document.title=cv(document.title);
}
document.addEventListener('DOMContentLoaded',()=>hrmsTranslate());
new MutationObserver(ms=>{if(window.HRMS_LANGUAGE==='pt')ms.forEach(m=>m.addedNodes.forEach(n=>{if(n.nodeType===1)hrmsTranslate(n)}))}).observe(document.documentElement,{childList:true,subtree:true});
</script>

<style id="hrms-toast-style">
#hrms-toast-host{position:fixed;top:18px;right:18px;z-index:99999;width:min(420px,calc(100vw - 36px));display:flex;flex-direction:column;gap:10px;pointer-events:none}
.hrms-toast{pointer-events:auto;display:flex;align-items:flex-start;gap:12px;padding:14px 16px;border-radius:14px;color:#fff;font:700 12px/1.45 Inter,system-ui,sans-serif;box-shadow:0 18px 45px rgba(15,23,42,.24);transform:translateX(0);opacity:1;transition:opacity .3s ease,transform .3s ease;overflow-wrap:anywhere}
.hrms-toast-success{background:#059669}.hrms-toast-danger,.hrms-toast-error{background:#e11d48}.hrms-toast-warning{background:#d97706}.hrms-toast-info{background:#2563eb}
.hrms-toast-hide{opacity:0;transform:translateX(30px)}.hrms-toast-message{flex:1}.hrms-toast-close{border:0;background:transparent;color:#fff;font-size:21px;line-height:1;cursor:pointer;min-height:0;padding:0 0 0 6px}
@media(max-width:640px){ #hrms-toast-host{top:10px;right:10px;width:calc(100vw - 20px)}}
</style>
<script id="hrms-toast-script">
(function(){
  function host(){let h=document.getElementById('hrms-toast-host');if(!h){h=document.createElement('div');h.id='hrms-toast-host';h.setAttribute('aria-live','polite');h.setAttribute('aria-atomic','false');document.body.appendChild(h)}return h}
  window.showToast=function(message,type='success'){
    if(message===undefined||message===null||String(message).trim()==='')return;
    const normalized=(type==='error'?'danger':type)||'info';
    const t=document.createElement('div');t.className='hrms-toast hrms-toast-'+normalized;t.setAttribute('role',normalized==='danger'?'alert':'status');
    const m=document.createElement('div');m.className='hrms-toast-message';m.textContent=String(message);
    const x=document.createElement('button');x.type='button';x.className='hrms-toast-close';x.innerHTML='&times;';x.setAttribute('aria-label','Close notification');
    const remove=()=>{if(t.dataset.closing)return;t.dataset.closing='1';t.classList.add('hrms-toast-hide');setTimeout(()=>t.remove(),300)};
    x.addEventListener('click',remove);t.append(m,x);host().appendChild(t);setTimeout(remove,5000);
  };
  document.addEventListener('DOMContentLoaded',()=>document.querySelectorAll('.hrms-flash').forEach(el=>{showToast(el.dataset.message,el.dataset.category);el.remove()}));
})();
</script>
</body></html>'''

@app.route('/workforce_hub')
def workforce_hub():
    if not session_has_permission('workforce_hub'): return redirect(url_for('index'))
    if session.get('role') not in ['admin','developer']: return redirect(url_for('index'))
    stores=suite_allowed_stores(); today=datetime.now().strftime('%Y-%m-%d'); store_rows=[]
    totals={'present':0,'absent':0,'off':0,'late':0}
    for st in stores:
        try:
            _,_,_,_,_,_,stats=fetch_attendance_data(today,today,'ALL',st)
            row={'store':st,'online':bool(stats.get('device_online')),'present':stats.get('present',0),'absent':stats.get('absent',0),'off':stats.get('off',0),'late':stats.get('late_arrival',0)}
        except Exception:
            row={'store':st,'online':False,'present':0,'absent':0,'off':0,'late':0}
        store_rows.append(row)
        for k in totals: totals[k]+=row[k]
    db=load_users_db(); alerts=[]; now=datetime.now().date()
    for uid,raw in db.items():
        info=normalize_user_record(raw); st=suite_employee_store(info)
        if info.get('role')!='employee' or st not in stores: continue
        ds=info.get('data_de_contrato','')
        try:
            d=datetime.strptime(ds,'%m/%d/%y').date() if '/' in ds else datetime.strptime(ds,'%Y-%m-%d').date()
            delta=(d-now).days
            if delta<=90:
                label='Expired' if delta<0 else f'{delta} days remaining'; color='text-rose-600' if delta<=30 else 'text-amber-600'
                alerts.append({'uid':uid,'name':info.get('name',uid),'store':st,'date':ds,'label':label,'color':color})
        except Exception: pass
    cards=[{'label':'Stores','value':len(stores),'color':'text-cyan-700'},{'label':'Present','value':totals['present'],'color':'text-emerald-600'},{'label':'Absent','value':totals['absent'],'color':'text-rose-600'},{'label':'Weekly Off','value':totals['off'],'color':'text-slate-700'},{'label':'Late','value':totals['late'],'color':'text-amber-600'},{'label':'Corrections','value':sum(1 for x in suite_load(CORRECTIONS_FILE,[]) if x.get('status')=='Pending'),'color':'text-indigo-600'},{'label':'Contracts','value':len(alerts),'color':'text-orange-600'},{'label':'Audit Events','value':len(suite_load(AUDIT_LOG_FILE,[])),'color':'text-purple-600'}]
    tools=[
      {'url':'/suite/all_stores','icon':'🏬','title':'All Stores Attendance','desc':'Combined attendance with store code and Excel export.'},
      {'url':'/suite/sync','icon':'🔄','title':'Biometric Sync History','desc':'Machine status, manual checks and sync history.'},
      {'url':'/suite/corrections','icon':'✍️','title':'Attendance Corrections','desc':'Approve missing-punch correction requests.'},
      {'url':'/suite/leave_balances','icon':'🏖️','title':'Leave Balances','desc':'Annual, used, pending and available balances.'},
      {'url':'/suite/documents','icon':'📁','title':'Employee Documents','desc':'Contracts, IDs and certificates.'},
      {'url':'/suite/roster_copy','icon':'📋','title':'Roster Copy','desc':'Copy week or month schedules.'},
      {'url':'/suite/staffing','icon':'👥','title':'Staffing Rules','desc':'Minimum staffing rules and shortage warnings.'},
      {'url':'/suite/audit','icon':'🛡️','title':'Audit Log','desc':'Track user, store, action, IP and changes.'},
      {'url':'/suite/notifications','icon':'🔔','title':'Notification Center','desc':'Internal leave, shift, payroll and contract notices.'},
      {'url':'/suite/qr_directory','icon':'▦','title':'Employee QR Directory','desc':'Privacy-safe QR verification cards.'}
    ]
    return render_template_string(WORKFORCE_HUB_TEMPLATE,cards=cards,store_rows=store_rows,contract_alerts=alerts,tools=tools,today=today,ui_language=session.get('ui_language','en'))

SUITE_TABLE='''<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover"><title>{{title}}</title><script src="https://cdn.tailwindcss.com"></script><style>html,body{max-width:100%;overflow-x:hidden}main{overflow-x:auto;-webkit-overflow-scrolling:touch}input,select,textarea,button{max-width:100%;min-height:42px}table{min-width:640px}th,td{padding:8px;white-space:nowrap}@media(max-width:640px){header{flex-direction:column;align-items:stretch!important;gap:10px}header>div{display:flex;flex-wrap:wrap}main{padding:10px!important}form{grid-template-columns:1fr!important}form button{width:100%}.grid{gap:8px}.rounded-2xl{border-radius:14px}}</style></head><body class="bg-slate-100"><header class="bg-slate-950 text-white p-3 sm:p-4 flex flex-col sm:flex-row gap-2 sm:justify-between sm:items-center"><h1 class="font-black">{{icon}} {{title}}</h1><div class="flex gap-2 items-center"><select onchange="location.href='/set_language/'+this.value" class="text-slate-900 text-xs p-2 rounded"><option value="en" {% if ui_language=='en' %}selected{% endif %}>English</option><option value="pt" {% if ui_language=='pt' %}selected{% endif %}>Português</option></select><a href="/workforce_hub" class="font-bold">← Hub</a></div></header><main class="p-3 sm:p-5 pb-[max(1rem,env(safe-area-inset-bottom))]">{{body|safe}}</main><script id="universal-table-sort">
(function(){
  function value(cell,type){const text=(cell?.innerText||'').trim();if(type==='number'){const n=parseFloat(text.replace(/[^0-9.-]/g,''));return Number.isNaN(n)?-Infinity:n;}return text.toLocaleLowerCase();}
  function enhance(root=document){root.querySelectorAll('table').forEach((table,ti)=>{if(table.dataset.sortReady)return;table.dataset.sortReady='1';const heads=table.querySelectorAll('thead th');heads.forEach((th,ci)=>{if(th.dataset.noSort==='1')return;th.style.cursor='pointer';th.style.userSelect='none';if(!/[↕↑↓]/.test(th.textContent))th.insertAdjacentText('beforeend',' ↕');th.title='Click to sort';th.addEventListener('click',()=>{const tbody=table.tBodies[0];if(!tbody)return;const rows=Array.from(tbody.rows).filter(r=>r.cells.length>ci&&!r.querySelector('[colspan]'));const asc=th.dataset.direction!=='asc';heads.forEach(h=>{h.dataset.direction='';h.textContent=h.textContent.replace(/ [↑↓]$/,' ↕');});th.dataset.direction=asc?'asc':'desc';th.textContent=th.textContent.replace(/ ↕$/,'')+(asc?' ↑':' ↓');const type=th.dataset.sortType||((rows.every(r=>/^[-+]?\d[\d.,]*$/.test((r.cells[ci]?.innerText||'').trim())))?'number':'text');rows.sort((a,b)=>{const av=value(a.cells[ci],type),bv=value(b.cells[ci],type);return (av>bv?1:av<bv?-1:0)*(asc?1:-1);});rows.forEach(r=>tbody.appendChild(r));});});});}
  window.applyKviFilters=function(){const q=(document.getElementById('kvi-search')?.value||'').trim().toLowerCase();const type=document.getElementById('kvi-type-filter')?.value||'ALL';document.querySelectorAll('.kvi-row').forEach(r=>{let rt=(r.dataset.type||'OTHER').trim().toUpperCase();if(rt.includes('KVI')&&rt.includes('POWER'))rt='KVI+POWER';else if(rt.includes('KVI'))rt='KVI';else if(rt.includes('POWER'))rt='POWER SKU';const typeOk=type==='ALL'||(type==='OTHER'&&!['KVI','KVI+POWER','POWER SKU'].includes(rt))||rt===type;const searchOk=!q||(r.dataset.search||'').includes(q);r.style.display=typeOk&&searchOk?'':'none';});};
  window.filterKviType=function(type){const sel=document.getElementById('kvi-type-filter');if(sel)sel.value=type;window.applyKviFilters();document.getElementById('kvi-article-table')?.scrollIntoView({behavior:'smooth',block:'start'});};
  document.addEventListener('DOMContentLoaded',()=>enhance(document));window.enhanceSortableTables=enhance;
})();
</script>
<style id="hrms-toast-style">
#hrms-toast-host{position:fixed;top:18px;right:18px;z-index:99999;width:min(420px,calc(100vw - 36px));display:flex;flex-direction:column;gap:10px;pointer-events:none}
.hrms-toast{pointer-events:auto;display:flex;align-items:flex-start;gap:12px;padding:14px 16px;border-radius:14px;color:#fff;font:700 12px/1.45 Inter,system-ui,sans-serif;box-shadow:0 18px 45px rgba(15,23,42,.24);transform:translateX(0);opacity:1;transition:opacity .3s ease,transform .3s ease;overflow-wrap:anywhere}
.hrms-toast-success{background:#059669}.hrms-toast-danger,.hrms-toast-error{background:#e11d48}.hrms-toast-warning{background:#d97706}.hrms-toast-info{background:#2563eb}
.hrms-toast-hide{opacity:0;transform:translateX(30px)}.hrms-toast-message{flex:1}.hrms-toast-close{border:0;background:transparent;color:#fff;font-size:21px;line-height:1;cursor:pointer;min-height:0;padding:0 0 0 6px}
@media(max-width:640px){ #hrms-toast-host{top:10px;right:10px;width:calc(100vw - 20px)}}
</style>
<script id="hrms-toast-script">
(function(){
  function host(){let h=document.getElementById('hrms-toast-host');if(!h){h=document.createElement('div');h.id='hrms-toast-host';h.setAttribute('aria-live','polite');h.setAttribute('aria-atomic','false');document.body.appendChild(h)}return h}
  window.showToast=function(message,type='success'){
    if(message===undefined||message===null||String(message).trim()==='')return;
    const normalized=(type==='error'?'danger':type)||'info';
    const t=document.createElement('div');t.className='hrms-toast hrms-toast-'+normalized;t.setAttribute('role',normalized==='danger'?'alert':'status');
    const m=document.createElement('div');m.className='hrms-toast-message';m.textContent=String(message);
    const x=document.createElement('button');x.type='button';x.className='hrms-toast-close';x.innerHTML='&times;';x.setAttribute('aria-label','Close notification');
    const remove=()=>{if(t.dataset.closing)return;t.dataset.closing='1';t.classList.add('hrms-toast-hide');setTimeout(()=>t.remove(),300)};
    x.addEventListener('click',remove);t.append(m,x);host().appendChild(t);setTimeout(remove,5000);
  };
  document.addEventListener('DOMContentLoaded',()=>document.querySelectorAll('.hrms-flash').forEach(el=>{showToast(el.dataset.message,el.dataset.category);el.remove()}));
})();
</script>
</body></html>'''


def tr_ui(text):
    if session.get('ui_language','en')!='pt': return text
    mapping={
      'All Stores Attendance':'Assiduidade de Todas as Lojas','Biometric Sync History':'Histórico de Sincronização Biométrica',
      'Attendance Corrections':'Correcções de Assiduidade','Leave Balances':'Saldos de Ausências','Employee Documents':'Documentos do Trabalhador',
      'Roster Copy':'Copiar Escala','Minimum Staffing Rules':'Regras de Dotação Mínima','Audit Log':'Registo de Auditoria',
      'Notification Center':'Centro de Notificações','Employee QR Directory':'Directório QR dos Trabalhadores',
      'Holiday Calendar':'Calendário de Feriados',
      'Bulk Roster Assignment':'Atribuição de Escala em Massa',
      'Overtime Approval':'Aprovação de Horas Extra','Employee Lifecycle':'Ciclo de Vida do Trabalhador','Employee Timeline':'Linha do Tempo do Trabalhador',
      'Performance & Training':'Desempenho e Formação','Scheduled Reports':'Relatórios Programados','Late Arrival Escalation':'Escalonamento de Chegadas Tardias',
      'Offline Biometric Queue':'Fila Biométrica Offline','Backup & Restore':'Cópia de Segurança e Restauro','System Health':'Estado do Sistema','Privacy Controls':'Controlos de Privacidade'
    }
    return mapping.get(text,text)

def suite_page(title,icon,body): return render_template_string(SUITE_TABLE,title=tr_ui(title),icon=icon,body=body,ui_language=session.get('ui_language','en'))

@app.route('/suite/all_stores')
def suite_all_stores():
    if not session_has_permission('all_stores_attendance'): return redirect(url_for('index'))
    if session.get('role') not in ['admin','developer']: return redirect(url_for('index'))
    date=request.args.get('date',datetime.now().strftime('%Y-%m-%d')); rows=[]
    for st in suite_allowed_stores():
        try:
            logs,*_=fetch_attendance_data(date,date,'ALL',st)
            for x in logs: rows.append({**x,'store':st})
        except Exception: pass
    tr=''.join(f"<tr class=\"border-b\"><td class=\"p-2 font-bold\">{x['store']}</td><td>{x['user_id']}</td><td>{x['name']}</td><td>{x['dept']}</td><td>{x['store_in']}</td><td>{x['out_time']}</td><td>{x['status']}</td></tr>" for x in rows)
    body=f'''<form class="mb-4 flex gap-2"><input type="date" name="date" value="{date}" class="border p-2 rounded"><button class="bg-cyan-600 text-white px-4 rounded font-bold">Load</button><a href="/suite/all_stores_export?date={date}" class="bg-emerald-600 text-white px-4 py-2 rounded font-bold">Export Excel</a></form><div class="bg-white rounded-xl border overflow-auto"><table class="w-full text-xs min-w-[800px]"><thead class="bg-slate-100"><tr><th class="p-2">Store</th><th>ID</th><th>Name</th><th>Dept</th><th>In</th><th>Out</th><th>Status</th></tr></thead><tbody>{tr or '<tr><td colspan=7 class="p-8 text-center">No data</td></tr>'}</tbody></table></div>'''
    return suite_page('All Stores Attendance','🏬',body)

@app.route('/suite/all_stores_export')
def suite_all_stores_export():
    if session.get('role') not in ['admin','developer']: return redirect(url_for('index'))
    date=request.args.get('date',datetime.now().strftime('%Y-%m-%d')); wb=openpyxl.Workbook();ws=wb.active;ws.title='All Stores'
    ws.append(['Store','Date','Employee Code','Employee Name','Department','Store In','Out Time','Working Hours','Status'])
    for st in suite_allowed_stores():
        try:
            logs,*_=fetch_attendance_data(date,date,'ALL',st)
            for x in logs: ws.append([st,x['date'],x['user_id'],x['name'],x['dept'],x['store_in'],x['out_time'],x['total_hours'],x['status']])
        except Exception: pass
    out=io.BytesIO();wb.save(out);out.seek(0);suite_audit('EXPORT_ALL_STORES',date,store='ALL')
    return send_file(out,as_attachment=True,download_name=f'All_Stores_Attendance_{date}.xlsx',mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/suite/sync',methods=['GET','POST'])
def suite_sync():
    if not session_has_permission('biometric_sync'): return redirect(url_for('index'))
    if session.get('role') not in ['admin','developer']: return redirect(url_for('index'))
    if request.method=='POST':
        st=request.form.get('store','').upper()
        if suite_store_allowed(st):
            online=check_device_connectivity(st); rows=suite_load(SYNC_HISTORY_FILE,[]); rows.append({'time':datetime.now().strftime('%Y-%m-%d %H:%M:%S'),'store':st,'online':online,'user':session.get('user_name')});suite_save(SYNC_HISTORY_FILE,rows[-1000:]);suite_audit('MANUAL_MACHINE_CHECK',st,new='Online' if online else 'Offline',store=st)
    rows=list(reversed(suite_load(SYNC_HISTORY_FILE,[])))[:100]
    options=''.join(f'<option>{x}</option>' for x in suite_allowed_stores()); tr=''.join(f"<tr class='border-b'><td class='p-2'>{x['time']}</td><td>{x['store']}</td><td>{'Online' if x['online'] else 'Offline'}</td><td>{x['user']}</td></tr>" for x in rows)
    return suite_page('Biometric Sync History','🔄',f'''<form method="post" class="bg-white p-4 rounded-xl border mb-4 flex gap-2"><select name="store" class="border p-2 rounded">{options}</select><button class="bg-cyan-600 text-white px-4 rounded font-bold">Check Now</button></form><div class="bg-white border rounded-xl"><table class="w-full text-xs"><thead><tr><th>Time</th><th>Store</th><th>Status</th><th>User</th></tr></thead><tbody>{tr}</tbody></table></div>''')

@app.route('/suite/corrections',methods=['GET','POST'])
def suite_corrections():
    if not session.get('logged_in'): return redirect(url_for('login'))
    rows=suite_load(CORRECTIONS_FILE,[])
    if request.method=='POST' and session.get('role')=='employee':
        rows.append({'id':uuid.uuid4().hex,'user_id':session.get('user_id'),'name':session.get('user_name'),'store':session.get('store'),'date':request.form.get('date'),'field':request.form.get('field'),'expected_time':request.form.get('expected_time'),'reason':request.form.get('reason'),'status':'Pending','created_at':datetime.now().strftime('%Y-%m-%d %H:%M:%S')});suite_save(CORRECTIONS_FILE,rows);suite_notify(session.get('store'),'Correction Request',f"{session.get('user_name')} requested attendance correction",session.get('store'));suite_audit('SUBMIT_CORRECTION',session.get('user_id'),new=request.form.get('date'))
    visible=[x for x in rows if session.get('role')=='developer' or (session.get('role')=='admin' and x.get('store') in suite_allowed_stores()) or x.get('user_id')==session.get('user_id')]
    form='''<form method="post" class="grid grid-cols-1 md:grid-cols-5 gap-2 bg-white p-4 rounded-xl border mb-4"><input type="date" name="date" required class="border p-2 rounded"><select name="field" class="border p-2 rounded"><option value="store_in">Store In</option><option value="lunch_out">Lunch Out</option><option value="lunch_in">Lunch In</option><option value="out_time">Out Time</option></select><input type="time" step="1" name="expected_time" required class="border p-2 rounded"><input name="reason" required placeholder="Reason" class="border p-2 rounded"><button class="bg-indigo-600 text-white rounded font-bold">Submit</button></form>''' if session.get('role')=='employee' else ''
    tr=''.join(f"<tr class='border-b'><td class='p-2'>{x['date']}</td><td>{x['user_id']}<br>{x['name']}</td><td>{x['field']} → {x['expected_time']}</td><td>{x['reason']}</td><td>{x['status']}</td><td>{('<a class=\"text-emerald-600 font-bold\" href=\"/suite/correction/'+x['id']+'/approve\">Approve</a> · <a class=\"text-rose-600 font-bold\" href=\"/suite/correction/'+x['id']+'/reject\">Reject</a>') if session.get('role') in ['admin','developer'] and x['status']=='Pending' else '-'}</td></tr>" for x in reversed(visible))
    return suite_page('Attendance Corrections','✍️',form+f'''<div class="bg-white border rounded-xl overflow-auto"><table class="w-full text-xs"><thead><tr><th>Date</th><th>Employee</th><th>Change</th><th>Reason</th><th>Status</th><th>Action</th></tr></thead><tbody>{tr}</tbody></table></div>''')

@app.route('/suite/correction/<rid>/<action>')
def suite_correction_action(rid,action):
    if session.get('role') not in ['admin','developer']: return redirect(url_for('index'))
    rows=suite_load(CORRECTIONS_FILE,[])
    for x in rows:
        if x.get('id')==rid and x.get('store') in suite_allowed_stores():
            if action=='approve':
                old=load_overrides(); old.setdefault(x['date'],{}).setdefault(x['user_id'],{})[x['field']]=x['expected_time'];save_overrides(old);x['status']='Approved'
            elif action=='reject': x['status']='Rejected'
            suite_notify(x['user_id'],'Correction '+x['status'],f"Your correction for {x['date']} was {x['status']}",x['store']);suite_audit('CORRECTION_'+x['status'].upper(),x['user_id'],new=x['date'],store=x['store'])
    suite_save(CORRECTIONS_FILE,rows);return redirect(url_for('suite_corrections'))

@app.route('/suite/leave_balances',methods=['GET','POST'])
def suite_leave_balances():
    if not session_has_permission('leave_balances'): return redirect(url_for('index'))
    if session.get('role') not in ['admin','developer']: return redirect(url_for('index'))
    balances=suite_load(LEAVE_BALANCES_FILE,{})
    if request.method=='POST':
        uid=request.form.get('uid','').upper(); db=load_users_db();info=normalize_user_record(db.get(uid,{}))
        if info and suite_employee_store(info) in suite_allowed_stores(): balances[uid]={'annual':int(request.form.get('annual',18)),'used':int(request.form.get('used',0)),'pending':int(request.form.get('pending',0))};suite_save(LEAVE_BALANCES_FILE,balances);suite_audit('UPDATE_LEAVE_BALANCE',uid,new=balances[uid])
    db=load_users_db(); opts=[];rows=[]
    for uid,raw in db.items():
        info=normalize_user_record(raw);st=suite_employee_store(info)
        if info.get('role')!='employee' or st not in suite_allowed_stores():continue
        b=suite_leave_balance(uid);opts.append(f'<option value="{uid}">{uid} - {info.get("name")}</option>');rows.append(f'<tr class="border-b"><td class="p-2">{uid}</td><td>{info.get("name")}</td><td>{st}</td><td>{b["annual"]}</td><td>{b["used"]}</td><td>{b["pending"]}</td><td class="font-bold text-emerald-600">{b["available"]}</td></tr>')
    return suite_page('Leave Balances','🏖️',f'''<form method="post" class="bg-white border p-4 rounded-xl grid md:grid-cols-5 gap-2 mb-4"><select name="uid" class="border p-2 rounded">{''.join(opts)}</select><input type="number" name="annual" value="18" class="border p-2 rounded"><input type="number" name="used" value="0" class="border p-2 rounded"><input type="number" name="pending" value="0" class="border p-2 rounded"><button class="bg-emerald-600 text-white rounded">Save</button></form><div class="bg-white border rounded-xl overflow-auto"><table class="w-full text-xs"><thead><tr><th>ID</th><th>Name</th><th>Store</th><th>Annual</th><th>Used</th><th>Pending</th><th>Available</th></tr></thead><tbody>{''.join(rows)}</tbody></table></div>''')

@app.route('/suite/documents',methods=['GET','POST'])
def suite_documents():
    if not session.get('logged_in'): return redirect(url_for('login'))
    rows=suite_load(DOCUMENTS_FILE,[])
    if request.method=='POST' and session.get('role') in ['admin','developer']:
        uid=request.form.get('uid','').upper();db=load_users_db();info=normalize_user_record(db.get(uid,{}));f=request.files.get('document')
        if f and f.filename and suite_employee_store(info) in suite_allowed_stores():
            fn=secure_filename(f'{uid}_{uuid.uuid4().hex}_{f.filename}');f.save(os.path.join(EMPLOYEE_DOC_FOLDER,fn));rows.append({'id':uuid.uuid4().hex,'user_id':uid,'name':info.get('name',uid),'store':suite_employee_store(info),'type':request.form.get('type'),'filename':fn,'original':f.filename,'uploaded_at':datetime.now().strftime('%Y-%m-%d %H:%M:%S')});suite_save(DOCUMENTS_FILE,rows);suite_notify(uid,'New Document',f"{request.form.get('type')} uploaded",suite_employee_store(info));suite_audit('UPLOAD_DOCUMENT',uid,new=request.form.get('type'))
    visible=[x for x in rows if x.get('user_id')==session.get('user_id') or (session.get('role') in ['admin','developer'] and x.get('store') in suite_allowed_stores())]
    opts=''.join(f'<option value="{uid}">{uid} - {normalize_user_record(raw).get("name")}</option>' for uid,raw in load_users_db().items() if normalize_user_record(raw).get('role')=='employee' and suite_employee_store(normalize_user_record(raw)) in suite_allowed_stores())
    form=f'''<form method="post" enctype="multipart/form-data" class="bg-white border rounded-xl p-4 grid md:grid-cols-4 gap-2 mb-4"><select name="uid" class="border p-2 rounded">{opts}</select><select name="type" class="border p-2 rounded"><option>ID Document</option><option>Contract</option><option>Medical Certificate</option><option>Training Certificate</option><option>Other</option></select><input type="file" name="document" required class="border p-2 rounded"><button class="bg-indigo-600 text-white rounded">Upload</button></form>''' if session.get('role') in ['admin','developer'] else ''
    cards=''.join(f'''<div class="bg-white border rounded-xl p-4"><b>{x['type']}</b><p class="text-xs text-slate-500">{x['user_id']} · {x['name']} · {x['uploaded_at']}</p><a class="text-cyan-700 font-bold text-sm" href="/suite/document/{x['id']}">Open</a></div>''' for x in reversed(visible))
    return suite_page('Employee Documents','📁',form+f'<div class="grid md:grid-cols-3 gap-3">{cards or "No documents"}</div>')

@app.route('/suite/document/<did>')
def suite_document_open(did):
    rows=suite_load(DOCUMENTS_FILE,[]);x=next((a for a in rows if a.get('id')==did),None)
    if not x:return 'Not found',404
    if not (x.get('user_id')==session.get('user_id') or (session.get('role') in ['admin','developer'] and x.get('store') in suite_allowed_stores())):return 'Unauthorized',403
    return send_from_directory(EMPLOYEE_DOC_FOLDER,x['filename'],as_attachment=False)

@app.route('/suite/roster_copy',methods=['GET','POST'])
def suite_roster_copy():
    if not session_has_permission('roster_copy'): return redirect(url_for('index'))
    if session.get('role') not in ['admin','developer']:return redirect(url_for('index'))
    msg=''
    if request.method=='POST':
        source=request.form.get('source');target=request.form.get('target');scope=request.form.get('scope');uid=request.form.get('uid','').upper();roster=load_roster();db=load_users_db();count=0
        try:
            src=datetime.strptime(source,'%Y-%m-%d');dst=datetime.strptime(target,'%Y-%m-%d');days=7 if scope=='week' else ((src.replace(day=28)+timedelta(days=4)).replace(day=1)-src.replace(day=1)).days;src=src if scope=='week' else src.replace(day=1);dst=dst if scope=='week' else dst.replace(day=1)
            ids=[uid] if uid else [k for k,v in db.items() if normalize_user_record(v).get('role')=='employee' and suite_employee_store(normalize_user_record(v)) in suite_allowed_stores()]
            for emp in ids:
                roster.setdefault(emp,{})
                for i in range(days):
                    sv=(src+timedelta(days=i)).strftime('%Y-%m-%d');tv=(dst+timedelta(days=i)).strftime('%Y-%m-%d')
                    if sv in roster.get(emp,{}):roster[emp][tv]=roster[emp][sv];count+=1
            save_roster(roster);suite_audit('COPY_ROSTER',uid or 'ALL',old=source,new=target);msg=f'{count} roster cells copied.'
        except Exception as e:msg='Copy failed: '+str(e)
    return suite_page('Roster Copy','📋',f'''<div class="bg-white border rounded-xl p-5"><p class="text-emerald-700 font-bold mb-3">{msg}</p><form method="post" class="grid md:grid-cols-5 gap-3"><select name="scope" class="border p-2 rounded"><option value="week">Copy Week</option><option value="month">Copy Month</option></select><input type="date" name="source" required class="border p-2 rounded"><input type="date" name="target" required class="border p-2 rounded"><input name="uid" placeholder="Employee ID or blank for all" class="border p-2 rounded"><button class="bg-indigo-600 text-white rounded font-bold">Copy</button></form><p class="text-xs text-slate-500 mt-3">Week copy preserves seven consecutive days; monthly copy preserves date positions.</p></div>''')

@app.route('/suite/staffing',methods=['GET','POST'])
def suite_staffing():
    if not session_has_permission('staffing_rules'): return redirect(url_for('index'))
    if session.get('role') not in ['admin','developer']:return redirect(url_for('index'))
    rules=suite_load(STAFFING_RULES_FILE,[])
    if request.method=='POST':
        rules=[x for x in rules if not (x.get('store')==request.form.get('store') and x.get('designation')==request.form.get('designation'))];rules.append({'store':request.form.get('store'),'designation':request.form.get('designation'),'shift_a':int(request.form.get('shift_a',0)),'shift_b':int(request.form.get('shift_b',0))});suite_save(STAFFING_RULES_FILE,rules);suite_audit('UPDATE_STAFFING_RULE',request.form.get('designation'),new=request.form.to_dict(),store=request.form.get('store'))
    opts=''.join(f'<option>{x}</option>' for x in suite_allowed_stores());tr=''.join(f"<tr class='border-b'><td class='p-2'>{x['store']}</td><td>{x['designation']}</td><td>{x['shift_a']}</td><td>{x['shift_b']}</td></tr>" for x in rules if x.get('store') in suite_allowed_stores())
    return suite_page('Minimum Staffing Rules','👥',f'''<form method="post" class="bg-white border rounded-xl p-4 grid md:grid-cols-5 gap-2 mb-4"><select name="store" class="border p-2 rounded">{opts}</select><input name="designation" required placeholder="Designation" class="border p-2 rounded"><input type="number" name="shift_a" value="1" class="border p-2 rounded"><input type="number" name="shift_b" value="1" class="border p-2 rounded"><button class="bg-emerald-600 text-white rounded">Save Rule</button></form><div class="bg-white border rounded-xl"><table class="w-full text-xs"><thead><tr><th>Store</th><th>Designation</th><th>Shift A Min</th><th>Shift B Min</th></tr></thead><tbody>{tr}</tbody></table></div>''')

@app.route('/suite/audit')
def suite_audit_page():
    if not session_has_permission('audit_log'): return redirect(url_for('index'))
    if session.get('role')!='developer':return redirect(url_for('index'))
    rows=list(reversed(suite_load(AUDIT_LOG_FILE,[])))[:500];tr=''.join(f"<tr class='border-b'><td class='p-2'>{x['time']}</td><td>{x['user_name']}<br>{x['role']}</td><td>{x['store']}</td><td>{x['action']}</td><td>{x['target']}</td><td>{x['old']} → {x['new']}</td><td>{x['ip']}</td></tr>" for x in rows)
    return suite_page('Audit Log','🛡️',f'<div class="bg-white border rounded-xl overflow-auto"><table class="w-full text-xs min-w-[900px]"><thead><tr><th>Time</th><th>User</th><th>Store</th><th>Action</th><th>Target</th><th>Change</th><th>IP</th></tr></thead><tbody>{tr}</tbody></table></div>')

@app.route('/suite/notifications')
def suite_notifications():
    if not session.get('logged_in'):return redirect(url_for('login'))
    rows=suite_load(NOTIFICATIONS_FILE,[]);visible=[x for x in rows if (session.get('role')=='employee' and x.get('user_id')==session.get('user_id') and 'approved' in x.get('title','').lower()) or (session.get('role')=='admin' and x.get('user_id') in [session.get('store'),session.get('user_id')]) or session.get('role')=='developer'];cards=''.join(f"<div class='bg-white border rounded-xl p-4'><b>{x['title']}</b><p class='text-sm'>{x['message']}</p><span class='text-xs text-slate-500'>{x['created_at']}</span></div>" for x in reversed(visible))
    return suite_page('Notification Center','🔔',f'<div class="space-y-2">{cards or "No notifications"}</div>')

@app.route('/suite/qr_directory')
def suite_qr_directory():
    if not session_has_permission('qr_directory'): return redirect(url_for('index'))
    if session.get('role') not in ['admin','developer']:return redirect(url_for('index'))
    cards=[]
    for uid,raw in load_users_db().items():
        info=normalize_user_record(raw);st=suite_employee_store(info)
        if info.get('role')=='employee' and st in suite_allowed_stores():cards.append(f'''<div class="bg-white border rounded-xl p-4 text-center"><img src="/suite/employee_qr/{uid}" class="w-32 h-32 mx-auto"><b>{info.get('name')}</b><p class="text-xs">{uid} · {st} · {info.get('designation')}</p></div>''')
    return suite_page('Employee QR Directory','▦',f'<div class="grid md:grid-cols-4 gap-3">{"".join(cards)}</div>')

@app.route('/suite/employee_qr/<uid>')
def suite_employee_qr(uid):
    db=load_users_db();info=normalize_user_record(db.get(uid,{}));st=suite_employee_store(info)
    if not info or not (session.get('role')=='developer' or st in suite_allowed_stores() or uid==session.get('user_id')):return 'Unauthorized',403
    payload=json.dumps({'employee_code':uid,'name':info.get('name'),'store':st,'designation':info.get('designation'),'status':info.get('status')},ensure_ascii=False)
    img=qrcode.make(payload);out=io.BytesIO();img.save(out,format='PNG');out.seek(0);return send_file(out,mimetype='image/png')



if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False, threaded=True)

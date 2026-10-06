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
from flask import Flask, render_template_string, request, Response, send_file, session, redirect, url_for, flash, send_from_directory
from github import Github
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from werkzeug.utils import secure_filename
from zk import ZK, const
from pypdf import PdfReader, PdfWriter

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'gamek_fresmart_secret_key_sonu')

UPLOAD_FOLDER = os.getenv('UPLOAD_FOLDER', 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
ALLOWED_EXTENSIONS = {'pdf', 'png', 'jpg', 'jpeg', 'doc', 'docx'}

# Machine Configurations with Multiple Stores
MACHINES = {
    'LM11': {'ip': os.getenv('MACHINE_IP', '192.168.1.153'), 'port': int(os.getenv('MACHINE_PORT', 4370)), 'name': 'Gamek HRMS'},
    'LF07': {'ip': '192.168.88.16', 'port': 4370, 'name': 'Benfica LF07 HRMS'},
    'DEV': {'ip': '127.0.0.1', 'port': 4370, 'name': 'Fresmart'}
}

GITHUB_TOKEN = os.getenv('GITHUB_TOKEN', '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')
GITHUB_REPO_NAME = os.getenv('GITHUB_REPO_NAME', 'devilsuza/Gamekfme')
GITHUB_BRANCH = os.getenv('GITHUB_BRANCH', 'main')

# Persistent Storage Files
LEAVE_JSON_FILE = 'leave_requests.json'
LEAVE_EXCEL_FILE = 'leave_records.xlsx'
ATTENDANCE_OVERRIDES_FILE = 'attendance_overrides.json'
MANUAL_PUNCHES_FILE = 'manual_punches.json'
ROSTER_JSON_FILE = 'roster.json' 
SHIFT_REQUESTS_FILE = 'shift_requests.json' 
PASSWORD_RESETS_FILE = 'password_resets.json'
USERS_DB_FILE = 'users_db.json'
SALARY_SLIPS_FILE = 'salary_slips.json'

SYNCED_ATTENDANCE_LOGS = []
LAST_DEVICE_SYNC_TIME = None

# Master Employees (Fallback / Default Setup)
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

# Database initialization
def load_users_db():
    db = load_json_file(USERS_DB_FILE)
    if not db:
        db = {}
        for k, v in MASTER_EMPLOYEES.items():
            db[k] = {'name': v['name'], 'password': '123', 'store': 'LM11', 'status': 'active', 'role': 'employee', 'off': v.get('off', 'SUNDAY'), 'dept': v.get('dept', 'General'), 'shift': v.get('shift', 'morning')}
        # Admins
        db['LM11'] = {'name': 'Admin (LM11)', 'password': os.getenv('ADMIN_PWD', 'Gamek@789'), 'store': 'LM11', 'status': 'active', 'role': 'admin', 'off': 'SUNDAY', 'dept': 'Admin', 'shift': 'morning'}
        db['LF07'] = {'name': 'Admin (LF07)', 'password': '123', 'store': 'LF07', 'status': 'active', 'role': 'admin', 'off': 'SUNDAY', 'dept': 'Admin', 'shift': 'morning'}
        save_json_file(USERS_DB_FILE, db)
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
    if LAST_DEVICE_SYNC_TIME and (datetime.now() - LAST_DEVICE_SYNC_TIME).total_seconds() < 300: return True
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
            
            raw_punches_list.append({'date': att_date_str, 'time': att_ts.strftime('%H:%M:%S'), 'user_id': emp_code, 'name': emp_name, 'timestamp': att_ts})
            if att_date_str not in period_data: period_data[att_date_str] = {}
            if emp_code not in period_data[att_date_str]: period_data[att_date_str][emp_code] = {'name': emp_name, 'timestamps': []}
            period_data[att_date_str][emp_code]['timestamps'].append(att_ts)
            
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
                
                net_duration_seconds = max(0, (s_out - s_in) - lunch_seconds) if (s_in is not None and s_out is not None and s_out > s_in) else 0
                total_duration_seconds += net_duration_seconds
                total_hours_str = f"{divmod(net_duration_seconds, 3600)[0]}h {divmod(net_duration_seconds, 3600)[1]//60}m" if s_in is not None and s_out is not None else "-"
                
                net_variance_str, variance_type = "-", "neutral"
                if s_in is not None and s_out is not None:
                    target_seconds = (8 * 3600) if has_lunch_punches else (7 * 3600)
                    diff = net_duration_seconds - target_seconds
                    total_net_variance_seconds += diff
                    if diff > (45 * 60):
                        extra_hours_val = math.floor(diff / 3600) or 1
                        code_prefix = 'H07' if is_weekend else 'H06'
                        net_variance_str, variance_type = f"{code_prefix};{extra_hours_val}", 'positive'
                    elif diff < 0:
                        s_hrs = divmod(abs(diff), 3600)
                        net_variance_str, variance_type = f"-{s_hrs[0]}h {s_hrs[1]//60}m", 'negative'
                    else: net_variance_str, variance_type = "0h 0m", 'neutral'
                
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
                    'status': status, 'is_late': 'Yes' if is_late else 'No', 'shift_type': shift_type
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
<body class="bg-slate-900 min-h-screen flex items-center justify-center p-4">
    <div class="bg-white/95 backdrop-blur-md rounded-3xl shadow-2xl border border-slate-200/50 p-8 w-full max-w-md space-y-6">
        <div class="text-center space-y-2">
            <div class="inline-flex bg-[#78b13f] px-5 py-3 rounded-2xl shadow-lg mb-2 items-center justify-center">
                <img src="{{ url_for('static', filename='fresmart.png') }}" alt="Logo" class="h-12 object-contain" onerror="this.style.display='none'">
            </div>
            <h1 class="text-2xl font-black text-slate-900 tracking-tight">Attendance Portal</h1>
            <p class="text-xs text-slate-500 font-medium">Developed by Sonu Kumar <span class="text-emerald-600 font-semibold">(NCSA0608)</span></p>
        </div>
        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}{% for category, message in messages %}<div class="bg-{{ 'emerald' if category == 'success' else 'rose' }}-50 border-{{ 'emerald' if category == 'success' else 'rose' }}-200 text-{{ 'emerald' if category == 'success' else 'rose' }}-800 border text-xs font-semibold p-3.5 rounded-xl text-center">{{ message }}</div>{% endfor %}{% endif %}
        {% endwith %}
        {% if error %}<div class="bg-rose-50 border border-rose-200 text-rose-700 text-xs font-semibold p-3.5 rounded-xl text-center">{{ error }}</div>{% endif %}
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
<body class="bg-slate-900 min-h-screen flex items-center justify-center p-4">
    <div class="bg-white/95 backdrop-blur-md rounded-3xl shadow-2xl border border-slate-200/50 p-8 w-full max-w-md space-y-6">
        <div class="text-center space-y-2">
            <div class="inline-flex bg-amber-500 px-5 py-3 rounded-2xl shadow-lg mb-2 items-center justify-center">
                <span class="text-3xl">🔑</span>
            </div>
            <h1 class="text-2xl font-black text-slate-900 tracking-tight">Reset Password</h1>
            <p class="text-xs text-slate-500 font-medium">Create a new password request</p>
        </div>

        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}
                {% for category, message in messages %}
                <div class="{% if category == 'success' %}bg-emerald-50 border-emerald-200 text-emerald-800{% else %}bg-rose-50 border-rose-200 text-rose-700{% endif %} border text-xs font-semibold p-3.5 rounded-xl text-center">
                    {{ message }}
                </div>
                {% endfor %}
            {% endif %}
        {% endwith %}

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
</body>
</html>
"""

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
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
                alert("Session timeout ho gaya hai. Dobara login karein.");
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

        let timeLeft = 180;
        let timerInterval;
        function startTimer() {
            clearInterval(timerInterval);
            timeLeft = 180;
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

        function secureShutdown() {
            let pwd = prompt("Server band karne ke liye password enter karein:");
            if (pwd) window.location.href = "/shutdown?pwd=" + encodeURIComponent(pwd);
        }

        function loadRosterPlanner() {
            let empId = document.getElementById('rp-emp').value;
            let monthStr = document.getElementById('rp-month').value;
            let tbody = document.getElementById('roster-planner-tbody');
            if(!empId || !monthStr) {
                tbody.innerHTML = '<tr><td colspan="2" class="text-center py-8 text-slate-400">Employee & Month select karein</td></tr>';
                return;
            }
            
            tbody.innerHTML = '<tr><td colspan="2" class="text-center py-8 text-slate-500">Loading roster...</td></tr>';
            
            fetch(`/api/get_roster?emp_id=${empId}`)
            .then(r => r.json())
            .then(data => {
                let parts = monthStr.split('-');
                let year = parseInt(parts[0]);
                let month = parseInt(parts[1]);
                let daysInMonth = new Date(year, month, 0).getDate();
                
                tbody.innerHTML = '';
                for(let i=1; i<=daysInMonth; i++) {
                    let dateStr = `${year}-${String(month).padStart(2,'0')}-${String(i).padStart(2,'0')}`;
                    let dateObj = new Date(year, month-1, i);
                    let dayName = dateObj.toLocaleDateString('en-US', {weekday:'long'});
                    let val = data[dateStr] || '';

                    let tr = document.createElement('tr');
                    tr.className = "hover:bg-slate-50";
                    tr.innerHTML = `
                        <td class="py-2.5 px-3 border-b border-slate-200 font-mono text-xs text-slate-600">${dateStr} <strong class="ml-2 text-indigo-500">${dayName}</strong></td>
                        <td class="py-2.5 px-3 border-b border-slate-200">
                            <select class="roster-select w-full bg-white border border-slate-300 rounded-xl px-3 py-1.5 text-xs font-bold focus:ring-2 focus:ring-indigo-500" data-date="${dateStr}" data-day="${dayName}" onchange="checkAutoReflect(this)">
                                <option value="" class="text-slate-400">-- Default --</option>
                                <option value="Shift A" ${val=='Shift A'?'selected':''} class="text-blue-600">Shift A (06:00)</option>
                                <option value="Shift B" ${val=='Shift B'?'selected':''} class="text-indigo-600">Shift B (10:00+)</option>
                                <option value="Weekly Off" ${val=='Weekly Off'?'selected':''} class="text-rose-600">Weekly Off</option>
                            </select>
                        </td>
                    `;
                    tbody.appendChild(tr);
                }
                document.getElementById('save-roster-btn').classList.remove('hidden');
            });
        }

        function checkAutoReflect(selectEl) {
            let val = selectEl.value;
            let dayName = selectEl.getAttribute('data-day');
            let dateStr = selectEl.getAttribute('data-date');
            if(val === 'Weekly Off') {
                if(confirm(`Kya aap iske aage ke sabhi "${dayName}" ko Weekly Off set karna chahte hain?`)) {
                    let selects = document.querySelectorAll('.roster-select');
                    selects.forEach(sel => {
                        let selDate = sel.getAttribute('data-date');
                        let selDay = sel.getAttribute('data-day');
                        if(selDay === dayName && selDate > dateStr) {
                            sel.value = 'Weekly Off';
                        }
                    });
                }
            }
        }

        function saveRoster() {
            let empId = document.getElementById('rp-emp').value;
            let selects = document.querySelectorAll('.roster-select');
            let payload = {emp_id: empId, updates: {}};
            selects.forEach(sel => {
                if(sel.value !== '') {
                    payload.updates[sel.getAttribute('data-date')] = sel.value;
                }
            });
            fetch('/api/save_roster', {
                method: 'POST', headers: {'Content-Type': 'application/json'},
                body: JSON.stringify(payload)
            }).then(r=>r.json()).then(res=>{
                alert('Roster (Weekly Off & Shifts) successfully save ho gaya hai!');
                toggleModal('roster-planner-modal', false);
                window.location.reload();
            });
        }

        function filterSalarySlips() {
            let month = document.getElementById('salary-month-select').value;
            let rows = document.querySelectorAll('.salary-row');
            rows.forEach(row => {
                if(month === '' || row.getAttribute('data-month') === month) {
                    row.style.display = '';
                } else {
                    row.style.display = 'none';
                }
            });
        }

        function openPdfViewer(url) {
            let viewer = document.getElementById('pdf-viewer-frame');
            viewer.src = url;
            toggleModal('pdf-viewer-modal', true);
        }

        function uploadLoading() {
            document.getElementById('upload-btn-text').innerText = 'Splitting... Please wait...';
            document.getElementById('upload-btn').disabled = true;
            document.getElementById('upload-btn').classList.add('opacity-75', 'cursor-wait');
        }

        document.addEventListener('DOMContentLoaded', function() {
            startTimer();
            setInterval(updateLiveClock, 1000);
            updateLiveClock();
        });
    </script>
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
                
                {% if role in ['admin', 'developer'] %}
                <a href="#" onclick="toggleModal('roster-planner-modal', true); return false;" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs transition border border-transparent hover:border-slate-200">
                    <span>🗓</span>
                    <span>Roster Planner</span>
                </a>
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
                <a href="#" onclick="toggleModal('user-mgmt-modal', true); return false;" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs transition border border-transparent hover:border-slate-200">
                    <span>🪪</span>
                    <span>User Management</span>
                </a>
                {% endif %}
                
                {% if role in ['admin', 'developer'] %}
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
                <a href="#" onclick="toggleModal('leave-modal', true); return false;" class="flex items-center justify-between px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs transition">
                    <div class="flex items-center space-x-3">
                        <span>🏖️</span>
                        <span>Leave Management</span>
                    </div>
                    {% if role in ['admin', 'developer'] and pending_leaves_count > 0 %}
                    <span class="bg-rose-500 text-white text-[10px] font-bold px-2 py-0.5 rounded-full animate-pulse">{{ pending_leaves_count }}</span>
                    {% endif %}
                </a>
                
                <p class="text-[10px] font-bold uppercase tracking-wider text-slate-400 px-3 mt-6 mb-2">Team Management</p>
                <a href="#" onclick="toggleModal('salary-modal', true); return false;" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs transition border border-transparent hover:border-slate-200">
                    <span>💰</span>
                    <span>Payroll & Reports</span>
                </a>
                <a href="#" onclick="toggleModal('roster-modal', true); return false;" class="flex items-center space-x-3 px-3 py-2.5 rounded-xl text-slate-600 hover:bg-slate-100 font-medium text-xs transition">
                    <span>👥</span>
                    <span>Employees Info</span>
                </a>
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
                    <h1 class="text-sm sm:text-base font-black text-slate-900 tracking-tight">Dashboard</h1>
                    <span class="hidden sm:inline-block text-xs text-slate-400 font-medium">| Good day, {{ logged_user_name }}</span>
                </div>
            </div>

            <div class="flex items-center gap-2 flex-wrap justify-end">
                <button onclick="toggleModal('leave-modal', true)" class="relative bg-emerald-50 hover:bg-emerald-100 text-emerald-800 text-xs font-bold px-2 py-1.5 sm:px-3 sm:py-2 rounded-xl transition border border-emerald-200 flex items-center space-x-1.5">
                    <span class="hidden sm:inline">🏖 Leave Portal</span>
                    <span class="sm:hidden">🏖 Leaves</span>
                    {% if role in ['admin', 'developer'] and pending_leaves_count > 0 %}
                    <span class="bg-rose-500 text-white text-[10px] px-1.5 py-0.2 rounded-full font-black animate-bounce">{{ pending_leaves_count }}</span>
                    {% endif %}
                </button>

                <!-- Clock and Status (Hidden on very small screens to save space) -->
                <div class="hidden md:flex text-xs bg-slate-50 px-3 py-2 rounded-xl border border-slate-200 items-center space-x-2">
                    <span class="h-2 w-2 {% if stats.device_online %}bg-emerald-500{% else %}bg-red-500{% endif %} rounded-full animate-pulse"></span>
                    <span class="text-slate-600 font-medium">Device: <strong class="{% if stats.device_online %}text-emerald-600{% else %}text-red-600{% endif %}">{% if stats.device_online %}Online{% else %}Offline{% endif %}</strong></span>
                    <span class="text-slate-300">|</span>
                    <span id="live-digital-clock" class="text-slate-700 font-semibold"></span>
                    <span class="text-slate-300">|</span>
                    <span class="text-slate-500">Sync: <strong id="countdown-timer" class="text-emerald-600 font-mono">03:00</strong></span>
                </div>

                <div class="flex items-center space-x-2 bg-slate-100 border border-slate-200 px-2.5 py-1.5 sm:px-3 sm:py-1.5 rounded-xl text-xs font-bold text-slate-700">
                    <span class="hidden sm:inline">👤 {{ logged_user_name }}</span>
                    <a href="/logout" class="text-rose-600 hover:text-rose-700 sm:ml-2 font-semibold">Logout 🔒</a>
                </div>

                {% if role == 'admin' or role == 'developer' %}
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
                {% if messages %}
                    {% for category, message in messages %}
                    <div class="{% if category == 'success' %}bg-emerald-50 border-emerald-200 text-emerald-800{% else %}bg-rose-50 border-rose-200 text-rose-700{% endif %} border text-xs font-bold p-4 rounded-2xl shadow-sm flex items-center justify-between">
                        <span>{{ message }}</span>
                        <span class="cursor-pointer text-lg px-2" onclick="this.parentElement.style.display='none'">✕</span>
                    </div>
                    {% endfor %}
                {% endif %}
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
                                    <td class="py-3 px-4 font-bold nowrap-cell font-mono text-xs {% if log.variance_type == 'positive' %}text-emerald-600{% elif log.variance_type == 'negative' %}text-rose-600{% else %}text-slate-600{% endif %}">{{ log.net_variance }}</td>
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

    <!-- Roster Planner Modal (Admin/Dev) -->
    <div id="roster-planner-modal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[60] flex items-center justify-center hidden p-2">
        <div class="bg-white rounded-2xl shadow-2xl border border-slate-200 p-4 w-full max-w-4xl mx-auto space-y-4 max-h-[90vh] flex flex-col">
            <div class="flex justify-between items-center border-b border-slate-100 pb-2">
                <h3 class="text-sm sm:text-lg font-bold text-slate-900">🗓️ Roster Planner</h3>
                <button onclick="toggleModal('roster-planner-modal', false)" class="text-slate-400 hover:text-slate-600 font-bold text-lg">✕</button>
            </div>
            <div class="flex flex-col sm:flex-row gap-2 items-end">
                <div class="w-full sm:flex-1">
                    <label class="block text-[10px] font-bold uppercase text-slate-500 mb-1">Employee</label>
                    <select id="rp-emp" class="w-full bg-slate-50 border border-slate-300 rounded-xl px-2 py-1.5 text-xs">
                        <option value="">-- Choose --</option>
                        {% for emp in all_users %}
                            <option value="{{ emp.user_id }}">{{ emp.name }} ({{ emp.user_id }})</option>
                        {% endfor %}
                    </select>
                </div>
                <div class="w-full sm:flex-1">
                    <label class="block text-[10px] font-bold uppercase text-slate-500 mb-1">Month</label>
                    <input type="month" id="rp-month" class="w-full bg-slate-50 border border-slate-300 rounded-xl px-2 py-1.5 text-xs">
                </div>
                <button onclick="loadRosterPlanner()" class="w-full sm:w-auto bg-indigo-600 text-white font-bold text-[10px] sm:text-xs uppercase px-4 py-2 rounded-xl">Load</button>
            </div>
            
            <div class="overflow-y-auto flex-1 border border-slate-200 rounded-xl">
                <table class="w-full text-left">
                    <thead class="bg-slate-100 text-slate-600 uppercase text-[9px] sm:text-[10px] font-bold sticky top-0">
                        <tr>
                            <th class="py-2 px-3">Date & Day</th>
                            <th class="py-2 px-3">Shift Allocation</th>
                        </tr>
                    </thead>
                    <tbody id="roster-planner-tbody" class="divide-y divide-slate-100">
                        <tr><td colspan="2" class="text-center py-6 text-slate-400 text-xs">Select Employee & Month</td></tr>
                    </tbody>
                </table>
            </div>
            <div class="pt-2 flex justify-end gap-2">
                <button id="save-roster-btn" onclick="saveRoster()" class="w-full sm:w-auto px-5 py-2 bg-emerald-600 text-white text-xs font-bold uppercase rounded-xl hidden">Save</button>
            </div>
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
    <div id="user-mgmt-modal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[60] flex items-center justify-center hidden p-2">
        <div class="bg-white rounded-2xl shadow-2xl border border-slate-200 p-4 w-full max-w-4xl mx-auto space-y-4 max-h-[90vh] flex flex-col">
            <div class="flex justify-between items-center border-b border-slate-100 pb-2">
                <h3 class="text-sm sm:text-lg font-bold text-slate-900">🪪 User Management</h3>
                <button onclick="toggleModal('user-mgmt-modal', false)" class="text-slate-400 hover:text-slate-600 font-bold text-lg">✕</button>
            </div>
            <div class="p-3 bg-slate-50 border border-slate-200 rounded-xl">
                <form action="/manage_user" method="POST" class="flex flex-wrap gap-2 items-end">
                    <input type="hidden" name="action" value="create">
                    <div>
                        <label class="block text-[10px] font-bold text-slate-500 mb-1">Emp ID</label>
                        <input type="text" name="uid" placeholder="NWC..." required class="w-24 px-2 py-1.5 border border-slate-300 rounded-lg text-xs focus:outline-none">
                    </div>
                    <div>
                        <label class="block text-[10px] font-bold text-slate-500 mb-1">Name</label>
                        <input type="text" name="name" placeholder="Full Name" required class="w-40 px-2 py-1.5 border border-slate-300 rounded-lg text-xs focus:outline-none">
                    </div>
                    <div>
                        <label class="block text-[10px] font-bold text-slate-500 mb-1">Password</label>
                        <input type="text" name="password" value="123" required class="w-20 px-2 py-1.5 border border-slate-300 rounded-lg text-xs focus:outline-none">
                    </div>
                    <div>
                        <label class="block text-[10px] font-bold text-slate-500 mb-1">Store Map</label>
                        <select name="store" class="w-24 px-2 py-1.5 border border-slate-300 rounded-lg text-xs focus:outline-none">
                            <option value="LM11">LM11</option>
                            <option value="LF07">LF07</option>
                        </select>
                    </div>
                    <button type="submit" class="bg-emerald-600 text-white font-bold px-4 py-1.5 rounded-lg text-xs uppercase tracking-wider mb-0.5 shadow">Add / Update</button>
                </form>
            </div>
            <div class="overflow-y-auto flex-1 border border-slate-200 rounded-xl">
                <table class="w-full text-left min-w-[500px]">
                    <thead class="bg-slate-100 text-slate-600 uppercase text-[9px] font-bold sticky top-0">
                        <tr>
                            <th class="py-2 px-3">Emp ID</th>
                            <th class="py-2 px-3">Name</th>
                            <th class="py-2 px-3 text-center">Store</th>
                            <th class="py-2 px-3 text-center">Status</th>
                            <th class="py-2 px-3 text-center">Action</th>
                        </tr>
                    </thead>
                    <tbody class="divide-y divide-slate-100 text-xs">
                        {% for u_id, u_info in users_db.items() %}
                        <tr class="hover:bg-slate-50">
                            <td class="py-2 px-3 font-mono text-slate-600">{{ u_id }}</td>
                            <td class="py-2 px-3 font-bold">{{ u_info.name }}</td>
                            <td class="py-2 px-3 text-center font-bold text-indigo-600">{{ u_info.store }}</td>
                            <td class="py-2 px-3 text-center">
                                <span class="px-2 py-0.5 rounded-full text-[10px] font-bold {% if u_info.status == 'active' %}bg-emerald-100 text-emerald-700{% else %}bg-rose-100 text-rose-700{% endif %}">
                                    {{ u_info.status|upper }}
                                </span>
                            </td>
                            <td class="py-2 px-3 text-center">
                                <div class="flex justify-center gap-1">
                                    <form action="/manage_user" method="POST" class="inline">
                                        <input type="hidden" name="uid" value="{{ u_id }}">
                                        <input type="hidden" name="action" value="toggle_status">
                                        <button type="submit" class="bg-amber-100 hover:bg-amber-200 text-amber-700 px-2 py-1 rounded text-[10px] font-bold border border-amber-200">Toggle</button>
                                    </form>
                                    <form action="/manage_user" method="POST" class="inline" onsubmit="return confirm('Delete this user?');">
                                        <input type="hidden" name="uid" value="{{ u_id }}">
                                        <input type="hidden" name="action" value="delete">
                                        <button type="submit" class="bg-rose-100 hover:bg-rose-200 text-rose-700 px-2 py-1 rounded text-[10px] font-bold border border-rose-200">Del</button>
                                    </form>
                                </div>
                            </td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </div>
    </div>
</body>
</html>
"""

@app.route('/login', methods=['GET', 'POST'])
def login():
    failed_attempts = session.get('failed_attempts', 0)
    
    if request.method == 'POST':
        uid = request.form.get('user_id').strip().upper()
        pwd = request.form.get('password').strip()
        
        dev_pass = os.getenv('DEV_PWD', 'Shama@8577')
        db = load_users_db()
        
        if uid == 'NCSA0608' and pwd == dev_pass:
            session.update({'logged_in': True, 'role': 'developer', 'user_id': uid, 'user_name': 'Sonu Kumar (Dev)', 'store': 'DEV'})
            return redirect(url_for('index'))
            
        emp_key = f"NWC{uid}" if not uid.startswith('NWC') and uid not in ['LM11', 'LF07'] else uid
        
        if emp_key in db:
            user_data = db[emp_key]
            if user_data.get('status') == 'blocked':
                return render_template_string(LOGIN_TEMPLATE, error="Account is Suspended/Blocked. Contact Developer.")
            
            if pwd == user_data.get('password'):
                session.update({'logged_in': True, 'role': user_data['role'], 'user_id': emp_key, 'user_name': user_data['name'], 'store': user_data['store']})
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
    if session.get('role') != 'developer': return redirect(url_for('login'))
    db = load_users_db()
    action = request.form.get('action')
    uid = request.form.get('uid').strip().upper()
    uid = f"NWC{uid}" if not uid.startswith('NWC') and uid not in ['LM11', 'LF07'] else uid
    
    if action == 'create':
        db[uid] = {
            'name': request.form.get('name').strip(),
            'password': request.form.get('password').strip(),
            'store': request.form.get('store'),
            'status': 'active', 'role': 'employee', 'off': 'SUNDAY', 'dept': 'General', 'shift': 'morning'
        }
        flash(f"User {uid} created/updated successfully.", "success")
    elif action == 'toggle_status' and uid in db:
        db[uid]['status'] = 'blocked' if db[uid].get('status') == 'active' else 'active'
        flash(f"User {uid} status changed to {db[uid]['status']}.", "success")
    elif action == 'delete' and uid in db:
        del db[uid]
        flash(f"User {uid} deleted permanently.", "success")
        
    save_json_file(USERS_DB_FILE, db)
    return redirect(url_for('index'))

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
        
    role = session.get('role')
    store = session.get('store', 'LM11')
    portal_name = MACHINES.get(store, MACHINES['LM11'])['name']
    
    logged_user_id = session.get('user_id')
    logged_user_name = session.get('user_name')
    
    if role == 'employee':
        selected_emp = logged_user_id
    else:
        selected_emp = request.args.get('employee', 'ALL')
        
    today_str = datetime.now().strftime('%Y-%m-%d')
    start_date = request.args.get('start_date', today_str)
    end_date = request.args.get('end_date', today_str)
    
    logs, all_users, g_hrs, g_l_hrs, g_var, raw_punches, stats = fetch_attendance_data(start_date, end_date, selected_emp, store)
    
    pending_leaves_count = sum(1 for req in LEAVE_REQUESTS if req['status'] == 'Pending')
    current_user_leave_requests = [req for req in LEAVE_REQUESTS if req['user_id'] == logged_user_id] if role == 'employee' else LEAVE_REQUESTS
        
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
    
    all_salary_slips = load_salary_slips()
    my_salary_slips = [s for s in all_salary_slips if s['user_id'] == logged_user_id] if role == 'employee' else all_salary_slips
    
    users_db = load_users_db() if role == 'developer' else {}
    
    return render_template_string(
        HTML_TEMPLATE,
        logs=logs, all_users=all_users, start_date=start_date, end_date=end_date, selected_emp=selected_emp,
        grand_total_hours=g_hrs, grand_total_lunch_hours=g_l_hrs, grand_total_variance=g_var, stats=stats,
        raw_punches=raw_punches, role=role, logged_user_name=logged_user_name, portal_name=portal_name,
        leave_requests=current_user_leave_requests, pending_leaves_count=pending_leaves_count,
        shift_requests=my_shift_reqs, pending_shifts_count=pending_shifts_count,
        reset_requests=reset_requests, pending_resets_count=pending_resets_count,
        salary_slips=my_salary_slips, users_db=users_db
    )

# --- BULK SALARY SLIP APIs ---
@app.route('/upload_bulk_salary', methods=['POST'])
def upload_bulk_salary():
    if session.get('role') not in ['admin', 'developer']:
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
    elif session.get('role') not in ['admin', 'developer']:
        return 'Unauthorized Access', 403

    file_path = os.path.join(app.config['UPLOAD_FOLDER'], target_slip.get('filename', ''))
    if not os.path.isfile(file_path):
        return 'Salary slip file missing on server', 404
    return send_file(file_path, mimetype='application/pdf')

@app.route('/delete_salary/<file_id>')
def delete_salary(file_id):
    if session.get('role') not in ['admin', 'developer']:
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

# --- ROSTER PLANNER APIs ---
@app.route('/api/get_roster', methods=['GET'])
def api_get_roster():
    if session.get('role') not in ['admin', 'developer']: return {}, 403
    emp_id = request.args.get('emp_id')
    roster = load_roster()
    return roster.get(emp_id, {})

@app.route('/api/save_roster', methods=['POST'])
def api_save_roster():
    if session.get('role') not in ['admin', 'developer']: return {}, 403
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
    if session.get('role') not in ['admin', 'developer']: return redirect(url_for('login'))
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

# --- QUICK EDIT API ---
@app.route('/quick_edit')
def quick_edit():
    if session.get('role') != 'developer': return redirect(url_for('index'))
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
        if len(new_time.split(':')) == 2: new_time += ":00"
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
        'end_date': end_date, 'leave_type': leave_type, 'filename': filename, 'status': 'Pending'
    })
    save_leave_requests(LEAVE_REQUESTS)
    flash('Aapki leave request successfully submit ho gayi hai!', 'success')
    return redirect(url_for('index'))

@app.route('/update_leave/<int:req_id>/<action>')
def update_leave(req_id, action):
    if session.get('role') not in ['admin', 'developer']: return redirect(url_for('login'))
    for req in LEAVE_REQUESTS:
        if req['id'] == req_id:
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

@app.route('/export')
def export_excel():
    if session.get('role') not in ['admin', 'developer']: return redirect(url_for('login'))
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
    if session.get('role') not in ['admin', 'developer']: return redirect(url_for('login'))
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
    if session.get('role') in ['admin', 'developer'] and request.args.get('pwd') == dev_pass:
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
        existing_keys = {(str(item.get('user_id')), str(item.get('timestamp'))) for item in SYNCED_ATTENDANCE_LOGS}
        added_count = 0
        for log in logs:
            key = (str(log.get('user_id')), str(log.get('timestamp')))
            if key not in existing_keys:
                SYNCED_ATTENDANCE_LOGS.append(log)
                existing_keys.add(key)
                added_count += 1
        LAST_DEVICE_SYNC_TIME = datetime.now()
        return {'status': 'success', 'message': f'{len(logs)} records synced ({added_count} new)'}, 200
    except Exception as e: return {'status': 'error', 'message': str(e)}, 500

if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)

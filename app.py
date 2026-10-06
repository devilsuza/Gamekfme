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
MACHINES_FILE = 'machines.json'

SYNCED_ATTENDANCE_LOGS = []
LAST_DEVICE_SYNC_TIME = None

# Master Employees (Fallback / Default Setup)
MASTER_EMPLOYEES = {
    'NWC2981': {'name': 'ANTONIO JOSE BANDOLA', 'off': 'SUNDAY', 'dept': 'ADMIN - MANAGER', 'shift': 'morning'},
    'NWC3127': {'name': 'MATEUS ANTONIO DA COSTA BALMIRO', 'off': 'FRIDAY', 'dept': 'ADMIN - MANAGER', 'shift': 'morning'},
    'NWC1525': {'name': 'ETY JOSÉ BANDUA MONTEIRO', 'off': 'MONDAY', 'dept': 'ADMIN - MANAGER', 'shift': 'morning'},
    'NWC8328': {'name': 'JOAO MATIAS DOMINGOS', 'off': 'SATURDAY', 'dept': 'ADMIN - CCTV', 'shift': 'morning'},
    'NWC6661': {'name': 'FRANCISCO MUNDELE CHIVELA', 'off': 'WEDNESDAY', 'dept': 'ADMIN - CCTV', 'shift': 'morning'}
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

def load_machines():
    machines = load_json_file(MACHINES_FILE)
    if not machines:
        machines = {
            'LM11': {'ip': os.getenv('MACHINE_IP', '192.168.1.153'), 'port': int(os.getenv('MACHINE_PORT', 4370)), 'name': 'Gamek HRMS', 'admin': 'NWC1234'},
            'LF07': {'ip': '192.168.88.16', 'port': 4370, 'name': 'Benfica LF07 HRMS', 'admin': 'LF07_ADMIN'},
            'DEV': {'ip': '127.0.0.1', 'port': 4370, 'name': 'Fresmart', 'admin': 'NCSA0608'}
        }
        save_json_file(MACHINES_FILE, machines)
    return machines

# Database initialization
def load_users_db():
    db = load_json_file(USERS_DB_FILE)
    if not db:
        db = {}
        for k, v in MASTER_EMPLOYEES.items():
            db[k] = {'name': v['name'], 'password': '123', 'allowed_stores': ['LM11'], 'status': 'active', 'role': 'employee', 'designation': v.get('dept', 'General'), 'off': v.get('off', 'SUNDAY'), 'dept': v.get('dept', 'General'), 'shift': v.get('shift', 'morning'), 'allowed_features': ['calendar', 'leave', 'shift', 'salary'], 'profile_pic': ''}
        db['LM11'] = {'name': 'Admin (LM11)', 'password': os.getenv('ADMIN_PWD', 'Gamek@789'), 'allowed_stores': ['LM11'], 'status': 'active', 'role': 'admin', 'designation': 'Store Admin', 'off': 'SUNDAY', 'dept': 'Admin', 'shift': 'morning', 'allowed_features': ['roster_planner', 'shift_approvals', 'password_resets', 'calendar', 'leave', 'salary', 'roster'], 'profile_pic': ''}
        save_json_file(USERS_DB_FILE, db)
    
    # Migration for legacy DB to new features format
    modified = False
    for k, v in db.items():
        if 'allowed_stores' not in v:
            v['allowed_stores'] = [v.get('store', 'LM11')]
            modified = True
        if 'allowed_features' not in v:
            if v.get('role') == 'admin':
                v['allowed_features'] = ['roster_planner', 'shift_approvals', 'password_resets', 'calendar', 'leave', 'salary', 'roster']
            else:
                v['allowed_features'] = ['calendar', 'leave', 'shift', 'salary']
            modified = True
        if 'designation' not in v:
            v['designation'] = v.get('dept', 'Employee')
            modified = True
        if 'profile_pic' not in v:
            v['profile_pic'] = ''
            modified = True
    if modified:
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

def get_emp_info(emp_code):
    emp_str = str(emp_code).strip()
    if not emp_str.startswith('NWC') and f"NWC{emp_str}" in load_users_db(): emp_str = f"NWC{emp_str}"
    
    db = load_users_db()
    if emp_str in db: return db[emp_str]
    return {'name': f'Employee {emp_code}', 'off': 'SUNDAY', 'dept': 'General', 'shift': 'morning'}

def check_device_connectivity(store_code):
    machines = load_machines()
    machine = machines.get(store_code, machines.get('LM11', {'ip': '127.0.0.1', 'port': 4370}))
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
    machines = load_machines()
    machine = machines.get(store_code, machines.get('LM11', {'ip': '127.0.0.1', 'port': 4370}))
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
                
            if store_code != 'DEV' and store_code not in db.get(emp_code, {}).get('allowed_stores', []):
                continue

            if filter_user_id and filter_user_id != 'ALL' and emp_code != filter_user_id and raw_uid != filter_user_id:
                continue
            
            raw_punches_list.append({'date': att_date_str, 'time': att_ts.strftime('%H:%M:%S'), 'user_id': emp_code, 'name': emp_name, 'timestamp': att_ts})
            if att_date_str not in period_data: period_data[att_date_str] = {}
            if emp_code not in period_data[att_date_str]: period_data[att_date_str][emp_code] = {'name': emp_name, 'timestamps': []}
            period_data[att_date_str][emp_code]['timestamps'].append(att_ts)
            
    users_list = []
    for k, v in db.items():
        if store_code == 'DEV' or store_code in v.get('allowed_stores', []):
            users_list.append({'user_id': k, 'name': v['name'], 'dept': v.get('dept', ''), 'off': v.get('off', '')})
    users_list = sorted(users_list, key=lambda x: x['name'])

    final_data = []
    total_duration_seconds, total_lunch_seconds, total_net_variance_seconds = 0, 0, 0
    present_count, absent_count, off_count, mis_punch_count, late_arrival_count, ml_count, shift_a_count, shift_b_count = 0, 0, 0, 0, 0, 0, 0, 0
    
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
            if store_code != 'DEV' and store_code not in emp_info.get('allowed_stores', []): continue
            
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
</head>
<body class="bg-slate-900 min-h-screen flex items-center justify-center p-4">
    <div class="bg-white/95 rounded-3xl shadow-2xl p-8 w-full max-w-md space-y-6">
        <h1 class="text-2xl font-black text-center text-slate-900">Attendance Portal</h1>
        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}{% for category, message in messages %}<div class="bg-blue-50 text-blue-800 border p-3 rounded-xl text-center">{{ message }}</div>{% endfor %}{% endif %}
        {% endwith %}
        {% if error %}<div class="bg-rose-50 text-rose-700 p-3 rounded-xl text-center">{{ error }}</div>{% endif %}
        <form method="POST" action="/login" class="space-y-4">
            <input type="text" name="user_id" required placeholder="User ID / Store Code" class="w-full bg-slate-50 border rounded-xl px-4 py-3">
            <input type="password" name="password" required placeholder="Password" class="w-full bg-slate-50 border rounded-xl px-4 py-3">
            <button type="submit" class="w-full bg-emerald-600 text-white font-bold py-3 rounded-xl">Login 🚀</button>
            <div class="text-center mt-3"><a href="/reset_password" class="text-xs text-emerald-600">Forgot Password?</a></div>
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
    <title>{{ portal_name }} | Portal</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script>
        function toggleModal(id, show) {
            let el = document.getElementById(id);
            if(el) { show ? el.classList.remove('hidden') : el.classList.add('hidden'); }
        }
    </script>
</head>
<body class="bg-slate-50 text-slate-800 flex h-screen overflow-hidden">
    <!-- Sidebar Navigation -->
    <aside id="sidebar" class="w-64 bg-white border-r border-slate-200 flex flex-col justify-between p-4 space-y-2">
        <div>
            <h2 class="text-sm font-bold text-slate-900 mb-4">{{ portal_name }}</h2>
            <a href="/" class="block px-3 py-2 rounded bg-slate-900 text-white text-xs">Dashboard</a>
            
            {% if 'roster_planner' in user_features %}
            <a href="#" onclick="toggleModal('roster-planner-modal', true)" class="block px-3 py-2 rounded text-slate-600 hover:bg-slate-100 text-xs">Roster Planner</a>
            {% endif %}
            
            {% if 'shift_approvals' in user_features or role == 'developer' %}
            <a href="#" onclick="toggleModal('shift-approvals-modal', true)" class="block px-3 py-2 rounded text-slate-600 hover:bg-slate-100 text-xs">Shift Approvals</a>
            {% endif %}
            
            {% if role == 'developer' %}
            <a href="#" onclick="toggleModal('machine-mgmt-modal', true)" class="block px-3 py-2 rounded text-slate-600 hover:bg-slate-100 text-xs">Machine Management</a>
            <a href="#" onclick="toggleModal('user-mgmt-modal', true)" class="block px-3 py-2 rounded text-slate-600 hover:bg-slate-100 text-xs">User Management</a>
            {% endif %}
            
            {% if 'leave' in user_features %}
            <a href="#" onclick="toggleModal('leave-modal', true)" class="block px-3 py-2 rounded text-slate-600 hover:bg-slate-100 text-xs">Leave Management</a>
            {% endif %}
        </div>
    </aside>

    <div class="flex-1 flex flex-col h-screen overflow-hidden">
        <!-- Top Navbar -->
        <header class="bg-white border-b px-6 py-3 flex justify-between items-center z-10">
            <h1 class="text-base font-black text-slate-900">Dashboard</h1>
            <div class="flex items-center gap-3">
                <form action="/switch_store" method="GET" class="flex items-center gap-2">
                    <select name="store_code" onchange="this.form.submit()" class="text-xs border rounded p-1">
                        {% for st in allowed_stores %}
                        <option value="{{ st }}" {% if current_store == st %}selected{% endif %}>{{ st }}</option>
                        {% endfor %}
                    </select>
                </form>
                
                <div class="flex items-center space-x-2 bg-slate-100 px-3 py-1.5 rounded-xl cursor-pointer" onclick="toggleModal('profile-modal', true)">
                    {% if profile_pic %}
                    <img src="/uploads/{{ profile_pic }}" class="w-6 h-6 rounded-full object-cover">
                    {% else %}
                    <span class="w-6 h-6 rounded-full bg-emerald-500 text-white flex items-center justify-center text-xs">👤</span>
                    {% endif %}
                    <span class="text-xs font-bold">{{ logged_user_name }} ({{ user_designation }})</span>
                </div>
                <a href="/logout" class="text-rose-600 text-xs font-bold">Logout 🔒</a>
            </div>
        </header>

        <main class="flex-1 overflow-y-auto p-6 space-y-6">
            {% with messages = get_flashed_messages(with_categories=true) %}
                {% if messages %}
                    {% for category, message in messages %}
                    <div class="bg-emerald-50 text-emerald-800 border p-4 rounded text-xs font-bold">{{ message }}</div>
                    {% endfor %}
                {% endif %}
            {% endwith %}

            <div class="bg-white rounded-2xl shadow-sm border p-5">
                <form method="GET" action="/" class="flex gap-4 items-end">
                    <input type="date" name="start_date" value="{{ start_date }}" class="border rounded px-3 py-2 text-sm">
                    <input type="date" name="end_date" value="{{ end_date }}" class="border rounded px-3 py-2 text-sm">
                    <select name="employee" class="border rounded px-3 py-2 text-sm">
                        <option value="ALL">-- All --</option>
                        {% for emp in all_users %}
                            <option value="{{ emp.user_id }}" {% if selected_emp == emp.user_id %}selected{% endif %}>{{ emp.name }}</option>
                        {% endfor %}
                    </select>
                    <button type="submit" class="bg-slate-900 text-white px-4 py-2 rounded text-sm">Filter</button>
                </form>
            </div>
            
            <div class="bg-white rounded-2xl shadow border overflow-x-auto">
                <table class="w-full text-left text-xs">
                    <thead class="bg-slate-100 text-slate-600 uppercase font-bold">
                        <tr>
                            <th class="py-3 px-4">Date</th>
                            <th class="py-3 px-4">ID</th>
                            <th class="py-3 px-4">Name</th>
                            <th class="py-3 px-4">Store In</th>
                            <th class="py-3 px-4">Out Time</th>
                            <th class="py-3 px-4">Total Hrs</th>
                            <th class="py-3 px-4">Status</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for log in logs %}
                        <tr class="border-b">
                            <td class="py-2 px-4">{{ log.date }}</td>
                            <td class="py-2 px-4">{{ log.user_id }}</td>
                            <td class="py-2 px-4">{{ log.name }}</td>
                            <td class="py-2 px-4">{{ log.store_in }}</td>
                            <td class="py-2 px-4">{{ log.out_time }}</td>
                            <td class="py-2 px-4">{{ log.total_hours }}</td>
                            <td class="py-2 px-4">{{ log.status }}</td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </main>
    </div>

    <!-- Profile Modal -->
    <div id="profile-modal" class="fixed inset-0 bg-slate-900/60 flex items-center justify-center hidden z-50">
        <div class="bg-white p-5 rounded-2xl w-96 space-y-4">
            <h3 class="font-bold text-lg">My Profile</h3>
            <form action="/upload_profile_pic" method="POST" enctype="multipart/form-data" class="space-y-3">
                <input type="file" name="profile_pic" accept="image/*" class="w-full text-sm">
                <div class="flex gap-2">
                    <button type="submit" class="bg-emerald-600 text-white px-3 py-1.5 rounded text-xs">Upload Photo</button>
                    <button type="button" onclick="toggleModal('profile-modal', false)" class="bg-slate-200 px-3 py-1.5 rounded text-xs">Close</button>
                </div>
            </form>
        </div>
    </div>

    <!-- Machine Management Modal (Dev) -->
    <div id="machine-mgmt-modal" class="fixed inset-0 bg-slate-900/60 flex items-center justify-center hidden z-50">
        <div class="bg-white p-5 rounded-2xl w-full max-w-2xl space-y-4">
            <div class="flex justify-between items-center border-b pb-2">
                <h3 class="font-bold text-lg">⚙️ Machine Management</h3>
                <button onclick="toggleModal('machine-mgmt-modal', false)" class="text-slate-400 font-bold">✕</button>
            </div>
            <form action="/manage_machine" method="POST" class="flex gap-2 items-end bg-slate-50 p-3 rounded">
                <input type="text" name="store_code" placeholder="Store Code (e.g. LM11)" required class="border p-1.5 text-xs w-24">
                <input type="text" name="name" placeholder="Name" required class="border p-1.5 text-xs w-32">
                <input type="text" name="ip" placeholder="IP Address" required class="border p-1.5 text-xs w-32">
                <input type="number" name="port" value="4370" required class="border p-1.5 text-xs w-20">
                <input type="text" name="admin" placeholder="Admin ID" required class="border p-1.5 text-xs w-24">
                <button type="submit" class="bg-emerald-600 text-white px-3 py-1.5 rounded text-xs font-bold">Add/Update</button>
            </form>
            <div class="overflow-auto h-64 border rounded">
                <table class="w-full text-xs text-left">
                    <thead class="bg-slate-100 font-bold"><tr><th class="p-2">Store Code</th><th class="p-2">Name</th><th class="p-2">IP</th><th class="p-2">Port</th><th class="p-2">Admin ID</th></tr></thead>
                    <tbody>
                        {% for mcode, mdata in all_machines.items() %}
                        <tr class="border-b"><td class="p-2 font-bold">{{ mcode }}</td><td class="p-2">{{ mdata.name }}</td><td class="p-2">{{ mdata.ip }}</td><td class="p-2">{{ mdata.port }}</td><td class="p-2">{{ mdata.admin }}</td></tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- User Management Modal (Dev) -->
    <div id="user-mgmt-modal" class="fixed inset-0 bg-slate-900/60 flex items-center justify-center hidden z-50 p-4">
        <div class="bg-white p-5 rounded-2xl w-full max-w-4xl space-y-4 max-h-[90vh] flex flex-col">
            <div class="flex justify-between items-center border-b pb-2">
                <h3 class="font-bold text-lg">🪪 User Management</h3>
                <button onclick="toggleModal('user-mgmt-modal', false)" class="text-slate-400 font-bold">✕</button>
            </div>
            <form action="/manage_user" method="POST" class="bg-slate-50 p-4 rounded space-y-3">
                <input type="hidden" name="action" value="create">
                <div class="flex gap-2">
                    <input type="text" name="uid" placeholder="ID (NWC...)" required class="border p-2 text-xs flex-1">
                    <input type="text" name="name" placeholder="Full Name" required class="border p-2 text-xs flex-1">
                    <input type="text" name="password" value="123" required class="border p-2 text-xs flex-1">
                    <input type="text" name="designation" placeholder="Designation" class="border p-2 text-xs flex-1">
                    <select name="role" class="border p-2 text-xs flex-1">
                        <option value="employee">Employee</option>
                        <option value="admin">Admin</option>
                    </select>
                </div>
                <div class="flex flex-col gap-1">
                    <label class="text-[10px] font-bold text-slate-500">Allowed Stores (Ctrl+Click for multiple)</label>
                    <select name="allowed_stores" multiple class="border p-2 text-xs h-20">
                        {% for mcode in all_machines.keys() %}<option value="{{ mcode }}">{{ mcode }}</option>{% endfor %}
                    </select>
                </div>
                <button type="submit" class="bg-emerald-600 text-white px-4 py-2 rounded text-xs font-bold w-full">Add / Update User</button>
            </form>
            <div class="overflow-y-auto border rounded flex-1">
                <table class="w-full text-xs text-left">
                    <thead class="bg-slate-100 font-bold"><tr><th class="p-2">ID</th><th class="p-2">Name</th><th class="p-2">Role/Desig</th><th class="p-2">Stores</th><th class="p-2">Action</th></tr></thead>
                    <tbody>
                        {% for u_id, u_info in users_db.items() %}
                        <tr class="border-b">
                            <td class="p-2 font-bold">{{ u_id }}</td><td class="p-2">{{ u_info.name }}</td><td class="p-2">{{ u_info.role }} / {{ u_info.designation }}</td>
                            <td class="p-2">{{ u_info.allowed_stores|join(', ') }}</td>
                            <td class="p-2">
                                <form action="/manage_user" method="POST" class="inline"><input type="hidden" name="uid" value="{{ u_id }}"><input type="hidden" name="action" value="toggle_status"><button class="bg-amber-100 text-amber-700 px-2 py-1 rounded text-[10px]">Toggle</button></form>
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
    if request.method == 'POST':
        uid = request.form.get('user_id').strip().upper()
        pwd = request.form.get('password').strip()
        
        dev_pass = os.getenv('DEV_PWD', 'Shama@8577')
        if uid == 'NCSA0608' and pwd == dev_pass:
            session.update({'logged_in': True, 'role': 'developer', 'user_id': uid, 'user_name': 'Sonu Kumar', 'current_store': 'DEV'})
            return redirect(url_for('index'))
            
        db = load_users_db()
        emp_key = f"NWC{uid}" if not uid.startswith('NWC') and uid not in ['LM11', 'LF07'] else uid
        
        if emp_key in db and pwd == db[emp_key].get('password'):
            user_data = db[emp_key]
            if user_data.get('status') == 'blocked': return render_template_string(LOGIN_TEMPLATE, error="Blocked!")
            session.update({
                'logged_in': True, 'role': user_data['role'], 'user_id': emp_key, 
                'user_name': user_data['name'], 'current_store': user_data.get('allowed_stores', ['LM11'])[0]
            })
            return redirect(url_for('index'))
        return render_template_string(LOGIN_TEMPLATE, error="Invalid Credentials!")
    return render_template_string(LOGIN_TEMPLATE)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/switch_store')
def switch_store():
    store = request.args.get('store_code')
    db = load_users_db()
    uid = session.get('user_id')
    
    if session.get('role') == 'developer' or store in db.get(uid, {}).get('allowed_stores', []):
        session['current_store'] = store
    return redirect(url_for('index'))

@app.route('/')
def index():
    if not session.get('logged_in'): return redirect(url_for('login'))
        
    role = session.get('role')
    uid = session.get('user_id')
    current_store = session.get('current_store', 'LM11')
    
    db = load_users_db()
    user_data = db.get(uid, {})
    
    allowed_stores = user_data.get('allowed_stores', ['LM11']) if role != 'developer' else list(load_machines().keys())
    user_features = user_data.get('allowed_features', []) if role != 'developer' else ['roster_planner', 'shift_approvals', 'password_resets', 'calendar', 'leave', 'salary', 'roster']
    profile_pic = user_data.get('profile_pic', '')
    user_designation = user_data.get('designation', 'Developer' if role == 'developer' else 'Employee')

    start_date = request.args.get('start_date', datetime.now().strftime('%Y-%m-%d'))
    end_date = request.args.get('end_date', datetime.now().strftime('%Y-%m-%d'))
    selected_emp = request.args.get('employee', 'ALL') if role != 'employee' else uid
    
    machines = load_machines()
    portal_name = machines.get(current_store, {}).get('name', 'HRMS Portal')
    
    logs, all_users, g_hrs, g_l_hrs, g_var, raw_punches, stats = fetch_attendance_data(start_date, end_date, selected_emp, current_store)
    
    return render_template_string(
        HTML_TEMPLATE, logs=logs, all_users=all_users, start_date=start_date, end_date=end_date, 
        selected_emp=selected_emp, role=role, logged_user_name=session.get('user_name'), portal_name=portal_name,
        current_store=current_store, allowed_stores=allowed_stores, user_features=user_features,
        users_db=db, all_machines=machines, profile_pic=profile_pic, user_designation=user_designation
    )

@app.route('/manage_machine', methods=['POST'])
def manage_machine():
    if session.get('role') != 'developer': return redirect(url_for('index'))
    machines = load_machines()
    m_code = request.form.get('store_code').strip().upper()
    machines[m_code] = {
        'name': request.form.get('name').strip(),
        'ip': request.form.get('ip').strip(),
        'port': int(request.form.get('port')),
        'admin': request.form.get('admin').strip().upper()
    }
    save_json_file(MACHINES_FILE, machines)
    flash(f"Machine {m_code} added/updated successfully.", "success")
    return redirect(url_for('index'))

@app.route('/manage_user', methods=['POST'])
def manage_user():
    if session.get('role') != 'developer': return redirect(url_for('index'))
    db = load_users_db()
    action = request.form.get('action')
    uid = request.form.get('uid').strip().upper()
    uid = f"NWC{uid}" if not uid.startswith('NWC') and uid not in load_machines().keys() else uid
    
    if action == 'create':
        db[uid] = {
            'name': request.form.get('name').strip(),
            'password': request.form.get('password').strip(),
            'role': request.form.get('role'),
            'designation': request.form.get('designation').strip(),
            'allowed_stores': request.form.getlist('allowed_stores'),
            'status': 'active', 'off': 'SUNDAY', 'shift': 'morning', 'profile_pic': '',
            'allowed_features': ['calendar', 'leave', 'shift', 'salary'] if request.form.get('role') == 'employee' else ['roster_planner', 'shift_approvals', 'password_resets', 'calendar', 'leave', 'salary', 'roster']
        }
        flash(f"User {uid} created/updated successfully.", "success")
    elif action == 'toggle_status' and uid in db:
        db[uid]['status'] = 'blocked' if db[uid].get('status') == 'active' else 'active'
        flash(f"User {uid} status changed.", "success")
        
    save_json_file(USERS_DB_FILE, db)
    return redirect(url_for('index'))

@app.route('/upload_profile_pic', methods=['POST'])
def upload_profile_pic():
    uid = session.get('user_id')
    file = request.files.get('profile_pic')
    if file and file.filename != '' and allowed_file(file.filename):
        filename = f"profile_{uid}_{secure_filename(file.filename)}"
        file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        
        db = load_users_db()
        if uid in db:
            db[uid]['profile_pic'] = filename
            save_json_file(USERS_DB_FILE, db)
            flash("Profile picture updated!", "success")
    return redirect(url_for('index'))

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)

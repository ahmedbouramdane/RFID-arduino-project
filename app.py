import serial
import serial.tools.list_ports
import json
import threading
import os
import time
from datetime import datetime, timedelta
import pandas as pd
from flask import Flask, render_template, request, jsonify
from flask_socketio import SocketIO

app = Flask(__name__)
socketio = SocketIO(app, cors_allowed_origins="*", async_mode='threading')

EXCEL_FILE = 'attendance.xlsx'
JSON_FILE = 'users.json'
AVATAR_DIR = 'static/avatars'
file_lock = threading.Lock()

REQUIRED_COLUMNS = ['UID', 'Full Name', 'Role', 'Date', 'Entry Time', 'Exit Time', 'Duration (Min)', 'Status']

def init_db():
    os.makedirs(AVATAR_DIR, exist_ok=True)
    if not os.path.exists(JSON_FILE):
        with open(JSON_FILE, 'w') as f:
            json.dump({}, f)

def ensure_excel_file():
    if not os.path.exists(EXCEL_FILE):
        df = pd.DataFrame(columns=REQUIRED_COLUMNS)
        df.to_excel(EXCEL_FILE, index=False)
        return
    try:
        df = pd.read_excel(EXCEL_FILE)
        missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
        if missing:
            print(f"⚠️ Excel file missing columns: {missing}. Recreating...")
            df = pd.DataFrame(columns=REQUIRED_COLUMNS)
            df.to_excel(EXCEL_FILE, index=False)
    except Exception as e:
        print(f"⚠️ Error reading Excel file: {e}. Recreating...")
        df = pd.DataFrame(columns=REQUIRED_COLUMNS)
        df.to_excel(EXCEL_FILE, index=False)

init_db()

def format_duration(seconds):
    if seconds < 0 or seconds == 0:
        return "0 seconds"
    units = [
        ('month', 30 * 24 * 3600),
        ('day', 24 * 3600),
        ('hour', 3600),
        ('minute', 60),
        ('second', 1)
    ]
    parts = []
    for name, unit in units:
        if seconds >= unit:
            count = int(seconds // unit)
            seconds %= unit
            parts.append(f"{count} {name}{'s' if count > 1 else ''}")
    return ", ".join(parts)

def process_scan(uid):
    with open(JSON_FILE, 'r') as f:
        users = json.load(f)
    if uid not in users:
        print(f"⚠️ Unknown Card: {uid}")
        socketio.emit('access_denied', {'uid': uid})
        return
    user = users[uid]
    with file_lock:
        ensure_excel_file()
        df = pd.read_excel(EXCEL_FILE)
        active_session = df[(df['UID'] == uid) & (df['Exit Time'].isna())]
    if active_session.empty:
        log_entry(uid, user)
    else:
        log_exit(uid, active_session.index[-1])

def log_entry(uid, user):
    with file_lock:
        ensure_excel_file()
        df = pd.read_excel(EXCEL_FILE)
        now = datetime.now()
        date_today = now.strftime('%Y-%m-%d')
        time_now = now.strftime('%H:%M:%S')
        new_row = {
            'UID': uid, 'Full Name': user['name'], 'Role': user['role'],
            'Date': date_today, 'Entry Time': time_now, 'Exit Time': None,
            'Duration (Min)': 0, 'Status': 'Auto Entry'
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
    socketio.emit('access_granted', {
        'user': user, 'type': 'ENTRY', 'time': time_now,
        'date': date_today, 'duration': 0
    })

def log_exit(uid, row_index):
    with file_lock:
        ensure_excel_file()
        df = pd.read_excel(EXCEL_FILE)
        now = datetime.now()
        time_now = now.strftime('%H:%M:%S')
        date_today = now.strftime('%Y-%m-%d')
        entry_date = str(df.at[row_index, 'Date'])
        entry_time = str(df.at[row_index, 'Entry Time'])
        entry_datetime = datetime.strptime(f"{entry_date} {entry_time}", '%Y-%m-%d %H:%M:%S')
        delta = now - entry_datetime
        total_seconds = delta.total_seconds()
        duration_minutes = total_seconds / 60.0
        df.at[row_index, 'Exit Time'] = time_now
        df.at[row_index, 'Duration (Min)'] = round(duration_minutes, 2)
        df.to_excel(EXCEL_FILE, index=False)
        users = json.load(open(JSON_FILE))
        formatted_duration = format_duration(total_seconds)
        socketio.emit('access_granted', {
            'user': users[uid], 'type': 'EXIT', 'time': time_now,
            'date': date_today, 'duration': formatted_duration
        })

@app.route('/')
def home():
    with file_lock:
        ensure_excel_file()
        df = pd.read_excel(EXCEL_FILE)
        today = datetime.now().strftime('%Y-%m-%d')
        date_str = request.args.get('date', today)   # use param or today
        df_selected = df[df['Date'] == date_str].copy()
    with open(JSON_FILE, 'r') as f:
        users = json.load(f)
    records = []
    for _, row in df_selected.iterrows():
        uid = row['UID']
        user = users.get(uid, {})
        name = user.get('name', row.get('Full Name', 'Unknown'))
        role = user.get('role', row.get('Role', ''))
        image = user.get('image', '/static/default-avatar.png')
        entry_time = row['Entry Time']
        exit_time = row['Exit Time'] if pd.notna(row['Exit Time']) else None
        status = 'Inside' if exit_time is None else 'Exited'
        if exit_time is not None:
            entry_date = str(row['Date'])
            entry_dt = datetime.strptime(f"{entry_date} {entry_time}", '%Y-%m-%d %H:%M:%S')
            exit_candidate = datetime.strptime(f"{entry_date} {exit_time}", '%Y-%m-%d %H:%M:%S')
            if exit_candidate <= entry_dt:
                while exit_candidate <= entry_dt:
                    exit_candidate += timedelta(days=1)
            exit_dt = exit_candidate
            total_seconds = (exit_dt - entry_dt).total_seconds()
            duration_str = format_duration(total_seconds)
        else:
            duration_str = "0 seconds"
        records.append({
            'uid': uid, 'name': name, 'role': role, 'image': image,
            'entry_time': entry_time, 'exit_time': exit_time,
            'duration': duration_str, 'status': status
        })
    total_users = len(users)
    total_selected = len(df_selected)
    active_now = len([r for r in records if r['status'] == 'Inside'])
    return render_template('index.html',
                           total_users=total_users,
                           total_today=total_selected,
                           active_now=active_now,
                           records=records,
                           selected_date=date_str)

@app.route('/scan')
def scan_page():
    return render_template('scan.html')

@app.route('/add_employee')
def add_employee():
    uid = request.args.get('uid', '')
    return render_template('add_employee.html', scanned_uid=uid)

@app.route('/api/add', methods=['POST'])
def add_user():
    uid = request.form.get('uid')
    name = request.form.get('name')
    role = request.form.get('role')
    image_b64 = request.form.get('image_b64')
    if not image_b64:
        import hashlib
        hash_obj = hashlib.md5(name.encode())
        color = int(hash_obj.hexdigest()[:6], 16) & 0xffffff
        image_b64 = f"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='100' height='100' viewBox='0 0 100 100'%3E%3Crect width='100' height='100' fill='%23{color:06x}'/%3E%3Ctext x='50' y='50' font-size='40' text-anchor='middle' dy='.3em' fill='white' font-family='Arial'%3E{name[0].upper()}%3C/text%3E%3C/svg%3E"
    with file_lock:
        with open(JSON_FILE, 'r') as f:
            users = json.load(f)
        users[uid] = {'name': name, 'role': role, 'image': image_b64}
        with open(JSON_FILE, 'w') as f:
            json.dump(users, f, indent=2)
    return jsonify({'success': True})

def select_serial_port():
    """List available COM ports and let the user choose one."""
    ports = list(serial.tools.list_ports.comports())
    if not ports:
        print("❌ No serial ports found. Using default 'COM11'.")
        return 'COM11'
    print("\n🔌 Available serial ports:")
    for i, port in enumerate(ports):
        print(f"  {i+1}. {port.device} - {port.description}")
    while True:
        try:
            choice = input(f"Select port (1-{len(ports)}), or press Enter for default 'COM11': ")
            if choice == "":
                return 'COM11'
            idx = int(choice) - 1
            if 0 <= idx < len(ports):
                return ports[idx].device
            else:
                print("Invalid choice. Try again.")
        except ValueError:
            print("Please enter a number.")

def serial_thread(port):
    while True:
        try:
            ser = serial.Serial(port, 9600, timeout=0.1)
            print(f"🔌 Connected to Scanner on {port}")
            while True:
                try:
                    if ser.in_waiting > 0:
                        uid = ser.readline().decode('utf-8').strip().upper()
                        if len(uid) > 4:
                            print(f"📡 Tag Detected: {uid}")
                            process_scan(uid)
                            time.sleep(2)
                    time.sleep(0.1)
                except Exception as e:
                    print(f"⚠️ Error during scan processing: {e}")
                    time.sleep(1)
        except Exception as e:
            print(f"❌ Serial connection failed on {port}: {e}. Retrying in 5 seconds...")
            time.sleep(5)

if __name__ == '__main__':
    # Interactive port selection
    selected_port = select_serial_port()
    threading.Thread(target=serial_thread, args=(selected_port,), daemon=True).start()
    socketio.run(app, debug=True, port=5000, use_reloader=False)
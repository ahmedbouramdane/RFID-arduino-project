import serial
import json
import threading
import os
import time
from datetime import datetime
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
    """Create or repair the Excel file if needed."""
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

def process_scan(uid):
    """Handle an RFID scan."""
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
        # No active session → auto entry
        log_entry(uid, user)
    else:
        # Active session → exit
        log_exit(uid, active_session.index[-1])

def log_entry(uid, user):
    with file_lock:
        ensure_excel_file()
        df = pd.read_excel(EXCEL_FILE)
        now = datetime.now()
        date_today = now.strftime('%Y-%m-%d')
        time_now = now.strftime('%H:%M:%S')

        new_row = {
            'UID': uid,
            'Full Name': user['name'],
            'Role': user['role'],
            'Date': date_today,
            'Entry Time': time_now,
            'Exit Time': None,
            'Duration (Min)': 0,
            'Status': 'Auto Entry'
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)

    socketio.emit('access_granted', {
        'user': user,
        'type': 'ENTRY',
        'time': time_now,
        'date': date_today,
        'duration': 0
    })

def log_exit(uid, row_index):
    with file_lock:
        ensure_excel_file()
        df = pd.read_excel(EXCEL_FILE)
        now = datetime.now()
        time_now = now.strftime('%H:%M:%S')
        date_today = now.strftime('%Y-%m-%d')

        start_t = datetime.strptime(str(df.at[row_index, 'Entry Time']), '%H:%M:%S')
        duration = round((now - now.replace(hour=start_t.hour, minute=start_t.minute, second=start_t.second)).total_seconds() / 60, 2)

        df.at[row_index, 'Exit Time'] = time_now
        df.at[row_index, 'Duration (Min)'] = duration
        df.to_excel(EXCEL_FILE, index=False)

        users = json.load(open(JSON_FILE))
        socketio.emit('access_granted', {
            'user': users[uid],
            'type': 'EXIT',
            'time': time_now,
            'date': date_today,
            'duration': duration
        })

@app.route('/')
def home():
    with file_lock:
        ensure_excel_file()
        df = pd.read_excel(EXCEL_FILE)
        today = datetime.now().strftime('%Y-%m-%d')
        date_str = request.args.get('date', None)
        df_today = ""
        if date_str:
            df_today = df[df['Date'] == date_str].copy() 
        else:
            df_today = df[df['Date'] == today].copy() 
        
    with open(JSON_FILE, 'r') as f:
        users = json.load(f)

    records = []
    for _, row in df_today.iterrows():
        uid = row['UID']
        user = users.get(uid, {})
        name = user.get('name', row.get('Full Name', 'Unknown'))
        role = user.get('role', row.get('Role', ''))
        image = user.get('image', '/static/default-avatar.png')

        entry_time = row['Entry Time']
        exit_time = row['Exit Time'] if pd.notna(row['Exit Time']) else None
        duration = row['Duration (Min)'] if pd.notna(row['Duration (Min)']) else 0
        status = 'Inside' if exit_time is None else 'Exited'

        records.append({
            'uid': uid,
            'name': name,
            'role': role,
            'image': image,
            'entry_time': entry_time,
            'exit_time': exit_time,
            'duration': duration,
            'status': status
        })

    total_users = len(users)
    total_today = len(df_today)
    active_now = len([r for r in records if r['status'] == 'Inside'])

    return render_template('index.html',
                           total_users=total_users,
                           total_today=total_today,
                           active_now=active_now,
                           records=records,
                           today=today)

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

    # Generate default avatar if none provided
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

# --- Robust serial thread ---
def serial_thread():
    while True:
        try:
            ser = serial.Serial('COM11', 9600, timeout=0.1)
            print("🔌 Connected to Scanner on COM11")
            while True:
                try:
                    if ser.in_waiting > 0:
                        uid = ser.readline().decode('utf-8').strip().upper()
                        if len(uid) > 4:
                            print(f"📡 Tag Detected: {uid}")
                            process_scan(uid)
                            time.sleep(2)  # debounce
                    time.sleep(0.1)
                except Exception as e:
                    print(f"⚠️ Error during scan processing: {e}")
                    time.sleep(1)
        except Exception as e:
            print(f"❌ Serial connection failed: {e}. Retrying in 5 seconds...")
            time.sleep(5)

if __name__ == '__main__':
    threading.Thread(target=serial_thread, daemon=True).start()
    socketio.run(app, debug=True, port=5000, use_reloader=False)
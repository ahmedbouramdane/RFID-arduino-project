import serial
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

def format_duration(seconds):
    """
    Convert seconds to a human-readable string like:
    "2 months, 5 days, 3 hours, 15 minutes, 30 seconds"
    Omitted if zero.
    """
    if seconds < 0:
        return "0 seconds"
    if seconds == 0:
        return "0 seconds"

    # Define units in seconds
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

        # Get entry date and time from the row
        entry_date = str(df.at[row_index, 'Date'])
        entry_time = str(df.at[row_index, 'Entry Time'])
        entry_datetime_str = f"{entry_date} {entry_time}"
        entry_datetime = datetime.strptime(entry_datetime_str, '%Y-%m-%d %H:%M:%S')

        # Exit datetime is now (the current datetime)
        exit_datetime = now

        # Compute total seconds
        delta = exit_datetime - entry_datetime
        total_seconds = delta.total_seconds()
        duration_minutes = total_seconds / 60.0

        # Update the row
        df.at[row_index, 'Exit Time'] = time_now
        df.at[row_index, 'Duration (Min)'] = round(duration_minutes, 2)
        df.to_excel(EXCEL_FILE, index=False)

        users = json.load(open(JSON_FILE))
        # Format duration for display (human-readable)
        formatted_duration = format_duration(total_seconds)

        socketio.emit('access_granted', {
            'user': users[uid],
            'type': 'EXIT',
            'time': time_now,
            'date': date_today,
            'duration': formatted_duration
        })

@app.route('/')
def home():
    with file_lock:
        ensure_excel_file()
        df = pd.read_excel(EXCEL_FILE)
        today = datetime.now().strftime('%Y-%m-%d')
        date_str = request.args.get('date', None)
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
        status = 'Inside' if exit_time is None else 'Exited'

        # Compute duration from entry and exit timestamps
        if exit_time is not None:
            # Build entry datetime
            entry_date = str(row['Date'])
            entry_time_str = str(entry_time)
            entry_datetime = datetime.strptime(f"{entry_date} {entry_time_str}", '%Y-%m-%d %H:%M:%S')
            # Build exit datetime (may be on a later day)
            # The exit date is not stored; we need to infer.
            # Since exit_time is recorded on the day of exit, but the row's Date is entry date.
            # If exit_time is less than entry_time, it's likely the next day.
            exit_time_str = str(exit_time)
            exit_candidate = datetime.strptime(f"{entry_date} {exit_time_str}", '%Y-%m-%d %H:%M:%S')
            if exit_candidate <= entry_datetime:
                # Assume it's the next day (or later)
                # We'll add days until it's after entry_datetime
                # In practice, one day is enough, but loop for safety
                while exit_candidate <= entry_datetime:
                    exit_candidate += timedelta(days=1)
            exit_datetime = exit_candidate
            delta = exit_datetime - entry_datetime
            total_seconds = delta.total_seconds()
            duration_str = format_duration(total_seconds)
        else:
            duration_str = "0 seconds"

        records.append({
            'uid': uid,
            'name': name,
            'role': role,
            'image': image,
            'entry_time': entry_time,
            'exit_time': exit_time,
            'duration': duration_str,
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
from flask import Flask, request, jsonify
from flask_cors import CORS
from werkzeug.security import generate_password_hash, check_password_hash
import sqlite3
import os
import re

app = Flask(__name__)
CORS(app) # Allow frontend to query backend port 5000

DB_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "registration.db")

def get_db():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    # BUG-03: Explicitly enable SQLite foreign key constraints
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn

# ═══════════════════════════════════════
# DATABASE INITIALIZATION & MIGRATIONS
# ═══════════════════════════════════════
def init_db():
    conn = get_db()
    cursor = conn.cursor()
    
    # 1. Users Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            first_name TEXT NOT NULL,
            last_name TEXT NOT NULL,
            reg_number TEXT UNIQUE NOT NULL,
            email TEXT UNIQUE NOT NULL,
            programme TEXT NOT NULL,
            year_of_study TEXT NOT NULL,
            password TEXT NOT NULL
        )
    ''')

    # 2. Courses Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS courses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT UNIQUE NOT NULL,
            title TEXT NOT NULL,
            department TEXT NOT NULL,
            lecturer TEXT NOT NULL,
            credits INTEGER NOT NULL,
            capacity INTEGER NOT NULL,
            enrolled_count INTEGER DEFAULT 0,
            day TEXT NOT NULL,
            time TEXT NOT NULL,
            slot INTEGER NOT NULL,
            prerequisite TEXT NOT NULL
        )
    ''')

    # 3. Sessions Table (BUG-02 Token Storage)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS sessions (
            user_id INTEGER PRIMARY KEY,
            token TEXT UNIQUE NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (user_id) REFERENCES users (id) ON DELETE CASCADE
        )
    ''')

    # 4. Safe Schema Migration for Enrolments Table (BUG-06: NOT NULL Constraints)
    # Check if table enrolments exists
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='enrolments'")
    table_exists = cursor.fetchone()
    
    if table_exists:
        # Check if columns are NOT NULL
        cursor.execute("PRAGMA table_info(enrolments)")
        columns = cursor.fetchall()
        user_id_nullable = True
        course_id_nullable = True
        for col in columns:
            if col[1] == 'user_id' and col[3] == 1:  # col[3] is notnull constraint (1 = NOT NULL, 0 = Nullable)
                user_id_nullable = False
            if col[1] == 'course_id' and col[3] == 1:
                course_id_nullable = False
        
        if user_id_nullable or course_id_nullable:
            # Perform safe migration without wiping data
            cursor.execute("ALTER TABLE enrolments RENAME TO enrolments_old")
            cursor.execute('''
                CREATE TABLE enrolments (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_id INTEGER NOT NULL,
                    course_id INTEGER NOT NULL,
                    enrolled_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (user_id) REFERENCES users (id) ON DELETE CASCADE,
                    FOREIGN KEY (course_id) REFERENCES courses (id) ON DELETE CASCADE,
                    UNIQUE(user_id, course_id)
                )
            ''')
            # Copy valid non-null rows
            cursor.execute('''
                INSERT INTO enrolments (id, user_id, course_id, enrolled_at)
                SELECT id, user_id, course_id, enrolled_at FROM enrolments_old
                WHERE user_id IS NOT NULL AND course_id IS NOT NULL
            ''')
            cursor.execute("DROP TABLE enrolments_old")
            conn.commit()
    else:
        # Table doesn't exist, create fresh with NOT NULL
        cursor.execute('''
            CREATE TABLE enrolments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                course_id INTEGER NOT NULL,
                enrolled_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (user_id) REFERENCES users (id) ON DELETE CASCADE,
                FOREIGN KEY (course_id) REFERENCES courses (id) ON DELETE CASCADE,
                UNIQUE(user_id, course_id)
            )
        ''')
        conn.commit()

    # Seed initial courses if empty
    cursor.execute("SELECT COUNT(*) FROM courses")
    if cursor.fetchone()[0] == 0:
        seed_courses = [
            ('CCS 3101', 'Database Management Systems', 'CS', 'Dr. Stephen Mburu', 3, 50, 42, 'Monday', '08:00 - 11:00', 0, 'CCS 2101'),
            ('CIT 3102', 'Network Administration', 'IT', 'Dr. Faith Kamau', 3, 45, 38, 'Monday', '11:00 - 14:00', 1, 'None'),
            ('CCS 3103', 'Artificial Intelligence', 'CS', 'Ms. Mercy Njagi', 3, 40, 39, 'Tuesday', '08:00 - 11:00', 0, 'CCS 2104'),
            ('CCS 3104', 'Software Engineering II', 'CS', 'Dr. George Musumba', 3, 60, 51, 'Wednesday', '08:00 - 11:00', 0, 'CCS 2105'),
            ('CIT 3105', 'Web Application Development', 'IT', 'Ms. Lydia Wanjiru', 3, 50, 48, 'Thursday', '14:00 - 17:00', 2, 'CIT 2102'),
            ('CCS 3106', 'Cybersecurity Principles', 'CS', 'Mr. James Mwangi', 3, 35, 35, 'Friday', '08:00 - 11:00', 0, 'CCS 3102'),
            ('CCS 4101', 'Human Computer Interaction', 'CS', 'Dr. Agnes Mindila', 3, 50, 25, 'Monday', '08:00 - 11:00', 0, 'None'),
            ('CIT 3108', 'Cloud Computing Infrastructure', 'IT', 'Mr. Anthony Njoroge', 3, 40, 18, 'Friday', '11:00 - 14:00', 1, 'CIT 3102'),
            ('CCS 3109', 'Operating Systems', 'CS', 'Mr. David Kimani', 3, 55, 40, 'Wednesday', '14:00 - 17:00', 2, 'CCS 2101'),
            ('CIT 3110', 'Mobile App Development', 'IT', 'Mr. Brian Ochieng', 3, 45, 30, 'Thursday', '08:00 - 11:00', 0, 'CIT 2102')
        ]
        cursor.executemany('''
            INSERT INTO courses (code, title, department, lecturer, credits, capacity, enrolled_count, day, time, slot, prerequisite)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', seed_courses)
        
    conn.commit()
    conn.close()

# ═══════════════════════════════════════
# AUTH TOKEN HELPER (BUG-02)
# ═══════════════════════════════════════
def verify_token():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return None
    token = auth_header.split(' ')[1]
    
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT user_id FROM sessions WHERE token = ?", (token,))
    row = cursor.fetchone()
    conn.close()
    
    return row['user_id'] if row else None

# ═══════════════════════════════════════
# AUTH API ENDPOINTS
# ═══════════════════════════════════════
@app.route('/api/signup', methods=['POST'])
def signup():
    # BUG-04: Server-Side Signup Validation
    data = request.json or {}
    required_fields = ['firstName', 'lastName', 'regNumber', 'email', 'programme', 'yearOfStudy', 'password']
    
    for f in required_fields:
        if f not in data or not str(data[f]).strip():
            return jsonify({"status": "error", "message": f"Field '{f}' is required and cannot be empty"}), 400
            
    email = str(data['email']).strip()
    if not re.match(r"[^@]+@[^@]+\.[^@]+", email):
        return jsonify({"status": "error", "message": "Invalid email address format"}), 400
        
    reg_number = str(data['regNumber']).strip()
    if len(reg_number) < 5:
        return jsonify({"status": "error", "message": "Invalid registration number format"}), 400
        
    password = str(data['password'])
    if len(password) < 6:
        return jsonify({"status": "error", "message": "Password must be at least 6 characters long"}), 400

    conn = get_db()
    cursor = conn.cursor()
    
    # Check duplicate email
    cursor.execute("SELECT id FROM users WHERE email = ?", (email,))
    if cursor.fetchone():
        conn.close()
        return jsonify({"status": "error", "message": "An account with this email already exists"}), 400

    # Check duplicate registration number
    cursor.execute("SELECT id FROM users WHERE reg_number = ?", (reg_number,))
    if cursor.fetchone():
        conn.close()
        return jsonify({"status": "error", "message": "An account with this registration number already exists"}), 400

    try:
        # BUG-01: Secure password hashing
        password_hash = generate_password_hash(password)
        
        cursor.execute('''
            INSERT INTO users (first_name, last_name, reg_number, email, programme, year_of_study, password)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (data['firstName'].strip(), data['lastName'].strip(), reg_number, email, data['programme'], data['yearOfStudy'], password_hash))
        conn.commit()
        return jsonify({"status": "success", "message": "User registered successfully"})
    except sqlite3.IntegrityError:
        return jsonify({"status": "error", "message": "Integrity violation: signup blocked"}), 400
    finally:
        conn.close()

@app.route('/api/login', methods=['POST'])
def login():
    data = request.json or {}
    email = data.get('email', '').strip()
    password = data.get('password', '')

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM users WHERE email = ?", (email,))
    user = cursor.fetchone()
    
    # BUG-01: Verify hashed password
    if user and check_password_hash(user["password"], password):
        # BUG-02: Generate session token
        import secrets
        token = secrets.token_hex(24)
        
        cursor.execute("INSERT OR REPLACE INTO sessions (user_id, token) VALUES (?, ?)", (user["id"], token))
        conn.commit()
        conn.close()
        
        return jsonify({
            "status": "success",
            "user": {
                "id": user["id"],
                "firstName": user["first_name"],
                "lastName": user["last_name"],
                "regNumber": user["reg_number"],
                "email": user["email"],
                "programme": user["programme"],
                "yearOfStudy": user["year_of_study"],
                "token": token
            }
        })
        
    conn.close()
    return jsonify({"status": "error", "message": "Invalid email or password"}), 401

@app.route('/api/logout', methods=['POST'])
def logout():
    auth_header = request.headers.get('Authorization')
    if auth_header and auth_header.startswith('Bearer '):
        token = auth_header.split(' ')[1]
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM sessions WHERE token = ?", (token,))
        conn.commit()
        conn.close()
    return jsonify({"status": "success", "message": "Logged out successfully"})

# ═══════════════════════════════════════
# COURSES & ENROLMENT API ENDPOINTS
# ═══════════════════════════════════════
@app.route('/api/courses', methods=['GET'])
def get_courses():
    # BUG-02: Verify token auth
    auth_user_id = verify_token()
    if not auth_user_id:
        return jsonify({"status": "error", "message": "Unauthorized access. Please login."}), 401

    req_user_id = request.args.get('user_id')
    if not req_user_id or int(req_user_id) != auth_user_id:
        return jsonify({"status": "error", "message": "Forbidden: You cannot view another student's courses."}), 403

    conn = get_db()
    cursor = conn.cursor()
    
    # Fetch all courses
    cursor.execute("SELECT * FROM courses")
    courses = [dict(row) for row in cursor.fetchall()]
    
    # Fetch user's registered course IDs
    cursor.execute("SELECT course_id FROM enrolments WHERE user_id = ?", (auth_user_id,))
    registered_ids = [row[0] for row in cursor.fetchall()]
    
    # Build list of active enrolled course codes
    cursor.execute('''
        SELECT c.code FROM enrolments e
        JOIN courses c ON e.course_id = c.id
        WHERE e.user_id = ?
    ''', (auth_user_id,))
    enrolled_codes = [row[0] for row in cursor.fetchall()]
        
    conn.close()
    
    # BUG-05: Real Prerequisite checking based on database enrolments
    # Prerequisites that are offered inside our catalog can be satisfied by active enrolment.
    # Prerequisites NOT offered inside our catalog are assumed to be satisfied from past semesters.
    offered_codes = [c['code'] for c in courses]
    
    for c in courses:
        c['registered'] = c['id'] in registered_ids
        prereq = c['prerequisite']
        
        # Map department to dept for frontend compatibility
        c['dept'] = c['department']
        
        if prereq == 'None':
            c['prereqMet'] = True
        elif prereq in offered_codes:
            # Prerequisite is offered this semester, student must be registered for it
            c['prereqMet'] = prereq in enrolled_codes
        else:
            # Prerequisite is not offered, assume satisfied in a prior semester
            c['prereqMet'] = True
            
    return jsonify({"status": "success", "courses": courses})

@app.route('/api/toggle-course', methods=['POST'])
def toggle_course():
    # BUG-02: Verify token auth
    auth_user_id = verify_token()
    if not auth_user_id:
        return jsonify({"status": "error", "message": "Unauthorized access. Please login."}), 401

    data = request.json or {}
    req_user_id = data.get('user_id')
    course_id = data.get('course_id')
    
    if not req_user_id or int(req_user_id) != auth_user_id:
        return jsonify({"status": "error", "message": "Forbidden: You cannot modify registration for another user ID."}), 403
    
    conn = get_db()
    cursor = conn.cursor()
    
    # Check if already registered
    cursor.execute("SELECT id FROM enrolments WHERE user_id = ? AND course_id = ?", (auth_user_id, course_id))
    enrolment = cursor.fetchone()
    
    cursor.execute("SELECT * FROM courses WHERE id = ?", (course_id,))
    course = cursor.fetchone()
    
    if not course:
        conn.close()
        return jsonify({"status": "error", "message": "Course not found"}), 404
        
    if enrolment:
        # Drop Course
        cursor.execute("DELETE FROM enrolments WHERE user_id = ? AND course_id = ?", (auth_user_id, course_id))
        cursor.execute("UPDATE courses SET enrolled_count = MAX(0, enrolled_count - 1) WHERE id = ?", (course_id,))
        conn.commit()
        conn.close()
        return jsonify({"status": "dropped", "message": f"Dropped {course['code']}"})
    else:
        # Register Course (Validations)
        
        # 1. Capacity limit check
        if course['enrolled_count'] >= course['capacity']:
            conn.close()
            return jsonify({"status": "error", "message": "Course is full"}), 400
            
        # 2. Time clash check
        cursor.execute('''
            SELECT c.code, c.time FROM enrolments e
            JOIN courses c ON e.course_id = c.id
            WHERE e.user_id = ? AND c.day = ? AND c.slot = ?
        ''', (auth_user_id, course['day'], course['slot']))
        clash = cursor.fetchone()
        if clash:
            conn.close()
            return jsonify({
                "status": "error", 
                "message": f"Schedule clash with {clash['code']} on {course['day']} at {course['time']}"
            }), 400
            
        # 3. Max credits check (limit to 24)
        cursor.execute('''
            SELECT SUM(c.credits) FROM enrolments e
            JOIN courses c ON e.course_id = c.id
            WHERE e.user_id = ?
        ''', (auth_user_id,))
        current_credits = cursor.fetchone()[0] or 0
        if current_credits + course['credits'] > 24:
            conn.close()
            return jsonify({"status": "error", "message": "Credit limit (24 CR) exceeded"}), 400

        # 4. BUG-05: Real Prerequisite Check
        prereq = course['prerequisite']
        if prereq != 'None':
            # Check if this prerequisite is offered in the catalog
            cursor.execute("SELECT id FROM courses WHERE code = ?", (prereq,))
            prereq_offered = cursor.fetchone()
            
            if prereq_offered:
                # Student must be enrolled in the prerequisite course
                cursor.execute("SELECT id FROM enrolments WHERE user_id = ? AND course_id = ?", (auth_user_id, prereq_offered[0]))
                has_prereq = cursor.fetchone()
                if not has_prereq:
                    conn.close()
                    return jsonify({
                        "status": "error", 
                        "message": f"Prerequisite requirement unmet: You must register for {prereq} first!"
                    }), 400
            
        # Complete registration
        # BUG-06: Database NOT NULL checks will block null user_id / course_id automatically
        cursor.execute("INSERT INTO enrolments (user_id, course_id) VALUES (?, ?)", (auth_user_id, course_id))
        cursor.execute("UPDATE courses SET enrolled_count = enrolled_count + 1 WHERE id = ?", (course_id,))
        conn.commit()
        conn.close()
        return jsonify({"status": "registered", "message": f"Registered for {course['code']}"})

if __name__ == '__main__':
    init_db()
    app.run(port=5000, debug=True)

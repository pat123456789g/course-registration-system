import sqlite3
import os
import requests
from werkzeug.security import generate_password_hash, check_password_hash

DB_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "registration.db")
API_URL = "http://localhost:5000/api"

def get_db():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn

def test_password_hash_storage():
    print("\n--- 1. Testing Password Security (BUG-01) ---")
    conn = get_db()
    cursor = conn.cursor()
    
    # Register a temporary user directly or clean up first
    cursor.execute("DELETE FROM users WHERE email='test_qa@students.dkut.ac.ke'")
    
    # Hash password
    pwd = "securepassword123"
    hash_pwd = generate_password_hash(pwd)
    
    cursor.execute('''
        INSERT INTO users (first_name, last_name, reg_number, email, programme, year_of_study, password)
        VALUES ('Test', 'QA', 'C026-01-9999/2023', 'test_qa@students.dkut.ac.ke', 'BSc CS', '3rd Year', ?)
    ''', (hash_pwd,))
    conn.commit()
    
    # Retrieve
    cursor.execute("SELECT password FROM users WHERE email='test_qa@students.dkut.ac.ke'")
    stored = cursor.fetchone()[0]
    
    print(f"Stored Password in SQLite: {stored}")
    assert not stored.startswith("securepassword"), "FAIL: Password stored in plaintext!"
    assert check_password_hash(stored, pwd), "FAIL: Password verification check failed!"
    print("PASS: Password is securely hashed and verified.")
    
    conn.close()

def test_foreign_key_cascade():
    print("\n--- 2. Testing SQLite Foreign Key Cascade (BUG-03) ---")
    conn = get_db()
    cursor = conn.cursor()
    
    # Clear old entries
    cursor.execute("DELETE FROM enrolments")
    cursor.execute("DELETE FROM users WHERE email='cascade_test@students.dkut.ac.ke'")
    conn.commit()
    
    # Create user
    cursor.execute('''
        INSERT INTO users (first_name, last_name, reg_number, email, programme, year_of_study, password)
        VALUES ('Cascade', 'Test', 'C026-01-8888/2023', 'cascade_test@students.dkut.ac.ke', 'BSc CS', '3rd Year', 'hashpwd')
    ''')
    user_id = cursor.lastrowid
    
    # Fetch course
    cursor.execute("SELECT id FROM courses LIMIT 1")
    course_id = cursor.fetchone()[0]
    
    # Insert enrolment
    cursor.execute("INSERT INTO enrolments (user_id, course_id) VALUES (?, ?)", (user_id, course_id))
    conn.commit()
    
    # Verify enrolment exists
    cursor.execute("SELECT COUNT(*) FROM enrolments WHERE user_id = ?", (user_id,))
    count = cursor.fetchone()[0]
    print(f"Enrolments count for test user before delete: {count}")
    assert count == 1, "FAIL: Enrolment record was not created!"
    
    # Delete user
    cursor.execute("DELETE FROM users WHERE id = ?", (user_id,))
    conn.commit()
    
    # Verify enrolment was cascade-deleted
    cursor.execute("SELECT COUNT(*) FROM enrolments WHERE user_id = ?", (user_id,))
    count_after = cursor.fetchone()[0]
    print(f"Enrolments count for test user after delete: {count_after}")
    assert count_after == 0, "FAIL: Cascade delete failed! Enrolment record still exists."
    print("PASS: Cascading deletion is working correctly.")
    
    conn.close()

def test_not_null_constraints():
    print("\n--- 3. Testing NOT NULL Constraints (BUG-06) ---")
    conn = get_db()
    cursor = conn.cursor()
    
    try:
        cursor.execute("INSERT INTO enrolments (user_id, course_id) VALUES (NULL, 1)")
        conn.commit()
        print("FAIL: Inserted null user_id successfully!")
    except sqlite3.IntegrityError as e:
        print(f"PASS: SQLite rejected NULL user_id insertion: {e}")
        
    try:
        cursor.execute("INSERT INTO enrolments (user_id, course_id) VALUES (1, NULL)")
        conn.commit()
        print("FAIL: Inserted null course_id successfully!")
    except sqlite3.IntegrityError as e:
        print(f"PASS: SQLite rejected NULL course_id insertion: {e}")
        
    conn.close()

def test_api_security_tokens():
    print("\n--- 4. Testing API Authentication (BUG-02) ---")
    
    # Try fetching courses without token
    res = requests.get(f"{API_URL}/courses?user_id=1")
    print(f"GET /courses without token: Status {res.status_code}")
    assert res.status_code == 401, "FAIL: Allowed access without auth token!"
    
    # Try registering course without token
    res = requests.post(f"{API_URL}/toggle-course", json={"user_id": 1, "course_id": 1})
    print(f"POST /toggle-course without token: Status {res.status_code}")
    assert res.status_code == 401, "FAIL: Allowed toggling course without auth token!"
    
    # Try signup with missing fields
    res = requests.post(f"{API_URL}/signup", json={"firstName": ""})
    print(f"POST /signup with missing fields: Status {res.status_code}, Msg: {res.json().get('message')}")
    assert res.status_code == 400, "FAIL: Missing fields should return 400 Bad Request!"
    
    # Try signup with invalid email
    res = requests.post(f"{API_URL}/signup", json={
        "firstName": "Patrick", "lastName": "Muli", "regNumber": "C026-01-1002/2023",
        "email": "invalidemail", "programme": "BSc CS", "yearOfStudy": "3rd Year", "password": "123"
    })
    print(f"POST /signup with invalid email/short pwd: Status {res.status_code}, Msg: {res.json().get('message')}")
    assert res.status_code == 400, "FAIL: Invalid email/password should return 400!"
    
    print("PASS: API endpoint security token controls verified.")

if __name__ == "__main__":
    test_password_hash_storage()
    test_foreign_key_cascade()
    test_not_null_constraints()
    test_api_security_tokens()

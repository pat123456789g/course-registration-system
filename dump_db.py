import sqlite3
import os

db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "registration.db")
out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "database_content.txt")

def dump():
    if not os.path.exists(db_path):
        print("Database file not found!")
        return

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    with open(out_path, "w", encoding="utf-8") as f:
        f.write("==================================================\n")
        f.write("         SQLITE DATABASE CURRENT CONTENT          \n")
        f.write("==================================================\n\n")

        # 1. Users Table
        f.write("--- 1. TABLE: users ---\n")
        cursor.execute("SELECT id, first_name, last_name, reg_number, email, programme, year_of_study FROM users")
        users = cursor.fetchall()
        f.write(f"{'ID':<4} | {'Name':<20} | {'Reg Number':<18} | {'Email':<30} | {'Programme':<25} | {'Year':<10}\n")
        f.write("-" * 115 + "\n")
        for u in users:
            name = f"{u[1]} {u[2]}"
            f.write(f"{u[0]:<4} | {name:<20} | {u[3]:<18} | {u[4]:<30} | {u[5]:<25} | {u[6]:<10}\n")
        f.write("\n\n")

        # 2. Courses Table
        f.write("--- 2. TABLE: courses ---\n")
        cursor.execute("SELECT id, code, title, department, lecturer, credits, capacity, enrolled_count, day, time FROM courses")
        courses = cursor.fetchall()
        f.write(f"{'ID':<4} | {'Code':<10} | {'Title':<30} | {'Dept':<6} | {'Lecturer':<20} | {'CR':<4} | {'Enrolled/Max':<12} | {'Schedule':<25}\n")
        f.write("-" * 120 + "\n")
        for c in courses:
            enr_cap = f"{c[7]}/{c[6]}"
            sched = f"{c[8]} {c[9]}"
            f.write(f"{c[0]:<4} | {c[1]:<10} | {c[2]:<30} | {c[3]:<6} | {c[4]:<20} | {c[5]:<4} | {enr_cap:<12} | {sched:<25}\n")
        f.write("\n\n")

        # 3. Enrolments Table
        f.write("--- 3. TABLE: enrolments ---\n")
        cursor.execute('''
            SELECT e.id, u.first_name, u.last_name, c.code, c.title, e.enrolled_at 
            FROM enrolments e
            JOIN users u ON e.user_id = u.id
            JOIN courses c ON e.course_id = c.id
        ''')
        enrolments = cursor.fetchall()
        f.write(f"{'ID':<4} | {'Student Name':<20} | {'Course Code':<12} | {'Course Title':<30} | {'Enrolled At':<20}\n")
        f.write("-" * 95 + "\n")
        for e in enrolments:
            s_name = f"{e[1]} {e[2]}"
            f.write(f"{e[0]:<4} | {s_name:<20} | {e[3]:<12} | {e[4]:<30} | {e[5]:<20}\n")
        f.write("\n")

    conn.close()
    print("Database content dumped to database_content.txt")

if __name__ == "__main__":
    dump()

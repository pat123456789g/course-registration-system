# 🎓 DeKUT Course Registration System

A full-stack, secure **University Course Registration System** built for the School of Computer Science & Information Technology at **Dedan Kimathi University of Technology (DeKUT)**. 

The application utilizes a **3-tier architecture** with a responsive frontend client, a Python Flask REST API server, and a persistent SQLite SQL database with foreign key constraints.

![HTML5](https://img.shields.io/badge/HTML5-E34F26?style=for-the-badge&logo=html5&logoColor=white)
![CSS3](https://img.shields.io/badge/CSS3-1572B6?style=for-the-badge&logo=css3&logoColor=white)
![JavaScript](https://img.shields.io/badge/JavaScript-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black)
![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)
![SQLite](https://img.shields.io/badge/SQLite-07405E?style=for-the-badge&logo=sqlite&logoColor=white)

---

## 🎨 System Diagrams

The system architecture and model designs are documented in high-resolution visual formats:
1. **Entity Relationship Diagram (ERD):** [course_registration_erd.png](course_registration_erd.png) — Details table schemas, foreign key relationships, and index constraints.
2. **Use Case Diagram:** [use_case_diagram.png](use_case_diagram.png) — Depicts actor interactions and workflow boundaries.

---

## ✨ Features

### 🔐 1. Cryptographic Security & Authentication
- **Secure Password Hashing:** Uses Werkzeug's `scrypt` hashing algorithm; plaintext passwords are never stored in the database.
- **Token-Based Session Authorization:** Login generates a unique hex token stored in a `sessions` table. Protected endpoints require the token passed via `Authorization: Bearer <token>` in HTTP request headers.
- **ID Manipulation Prevention:** Backend validates token ownership before allowing student queries or course toggle actions (returns `403 Forbidden` on mismatched IDs).

### 📚 2. Catalog & Enrolment Management
- Real-time seat capacity tracker with visual progress bars.
- Dynamic weekly schedule/timetable grid mapping enrolled courses.
- Filter by department (CS/IT) and search by code or title.

### ⚡ 3. Validation Engine (5 Constraints)
1. **Prerequisite Verification:** Queries enrolments to block registration if prerequisite courses offered in the catalog are not satisfied.
2. **Timetable Clash Detection:** Prevents registering for modules occupying overlapping slots.
3. **Capacity Limiter:** Disables registration when a course is full.
4. **Credit Cap Enforcer:** Restricts registration to a maximum of 24 credits per semester.
5. **Signup Input Validation:** Validates email formatting, registration number constraints, and minimum password lengths (6 chars) on the server side.

---

## 📁 Project Structure

```
course-registration-system/
├── app.py                                   # Python Flask REST API server & validations
├── index.html                               # Frontend single-page application client
├── schema.sql                               # Clean SQL schema database file
├── registration.db                          # Live SQLite database binary file
├── verify_fixes.py                          # Automated database & security test script
├── generate_final_reports.py                # QA reporting generator script
├── course_registration_erd.png              # Database Entity Relationship Diagram
├── use_case_diagram.png                     # System Use Case Diagram
├── Student_Course_Registration_FINAL_QA.xlsx# Excel regression testing workbook
└── Student_Course_Registration_FINAL_QA_REPORT.pdf # PDF regression testing report
```

---

## 🚀 Setup & Installation

### 1. Prerequisites
Ensure Python 3.x is installed, then install package dependencies:
```bash
pip install flask flask-cors openpyxl reportlab pillow requests
```

### 2. Start the Backend API Server
Launch the Flask backend server from the project directory:
```bash
python app.py
```
*The backend server initializes/migrates the SQLite database file `registration.db` and listens on `http://127.0.0.1:5000`.*

### 3. Run the Frontend Client
You can open `index.html` directly in any web browser, or serve it using Python's static server:
```bash
python -m http.server 8080
```
Then visit: [http://localhost:8080](http://localhost:8080)

---

## 🧪 Testing & Verification

### Running Automated QA Verification
Execute the test runner script to check password hashing, cascading deletions, NOT NULL constraints, and header authentication controls:
```bash
python verify_fixes.py
```

### Generating QA Excel and PDF Reports
To rebuild the Excel workbook and PDF Quality Assurance reports:
```bash
python generate_final_reports.py
```

---

## 👨‍💻 Author

**Patrick Muli** — Dedan Kimathi University of Technology (DeKUT)

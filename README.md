# 🎓 DeKUT Course Registration System

A full-stack web-based **University Course Registration System** built for the School of Computer Science & Information Technology at **Dedan Kimathi University of Technology (DeKUT)**.

![HTML5](https://img.shields.io/badge/HTML5-E34F26?style=for-the-badge&logo=html5&logoColor=white)
![CSS3](https://img.shields.io/badge/CSS3-1572B6?style=for-the-badge&logo=css3&logoColor=white)
![JavaScript](https://img.shields.io/badge/JavaScript-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black)

## ✨ Features

### 🔐 Authentication System
- **Login** with email & password validation
- **Create Account** with full student registration (Name, Reg No, Email, Programme, Year of Study)
- Session persistence via `localStorage`
- Password show/hide toggle
- Form validation with error alerts

### 📚 Course Catalog & Registration
- Browse **10 available courses** across Computer Science and Information Technology departments
- **Search & Filter** by course code, title, department, or day
- Real-time **seat capacity progress bars** with color indicators (green/amber/red)

### ⚡ Smart Validation Engine (4 Algorithms)
1. **Prerequisite Checker** — Blocks enrolment if required lower-level modules aren't completed
2. **Timetable Clash Detection** — Prevents registering for overlapping lecture slots
3. **Seat Capacity Limiter** — Blocks registration when a course is full
4. **Max Credits Enforcer** — Prevents exceeding the 24-credit semester limit

### 📅 Dynamic Timetable
- Auto-generated **weekly schedule grid** (Monday–Friday, 08:00–17:00)
- Updates in real-time when courses are added or dropped

### 📊 Student Dashboard
- Live stats: Registered Credits, Enrolled Courses, GPA, Semester Status
- Enrolment History audit trail table
- Student Profile page with academic summary

### 🎨 Premium UI/UX
- Modern sidebar navigation with dark theme
- Toast notification system for all actions
- Smooth animations and hover effects
- Fully **responsive** — works on desktop, tablet, and mobile

## 🚀 Quick Start

### Option 1: Open directly
Simply open `index.html` in any modern web browser (Chrome, Edge, Firefox).

### Option 2: Run with local server
```bash
python -m http.server 8080
```
Then visit: [http://localhost:8080](http://localhost:8080)

## 🔑 Default Login Credentials

| Field | Value |
|-------|-------|
| Email | `patrick@students.dkut.ac.ke` |
| Password | `123456` |

Or click **"Create Account"** to register a new student.

## 🏗️ Tech Stack

| Layer | Technology |
|-------|-----------|
| Structure | HTML5 (Semantic) |
| Styling | CSS3 (Custom Properties, Grid, Flexbox) |
| Logic | Vanilla JavaScript (ES6+) |
| Typography | Google Fonts (Outfit, Inter, JetBrains Mono) |
| Storage | Browser localStorage |

## 📁 Project Structure

```
course-registration-system/
├── index.html          # Complete single-file application (1,571 lines)
├── README.md           # Project documentation
└── LICENSE             # MIT License
```

## 📐 System Architecture

```
┌─────────────────────────────────────────────────┐
│                  PRESENTATION LAYER              │
│  Login Screen │ Sign Up Screen │ Dashboard App   │
├─────────────────────────────────────────────────┤
│                  BUSINESS LOGIC LAYER            │
│  Auth Engine │ Prerequisite Checker │ Clash Det.  │
│  Capacity Limiter │ Credits Enforcer │ Renderer  │
├─────────────────────────────────────────────────┤
│                  DATA LAYER                      │
│  localStorage (Users DB) │ In-Memory Courses DB  │
└─────────────────────────────────────────────────┘
```

## 👨‍💻 Author

**Patrick Muli** — Dedan Kimathi University of Technology

## 📄 License

This project is open source and available under the [MIT License](LICENSE).

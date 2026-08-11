-- ==========================================
-- DEKUT COURSE REGISTRATION SYSTEM SCHEMA
-- Database: SQLite
-- ==========================================

-- 1. Departments Table
CREATE TABLE IF NOT EXISTS departments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE
);

-- 2. Users Table (Students & Faculty)
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    first_name TEXT NOT NULL,
    last_name TEXT NOT NULL,
    reg_number TEXT UNIQUE NOT NULL,
    email TEXT UNIQUE NOT NULL,
    programme TEXT NOT NULL,
    year_of_study TEXT NOT NULL,
    password TEXT NOT NULL
);

-- 3. Courses Table
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
);

-- 4. Enrolments Table (Tracks Module Registration)
CREATE TABLE IF NOT EXISTS enrolments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER,
    course_id INTEGER,
    enrolled_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (user_id) REFERENCES users (id) ON DELETE CASCADE,
    FOREIGN KEY (course_id) REFERENCES courses (id) ON DELETE CASCADE,
    UNIQUE(user_id, course_id)
);

-- ==========================================
-- SEED INITIAL DATA
-- ==========================================

INSERT OR IGNORE INTO courses (code, title, department, lecturer, credits, capacity, enrolled_count, day, time, slot, prerequisite) VALUES
('CCS 3101', 'Database Management Systems', 'CS', 'Dr. Stephen Mburu', 3, 50, 42, 'Monday', '08:00 - 11:00', 0, 'CCS 2101'),
('CIT 3102', 'Network Administration', 'IT', 'Dr. Faith Kamau', 3, 45, 38, 'Monday', '11:00 - 14:00', 1, 'None'),
('CCS 3103', 'Artificial Intelligence', 'CS', 'Ms. Mercy Njagi', 3, 40, 39, 'Tuesday', '08:00 - 11:00', 0, 'CCS 2104'),
('CCS 3104', 'Software Engineering II', 'CS', 'Dr. George Musumba', 3, 60, 51, 'Wednesday', '08:00 - 11:00', 0, 'CCS 2105'),
('CIT 3105', 'Web Application Development', 'IT', 'Ms. Lydia Wanjiru', 3, 50, 48, 'Thursday', '14:00 - 17:00', 2, 'CIT 2102'),
('CCS 3106', 'Cybersecurity Principles', 'CS', 'Mr. James Mwangi', 3, 35, 35, 'Friday', '08:00 - 11:00', 0, 'CCS 3102'),
('CCS 4101', 'Human Computer Interaction', 'CS', 'Dr. Agnes Mindila', 3, 50, 25, 'Monday', '08:00 - 11:00', 0, 'None'),
('CIT 3108', 'Cloud Computing Infrastructure', 'IT', 'Mr. Anthony Njoroge', 3, 40, 18, 'Friday', '11:00 - 14:00', 1, 'CIT 3102'),
('CCS 3109', 'Operating Systems', 'CS', 'Mr. David Kimani', 3, 55, 40, 'Wednesday', '14:00 - 17:00', 2, 'CCS 2101'),
('CIT 3110', 'Mobile App Development', 'IT', 'Mr. Brian Ochieng', 3, 45, 30, 'Thursday', '08:00 - 11:00', 0, 'CIT 2102');

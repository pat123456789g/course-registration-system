import matplotlib.pyplot as plt
import matplotlib.patches as patches

def draw_erd():
    fig, ax = plt.subplots(figsize=(12, 8), dpi=150)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 80)
    ax.axis('off')
    
    # Background
    fig.patch.set_facecolor('white')
    
    # Table definitions: (x, y, width, height, title, rows)
    tables = [
        # 1. users
        (5, 45, 25, 28, "users", [
            ("id", "INTEGER PRIMARY KEY (PK)"),
            ("first_name", "TEXT NOT NULL"),
            ("last_name", "TEXT NOT NULL"),
            ("reg_number", "TEXT UNIQUE NOT NULL"),
            ("email", "TEXT UNIQUE NOT NULL"),
            ("programme", "TEXT NOT NULL"),
            ("year_of_study", "TEXT NOT NULL"),
            ("password", "TEXT NOT NULL")
        ]),
        # 2. sessions
        (5, 5, 25, 18, "sessions", [
            ("user_id", "INTEGER PRIMARY KEY (PK, FK)"),
            ("token", "TEXT UNIQUE NOT NULL"),
            ("created_at", "TIMESTAMP DEFAULT NOW")
        ]),
        # 3. enrolments
        (45, 25, 25, 22, "enrolments", [
            ("id", "INTEGER PRIMARY KEY (PK)"),
            ("user_id", "INTEGER NOT NULL (FK)"),
            ("course_id", "INTEGER NOT NULL (FK)"),
            ("enrolled_at", "TIMESTAMP DEFAULT NOW")
        ]),
        # 4. courses
        (72, 15, 25, 38, "courses", [
            ("id", "INTEGER PRIMARY KEY (PK)"),
            ("code", "TEXT UNIQUE NOT NULL"),
            ("title", "TEXT NOT NULL"),
            ("department", "TEXT NOT NULL"),
            ("lecturer", "TEXT NOT NULL"),
            ("credits", "INTEGER NOT NULL"),
            ("capacity", "INTEGER NOT NULL"),
            ("enrolled_count", "INTEGER DEFAULT 0"),
            ("day", "TEXT NOT NULL"),
            ("time", "TEXT NOT NULL"),
            ("slot", "INTEGER NOT NULL"),
            ("prerequisite", "TEXT NOT NULL")
        ])
    ]
    
    # Draw Tables
    for x, y, w, h, title, rows in tables:
        # Table Box Outline
        rect = patches.Rectangle((x, y), w, h, linewidth=1.5, edgecolor='#1A202C', facecolor='#F7FAFC')
        ax.add_patch(rect)
        
        # Table Header
        header_height = 4
        header_rect = patches.Rectangle((x, y + h - header_height), w, header_height, linewidth=1.5, edgecolor='#1A202C', facecolor='#EDF2F7')
        ax.add_patch(header_rect)
        
        # Table Title Text
        ax.text(x + w/2, y + h - header_height/2, title, fontsize=10, fontweight='bold', ha='center', va='center', color='#2D3748')
        
        # Draw Rows
        row_height = (h - header_height) / len(rows)
        for idx, (col_name, col_type) in enumerate(rows):
            curr_y = y + h - header_height - (idx + 0.5) * row_height
            
            # Key indicator
            is_key = "PK" in col_type or "FK" in col_type
            font_w = 'bold' if is_key else 'normal'
            color_c = '#2B6CB0' if is_key else '#4A5568'
            
            # Column Name
            ax.text(x + 1, curr_y, col_name, fontsize=8, fontweight=font_w, va='center', color='#1A202C')
            # Column Type / Constraints
            ax.text(x + w - 1, curr_y, col_type, fontsize=7, va='center', ha='right', color=color_c)
            
            # Row Divider Line
            if idx < len(rows) - 1:
                line_y = y + h - header_height - (idx + 1) * row_height
                ax.plot([x, x + w], [line_y, line_y], color='#E2E8F0', linewidth=0.5)

    # Draw Relation Lines
    # 1. users to sessions (1-to-1)
    ax.plot([17.5, 17.5], [45, 23], color='#4A5568', linestyle='-', linewidth=1.2)
    # Relation notation symbols
    ax.text(18.5, 43, "1", fontsize=8, color='#4A5568')
    ax.text(18.5, 24, "1", fontsize=8, color='#4A5568')
    
    # 2. users to enrolments (1-to-many)
    ax.plot([30, 38, 38, 45], [59, 59, 36, 36], color='#4A5568', linestyle='-', linewidth=1.2)
    ax.text(31, 60, "1", fontsize=8, color='#4A5568')
    ax.text(43, 37, "N", fontsize=8, color='#4A5568')
    
    # 3. courses to enrolments (1-to-many)
    ax.plot([72, 70, 70, 70], [34, 34, 36, 36], color='#4A5568', linestyle='-', linewidth=1.2) # Horizontal part
    ax.plot([70, 70, 45], [36, 36, 36], color='#4A5568', linestyle='-', linewidth=1.2) # Connect to enrolments
    ax.text(71, 35, "1", fontsize=8, color='#4A5568')
    ax.text(46, 37, "N", fontsize=8, color='#4A5568')

    plt.tight_layout()
    plt.savefig('ERD.png', bbox_inches='tight', facecolor='white')
    plt.close()
    print("ERD.png generated successfully.")

def draw_usecase():
    fig, ax = plt.subplots(figsize=(10, 8), dpi=150)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')
    
    # Background
    fig.patch.set_facecolor('white')
    
    # System boundary box
    boundary = patches.Rectangle((20, 5), 60, 90, linewidth=1.5, edgecolor='#1A202C', facecolor='#FFFFFF')
    ax.add_patch(boundary)
    ax.text(50, 92, "Course Registration Portal", fontsize=12, fontweight='bold', ha='center', va='center', color='#1A202C')
    
    # Use cases: (x, y, title)
    usecases = [
        (50, 83, "Sign Up (New Account)"),
        (50, 75, "Log In (Sign In)"),
        (50, 67, "View Course Catalog"),
        (50, 59, "Search & Filter Courses"),
        (50, 51, "Register / Enrol Course"),
        (50, 43, "Drop Active Course"),
        (50, 35, "View Weekly Timetable"),
        (50, 27, "View Enrolment History"),
        (50, 19, "View Student Profile"),
        (50, 11, "Logout (Session Clear)")
    ]
    
    # Draw Use Case Bubbles
    for x, y, title in usecases:
        ellipse = patches.Ellipse((x, y), 28, 5, edgecolor='#1A202C', facecolor='#F7FAFC', linewidth=1.2)
        ax.add_patch(ellipse)
        ax.text(x, y, title, fontsize=8, ha='center', va='center', color='#1A202C')
        
    # Draw Student Actor (Stick Figure on Left)
    actor_x = 10
    actor_y = 50
    # Head
    head = patches.Circle((actor_x, actor_y + 4), 2.5, edgecolor='#1A202C', facecolor='#FFFFFF', linewidth=1.5)
    ax.add_patch(head)
    # Body
    ax.plot([actor_x, actor_x], [actor_y + 1.5, actor_y - 5], color='#1A202C', linewidth=1.5)
    # Arms
    ax.plot([actor_x - 3, actor_x + 3], [actor_y - 1, actor_y - 1], color='#1A202C', linewidth=1.5)
    # Legs
    ax.plot([actor_x, actor_x - 3], [actor_y - 5, actor_y - 10], color='#1A202C', linewidth=1.5)
    ax.plot([actor_x, actor_x + 3], [actor_y - 5, actor_y - 10], color='#1A202C', linewidth=1.5)
    # Actor Label
    ax.text(actor_x, actor_y - 13, "Student", fontsize=10, fontweight='bold', ha='center', va='center', color='#1A202C')
    
    # Draw System Backend Actor (Right Box)
    sys_x = 88
    sys_y = 50
    sys_rect = patches.Rectangle((sys_x - 6, sys_y - 5), 12, 10, linewidth=1.5, edgecolor='#1A202C', facecolor='#EDF2F7')
    ax.add_patch(sys_rect)
    ax.text(sys_x, sys_y, "Flask API\nValidation\nEngine", fontsize=8, fontweight='bold', ha='center', va='center', color='#1A202C')
    
    # Draw Association Lines (Student to Use Cases)
    student_target_ys = [83, 75, 67, 59, 51, 43, 35, 27, 19, 11]
    for ty in student_target_ys:
        ax.plot([actor_x + 3, 36], [actor_y, ty], color='#718096', linestyle='-', linewidth=0.8)
        
    # Draw System connection to validation cases (Register, Drop, Login checks)
    sys_target_ys = [75, 51, 43] # Log in check, register, drop validations
    for ty in sys_target_ys:
        ax.plot([sys_x - 6, 64], [sys_y, ty], color='#718096', linestyle='--', linewidth=0.8)
        
    plt.tight_layout()
    plt.savefig('UseCaseDiagram.png', bbox_inches='tight', facecolor='white')
    plt.close()
    print("UseCaseDiagram.png generated successfully.")

if __name__ == "__main__":
    draw_erd()
    draw_usecase()

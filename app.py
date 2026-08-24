from flask import Flask, render_template, request, redirect, session, url_for, send_file, flash
from werkzeug.security import generate_password_hash, check_password_hash
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import text
from datetime import datetime

import pandas as pd
from dotenv import load_dotenv
import os

from werkzeug.utils import secure_filename
from openpyxl import Workbook

from werkzeug.security import generate_password_hash
print(generate_password_hash("admin123"))





# =========================================================
# DATABASE CONFIGURATION
# =========================================================
load_dotenv()

app = Flask(__name__)

app.secret_key = os.getenv("SECRET_KEY")

app.config["SQLALCHEMY_DATABASE_URI"] = (
    f"mysql+pymysql://"
    f"{os.getenv('DB_USER')}:"
    f"{os.getenv('DB_PASSWORD')}@"
    f"{os.getenv('DB_HOST')}/"
    f"{os.getenv('DB_NAME')}"
)

app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)


# =========================================================
# EXCEL IMPORT CONFIGURATION
# =========================================================

UPLOAD_FOLDER = "uploads"

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

ALLOWED_EXTENSIONS = {"xlsx", "xls"}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def allowed_file(filename):

    return (
        "." in filename
        and filename.rsplit(".", 1)[1].lower()
        in ALLOWED_EXTENSIONS
    )


# =========================================================
# DATABASE MODEL
# =========================================================


# =========maintenance_reports table=========
class MaintenanceReport(db.Model):

    __tablename__ = "maintenance_reports"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)

    room_number = db.Column(db.String(20), nullable=False)

    floor_number = db.Column(db.String(20), nullable=False)

    category = db.Column(db.String(100))

    issue_description = db.Column(db.Text, nullable=False)

    photo = db.Column(db.String(255))

    status = db.Column(
        db.String(50),
        default="Pending"
    )

    priority = db.Column(
        db.String(50),
        default="Pending"
    )

    assigned_worker = db.Column(
        db.String(100)
    )

    created_at = db.Column(
        db.DateTime,
        default=datetime.utcnow
    )

# =========workers table=========
class Worker(db.Model):

    __tablename__ = "workers"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    name = db.Column(db.String(100), nullable=False)
    phone = db.Column(db.String(30))
    department = db.Column(db.String(100))
    skills = db.Column(db.String(255))
    availability = db.Column(
        db.String(50),
        default="Available"
    )
    current_tasks = db.Column(
        db.Integer,
        default=0
    )
    created_at = db.Column(
        db.DateTime,
        default=datetime.utcnow
    )



# =========HOTEL TABLE =========

class Hotel(db.Model):

    __tablename__ = "hotel"

    id = db.Column(
        db.Integer,
        primary_key=True,
        autoincrement=True
    )

    hotel_name = db.Column(
        db.String(150),
        nullable=False
    )

    address = db.Column(
        db.String(255),
        nullable=False
    )

    floors = db.Column(
        db.Integer,
        nullable=False
    )

    rooms = db.Column(
        db.Integer,
        nullable=False
    )

    created_at = db.Column(
        db.DateTime,
        default=datetime.utcnow
    )

# =========ROOM table=========
class Room(db.Model):

    __tablename__ = "rooms"

    id = db.Column(
        db.Integer,
        primary_key=True,
        autoincrement=True
    )

    room_number = db.Column(
        db.String(20),
        nullable=False,
        unique=True
    )

    floor_number = db.Column(
        db.Integer,
        nullable=False
    )

    room_type = db.Column(
        db.String(100)
    )

    status = db.Column(
        db.String(50),
        default="Available"
    )

    created_at = db.Column(
        db.DateTime,
        default=datetime.utcnow
    )




# ========= USER TABLE =========

class User(db.Model):

    __tablename__ = "users"

    id = db.Column(
        db.Integer,
        primary_key=True,
        autoincrement=True
    )

    username = db.Column(
        db.String(100),
        unique=True,
        nullable=False
    )

    password = db.Column(
        db.String(255),
        nullable=False
    )

    role = db.Column(
        db.String(50),
        nullable=False
    )

    worker_id = db.Column(
        db.Integer,
        db.ForeignKey("workers.id"),
        nullable=True
    )

    created_at = db.Column(
        db.DateTime,
        default=datetime.utcnow
    )

    # NEW
    must_change_password = db.Column(
        db.Boolean,
        default=True,
        nullable=False
    )

    worker = db.relationship(
        "Worker",
        backref="user",
        foreign_keys=[worker_id]
    )






# =========================================================
# AI SMART TECHNICIAN ASSIGNMENT
# =========================================================

def ai_assign_worker(report):

    print("\n====================================")
    print("AI SMART TECHNICIAN ASSIGNMENT")
    print("====================================")

    print(
        "CATEGORY ENTERING AI:",
        repr(report.category)
    )

    # =====================================================
    # 1. AI PRIORITY DETECTION
    # =====================================================

    detected_priority = ai_detect_priority(
        report.category,
        report.issue_description
    )

    report.priority = detected_priority

    print("Report:", report.id)
    print("Category:", report.category)
    print("Priority:", report.priority)

    # =====================================================
    # 2. REQUIRED SKILLS
    # =====================================================

    category = (
        report.category or ""
    ).strip().lower()

    required_skills_map = {

        "hvac": [
            "hvac",
            "air conditioning"
        ],

        "plumbing": [
            "plumbing",
            "water pipes",
            "water"
        ],

        "electrical": [
            "electrical",
            "electricity",
            "wiring"
        ],

        "door": [
            "door",
            "repairs"
        ],

        "furniture": [
            "furniture",
            "woodwork"
        ],

        "tv": [
            "tv",
            "electronics"
        ],

        "it": [
            "it",
            "network",
            "computer"
        ],

        "other": [
            "general maintenance",
            "repairs"
        ]
    }

    required_skills = required_skills_map.get(
        category,
        ["general maintenance", "repairs"]
    )

    print(
        "Required skills:",
        required_skills
    )

    # =====================================================
    # 4. GET WORKERS
    # =====================================================

    workers = Worker.query.all()

    best_worker = None
    best_score = -1

    best_skill_score = 0
    best_workload_score = 0
    best_priority_score = 0

    # =====================================================
    # 5. CHECK EVERY WORKER
    # =====================================================

    for worker in workers:

        worker_name = worker.name or ""

        availability = (
            worker.availability or ""
        ).strip().lower()

        current_tasks = (
            worker.current_tasks or 0
        )

        # -------------------------------------------------
        # Only available workers
        # -------------------------------------------------

        if availability != "available":

            print(
                f"SKIP {worker_name}: "
                f"{worker.availability}"
            )

            continue

        # -------------------------------------------------
        # Worker skills
        # -------------------------------------------------

        worker_skills = []

        if worker.skills:

            worker_skills = [
                skill.strip().lower()
                for skill in worker.skills.split(",")
                if skill.strip()
            ]

        # -------------------------------------------------
        # Check skill match
        # -------------------------------------------------

        skill_match = False

        for required_skill in required_skills:

            required_skill = (
                required_skill.strip().lower()
            )

            for worker_skill in worker_skills:

                worker_skill = (
                    worker_skill.strip().lower()
                )

                if (
                    required_skill == worker_skill
                    or required_skill in worker_skill
                    or worker_skill in required_skill
                ):

                    skill_match = True
                    break

            if skill_match:
                break

        # -------------------------------------------------
        # Worker doesn't have required skill
        # -------------------------------------------------

        if not skill_match:

            print(
                f"SKIP {worker_name}: "
                f"No required skill"
            )

            continue

        # =================================================
        # 6. SKILL SCORE
        # =================================================

        skill_score = 30

        # =================================================
        # 7. AVAILABILITY SCORE
        # =================================================

        availability_score = 20

        # =================================================
        # 8. WORKLOAD SCORE
        # =================================================

        if current_tasks == 0:

            workload_score = 20

        elif current_tasks == 1:

            workload_score = 15

        elif current_tasks == 2:

            workload_score = 10

        else:

            workload_score = 5

        # =================================================
        # 9. PRIORITY SCORE
        # =================================================

        if report.priority == "Urgent":

            priority_score = 20

        elif report.priority == "High":

            priority_score = 15

        elif report.priority == "Medium":

            priority_score = 10

        else:

            priority_score = 5

        # =================================================
        # 10. TOTAL SCORE
        # =================================================

        total_score = (
            skill_score
            + availability_score
            + workload_score
            + priority_score
        )

        print(
            f"{worker_name}: "
            f"Total={total_score} | "
            f"Skill={skill_score} | "
            f"Availability={availability_score} | "
            f"Workload={workload_score} | "
            f"Priority={priority_score}"
        )

        # =================================================
        # 11. BEST WORKER
        # =================================================

        if total_score > best_score:

            best_score = total_score

            best_worker = worker

            best_skill_score = skill_score

            best_workload_score = workload_score

            best_priority_score = priority_score

    # =====================================================
    # 12. NO AVAILABLE TECHNICIAN
    # =====================================================

    if not best_worker:

        print(
            "AI WARNING: "
            "No available technician has "
            "the required skill."
        )

        report.assigned_worker = None
        report.status = "Pending"

        print("\n====================================")
        print("AI FINAL DECISION")
        print("====================================")
        print("Technician: None")
        print("Status: Pending")
        print("Priority:", report.priority)
        print("====================================")

        return None

    # =====================================================
    # 13. ASSIGN TECHNICIAN
    # =====================================================

    report.assigned_worker = best_worker.name

    report.status = "Assigned"

    # Increase worker workload

    best_worker.current_tasks = (
        best_worker.current_tasks or 0
    ) + 1

    # Change availability

    best_worker.availability = "Working"

    # =====================================================
    # 14. FINAL AI DECISION
    # =====================================================

    print("\n====================================")
    print("AI FINAL DECISION")
    print("====================================")

    print(
        "Technician:",
        best_worker.name
    )

    print(
        "Score:",
        best_score
    )

    print(
        "Skill score:",
        best_skill_score
    )

    print(
        "Workload score:",
        best_workload_score
    )

    print(
        "Priority:",
        report.priority
    )

    print(
        "Status:",
        report.status
    )

    print("====================================")

    return best_worker



# =========================================================
# AI CATEGORY DETECTION
# =========================================================

def ai_detect_category(issue_description):

    description = (
        issue_description or ""
    ).strip().lower()

    # =====================================================
    # HVAC
    # =====================================================

    hvac_keywords = [
        "air conditioner",
        "air conditioning",
        "air conditioning unit",
        "aircon",
        "a/c",
        "cooling",
        "heating",
        "heater",
        "thermostat",
        "temperature",
        "ventilation",
        "vent",
        "fan"
    ]

    if any(
        keyword in description
        for keyword in hvac_keywords
    ):
        return "HVAC"

    # =====================================================
    # PLUMBING
    # =====================================================

    plumbing_keywords = [
        "water leak",
        "water leakage",
        "water leaking",
        "leak",
        "leaking",
        "pipe",
        "pipes",
        "faucet",
        "tap",
        "sink",
        "toilet",
        "drain",
        "sewer",
        "shower",
        "flood",
        "flooding",
        "water system"
    ]

    if any(
        keyword in description
        for keyword in plumbing_keywords
    ):
        return "Plumbing"

    # =====================================================
    # ELECTRICAL
    # =====================================================

    electrical_keywords = [
        "electric",
        "electricity",
        "electrical",
        "wiring",
        "wire",
        "power outage",
        "power failure",
        "electric shock",
        "socket",
        "outlet",
        "switch",
        "light",
        "lamp",
        "bulb",
        "circuit",
        "fuse",
        "breaker",
        "sparking",
        "spark"
    ]

    if any(
        keyword in description
        for keyword in electrical_keywords
    ):
        return "Electrical"

    # =====================================================
    # IT
    # =====================================================

    it_keywords = [
        "internet",
        "wifi",
        "wi-fi",
        "network",
        "computer",
        "pc",
        "laptop",
        "printer",
        "server",
        "software",
        "system",
        "login",
        "password",
        "keyboard",
        "mouse"
    ]

    if any(
        keyword in description
        for keyword in it_keywords
    ):
        return "IT"

    # =====================================================
    # DOOR
    # =====================================================

    door_keywords = [
        "door",
        "door handle",
        "door lock",
        "lock",
        "key",
        "knob",
        "hinge"
    ]

    if any(
        keyword in description
        for keyword in door_keywords
    ):
        return "Door"

    # =====================================================
    # FURNITURE
    # =====================================================

    furniture_keywords = [
        "furniture",
        "chair",
        "table",
        "desk",
        "bed",
        "wardrobe",
        "cabinet",
        "drawer",
        "sofa",
        "couch",
        "shelf",
        "shelves"
    ]

    if any(
        keyword in description
        for keyword in furniture_keywords
    ):
        return "Furniture"

    # =====================================================
    # TV
    # =====================================================

    tv_keywords = [
        "tv",
        "television",
        "remote control",
        "remote",
        "screen",
        "display"
    ]

    if any(
        keyword in description
        for keyword in tv_keywords
    ):
        return "TV"

    # =====================================================
    # DEFAULT
    # =====================================================

    return "Other"



    # =====================================================
    # 3. REQUIRED SKILLS
    # =====================================================

    category = (
        report.category or ""
    ).strip().lower()

    skill_map = {

        "hvac": [
            "hvac",
            "air conditioning"
        ],

        "plumbing": [
            "plumbing",
            "water systems"
        ],

        "electrical": [
            "electrical",
            "wiring"
        ],

        "it": [
            "it",
            "pc",
            "computer",
            "network"
        ],

        "door": [
            "door",
            "lock"
        ],

        "furniture": [
            "furniture"
        ],

        "tv": [
            "tv",
            "electronics"
        ],

        "other": []
    }

    required_skills = skill_map.get(
        category,
        []
    )

    print(
        "Required skills:",
        required_skills
    )

    # =====================================================
    # 4. GET AVAILABLE WORKERS
    # =====================================================

    workers = Worker.query.all()

    available_workers = []

    for worker in workers:

        availability = (
            worker.availability or ""
        ).strip().lower()

        # Only Available workers
        if availability != "available":

            print(
                f"SKIP {worker.name}: "
                f"{worker.availability}"
            )

            continue

        available_workers.append(worker)

    # =====================================================
    # 5. FIND QUALIFIED WORKERS
    # =====================================================

    skilled_candidates = []

    for worker in available_workers:

        worker_skills = (
            worker.skills or ""
        ).lower()

        skill_match = False

        # -------------------------------------------------
        # Check required skills
        # -------------------------------------------------

        for required_skill in required_skills:

            if required_skill.lower() in worker_skills:

                skill_match = True
                break

        # -------------------------------------------------
        # If category has no specific skill requirement
        # -------------------------------------------------

        if not required_skills:

            skill_match = True

        # -------------------------------------------------
        # Add qualified worker
        # -------------------------------------------------

        if skill_match:

            skilled_candidates.append(worker)

        else:

            print(
                f"SKIP {worker.name}: "
                f"No matching skill"
            )

    # =====================================================
    # 6. NO QUALIFIED TECHNICIAN AVAILABLE
    # =====================================================

    if not skilled_candidates:

        print(
            "AI: No available technician has "
            "the required skill."
        )

        print(
            "AI: Task will remain Pending."
        )

        # Do NOT assign an unqualified technician

        report.status = "Pending"

        report.assigned_worker = None

        db.session.commit()

        print("====================================")
        print("AI FINAL DECISION")
        print("====================================")
        print("Technician: None")
        print("Status: Pending")
        print("Reason: No qualified technician available")
        print("====================================\n")

        return None

    # =====================================================
    # 7. SCORE QUALIFIED TECHNICIANS
    # =====================================================

    best_worker = None

    best_score = -1

    best_skill_score = 0

    best_workload_score = 0

    best_priority_score = 0

    for worker in skilled_candidates:

        # -------------------------------------------------
        # SKILL SCORE
        # -------------------------------------------------

        skill_score = 50

        # -------------------------------------------------
        # AVAILABILITY SCORE
        # -------------------------------------------------

        availability_score = 20

        # -------------------------------------------------
        # WORKLOAD SCORE
        # -------------------------------------------------

        current_tasks = (
            worker.current_tasks or 0
        )

        if current_tasks == 0:

            workload_score = 30

        elif current_tasks == 1:

            workload_score = 25

        elif current_tasks == 2:

            workload_score = 15

        elif current_tasks == 3:

            workload_score = 5

        else:

            workload_score = 0

        # -------------------------------------------------
        # PRIORITY SCORE
        # -------------------------------------------------

        if report.priority == "Urgent":

            priority_score = 20

        elif report.priority == "High":

            priority_score = 15

        elif report.priority == "Medium":

            priority_score = 10

        else:

            priority_score = 5

        # -------------------------------------------------
        # TOTAL SCORE
        # -------------------------------------------------

        total_score = (
            skill_score
            + availability_score
            + workload_score
            + priority_score
        )

        print(
            f"{worker.name}: "
            f"Total={total_score} | "
            f"Skill={skill_score} | "
            f"Availability={availability_score} | "
            f"Workload={workload_score} | "
            f"Priority={priority_score}"
        )

        # -------------------------------------------------
        # BEST WORKER
        # -------------------------------------------------

        if total_score > best_score:

            best_score = total_score

            best_worker = worker

            best_skill_score = skill_score

            best_workload_score = workload_score

            best_priority_score = priority_score

    # =====================================================
    # 8. SAFETY CHECK
    # =====================================================

    if not best_worker:

        report.status = "Pending"

        report.assigned_worker = None

        db.session.commit()

        print(
            "AI: No technician selected."
        )

        return None

    # =====================================================
    # 9. ASSIGN TASK
    # =====================================================

    report.assigned_worker = best_worker.name

    report.status = "Assigned"

    # =====================================================
    # 10. UPDATE WORKER STATUS
    # =====================================================

    if best_worker.current_tasks is None:

        best_worker.current_tasks = 0

    best_worker.current_tasks += 1

    best_worker.availability = "Working"

    # =====================================================
    # 11. SAVE DATABASE
    # =====================================================

    db.session.commit()

    # =====================================================
    # 12. FINAL AI DECISION
    # =====================================================

    print("\n====================================")
    print("AI FINAL DECISION")
    print("====================================")

    print(
        "Technician:",
        best_worker.name
    )

    print(
        "Score:",
        best_score
    )

    print(
        "Skill score:",
        best_skill_score
    )

    print(
        "Workload score:",
        best_workload_score
    )

    print(
        "Priority score:",
        best_priority_score
    )

    print(
        "Category:",
        report.category
    )

    print(
        "Priority:",
        report.priority
    )

    print(
        "Status:",
        report.status
    )

    print("====================================\n")

    return best_worker


# =========================================================
# AI PRIORITY DETECTION
# =========================================================

def ai_detect_priority(category, issue_description):

    category = (category or "").strip().lower()
    description = (issue_description or "").strip().lower()

    # -----------------------------------------------------
    # URGENT
    # -----------------------------------------------------

    urgent_keywords = [
        "fire",
        "smoke",
        "flood",
        "flooding",
        "gas leak",
        "gas leakage",
        "major leak",
        "water flooding",
        "electric shock",
        "danger",
        "dangerous",
        "emergency",
        "sparking",
        "spark",
        "explosion"
    ]

    if any(
        keyword in description
        for keyword in urgent_keywords
    ):
        return "Urgent"

    # -----------------------------------------------------
    # HIGH
    # -----------------------------------------------------

    high_keywords = [
        "not working",
        "doesn't work",
        "does not work",
        "broken",
        "no electricity",
        "power outage",
        "no water",
        "water leak",
        "leaking",
        "air conditioner",
        "air conditioning",
        "ac not working",
        "heating not working",
        "internet not working",
        "wifi not working",
        "toilet blocked",
        "toilet not working",
        "door cannot close",
        "door won't close",
        "door won't open"
    ]

    if any(
        keyword in description
        for keyword in high_keywords
    ):
        return "High"

    # -----------------------------------------------------
    # CATEGORY-BASED HIGH PRIORITY
    # -----------------------------------------------------

    if category in [
        "electrical",
        "electricity",
        "wiring"
    ]:
        return "High"

    if category in [
        "plumbing",
        "water",
        "water systems"
    ]:
        return "High"

    if category in [
        "hvac",
        "air conditioning",
        "air_conditioning"
    ]:
        return "High"

    # -----------------------------------------------------
    # MEDIUM
    # -----------------------------------------------------

    medium_keywords = [
        "door",
        "window",
        "furniture",
        "sink",
        "toilet",
        "light",
        "lamp",
        "tv",
        "remote",
        "chair",
        "table",
        "bed"
    ]

    if any(
        keyword in description
        for keyword in medium_keywords
    ):
        return "Medium"

    # -----------------------------------------------------
    # DEFAULT
    # -----------------------------------------------------

    return "Low"




# =========================================================
# ai_assignment
# =========
@app.route("/ai_assignments")
def ai_assignments():
    
    return render_template("ai_assignments.html")

# =========================================================
# HOME
# =========================================================

@app.route("/")
def home():
    return render_template("home.html")


# =========================================================
# REPORTER
# =========================================================

@app.route("/report")
def report():
    return render_template("report.html")


# =========================================================
# SUBMIT MAINTENANCE REPORT
# =========================================================

@app.route("/submit_report", methods=["POST"])
def submit_report():

    room_number = request.form.get("room_number")
    floor_number = request.form.get("floor_number")
    category = request.form.get("category")
    print("====================================")
    print("FORM CATEGORY:", category)
    print("ALL FORM DATA:", request.form)
    print("====================================")



    issue_description = request.form.get("issue_description")

    emergency_number = request.form.get("emergency_number")
    contact_info = request.form.get("contact_info")

    photo = request.files.get("photo")

    photo_filename = None

    if photo and photo.filename:
        photo_filename = photo.filename

    # =====================================================
    # CREATE REPORT
    # =====================================================

    new_report = MaintenanceReport(
        room_number=room_number,
        floor_number=floor_number,
        category=category,
        issue_description=issue_description,
        photo=photo_filename,
        status="Pending",
        priority="Pending",
        assigned_worker=None
    )

    db.session.add(new_report)

    # Save first so the report gets its ID
    db.session.commit()

    # =====================================================
    # AI CATEGORY DETECTION
    # =====================================================

    category = request.form.get("category")
    issue_description = request.form.get("issue_description")

    print("FORM CATEGORY:", category)

    if category and category.strip().lower() != "other":

        final_category = category.strip()

    else:

        final_category = ai_detect_category(
            issue_description
        )

    new_report.category = final_category
    print("FINAL CATEGORY:", final_category)

    # =====================================================
    # AI PRIORITY DETECTION
    # =====================================================

    detected_priority = ai_detect_priority(
        final_category,
        issue_description
    )

    new_report.priority = detected_priority

    db.session.commit()

    # =====================================================
    # AI TECHNICIAN ASSIGNMENT
    # =====================================================

    ai_assign_worker(new_report)

    # Save assignment/status changes
    db.session.commit()

    # =====================================================
    # DEBUG OUTPUT
    # =====================================================

    print("\n====================================")
    print("AI SMART TECHNICIAN ASSIGNMENT")
    print("====================================")
    print("Report:", new_report.id)
    print("Category:", new_report.category)
    print("Priority:", new_report.priority)
    print("Status:", new_report.status)
    print("Assigned:", new_report.assigned_worker)
    print("====================================\n")

    print("\n==============================")
    print("NEW MAINTENANCE REPORT")
    print("==============================")
    print("ID:", new_report.id)
    print("Room:", room_number)
    print("Floor:", floor_number)
    print("Category:", new_report.category)
    print("Problem:", issue_description)
    print("Photo:", photo_filename)
    print("Status:", new_report.status)
    print("Assigned:", new_report.assigned_worker)
    print("==============================\n")

    return f"""
        <h2>Report submitted successfully!</h2>
        <p>Your maintenance request has been received.</p>
        <p>Report ID: {new_report.id}</p>
        <a href="/">Back to Home</a>
    """

# =========================================================
# LOGIN
# =========================================================


# =========================================================
# LOGIN PAGE
# =========================================================

@app.route("/login", methods=["GET"])
def login():
    return render_template("login.html")


# =========================================================
# LOGIN PROCESS
# =========================================================
@app.route("/login", methods=["POST"])
def login_post():

    username = request.form.get("username", "").strip()
    password = request.form.get("password", "")

    if not username or not password:
        return render_template(
            "login.html",
            error="Username and password are required."
        )

    # Find user by username
    user = User.query.filter_by(
        username=username
    ).first()

    # User doesn't exist
    if not user:

        return render_template(
            "login.html",
            error="Invalid username or password."
        )

    # Check password
    if not check_password_hash(
        user.password,
        password
    ):

        return render_template(
            "login.html",
            error="Invalid username or password."
        )

    # =========================
    # CREATE LOGIN SESSION
    # =========================

    session["logged_in"] = True

    session["user_id"] = user.id

    session["username"] = user.username

    session["role"] = user.role

  
    # =========================
    # REDIRECT BY ROLE
    # =========================

    if user.role == "admin":

        return redirect(
            url_for("admin_dashboard")
        )

    elif user.role == "supervisor":

        return redirect(
            url_for("supervisor_dashboard")
        )

    elif user.role == "technician":

        return redirect(
            url_for("technician_dashboard")
        )

    else:

        session.clear()

        return render_template(
            "login.html",
            error="Unknown user role."
        )
 
# =========================================================
# LOGOUT
# =========================================================

@app.route("/logout")
def logout():

    session.clear()

    return redirect(url_for("home"))


@app.route("/admin_settings/change-password", methods=["POST"])
def admin_change_password():

    # Check login
    if not session.get("logged_in"):
        return redirect(url_for("login"))

    # Check admin role
    if session.get("role") != "admin":
        return redirect(url_for("login"))

    # Get current admin username
    current_username = session.get("username")

    if not current_username:
        return redirect(url_for("login"))

    # Find admin account
    user = User.query.filter_by(
        username=current_username
    ).first()

    if not user:
        return redirect(
            url_for(
                "admin_settings",
                error="Admin account not found."
            )
        )

    # Get form data
    new_username = request.form.get(
        "username",
        ""
    ).strip()

    current_password = request.form.get(
        "current_password",
        ""
    )

    new_password = request.form.get(
        "new_password",
        ""
    )

    confirm_password = request.form.get(
        "confirm_password",
        ""
    )

    # =====================================================
    # CHECK CURRENT PASSWORD
    # =====================================================

    if not current_password:

        return redirect(
            url_for(
                "admin_settings",
                error="Please enter your current password."
            )
        )

    if not check_password_hash(
        user.password,
        current_password
    ):

        return redirect(
            url_for(
                "admin_settings",
                error="Current password is incorrect."
            )
        )

    # =====================================================
    # USERNAME
    # =====================================================

    if new_username:

        # Check if username changed
        if new_username != user.username:

            existing_user = User.query.filter_by(
                username=new_username
            ).first()

            if existing_user:

                return redirect(
                    url_for(
                        "admin_settings",
                        error="This username is already in use."
                    )
                )

            user.username = new_username

            # Update session
            session["username"] = new_username

    # =====================================================
    # PASSWORD
    # =====================================================

    if new_password or confirm_password:

        if not new_password or not confirm_password:

            return redirect(
                url_for(
                    "admin_settings",
                    error="Please fill in both new password fields."
                )
            )

        if new_password != confirm_password:

            return redirect(
                url_for(
                    "admin_settings",
                    error="New passwords do not match."
                )
            )

        if len(new_password) < 6:

            return redirect(
                url_for(
                    "admin_settings",
                    error="Password must contain at least 6 characters."
                )
            )

        user.password = generate_password_hash(
            new_password
        )

    # =====================================================
    # SAVE
    # =====================================================

    db.session.commit()

    return redirect(
        url_for(
            "admin_settings",
            success="Admin account updated successfully."
        )
    )



@app.route("/create-admin-test")
def create_admin_test():

    password_hash = generate_password_hash("admin123")

    admin = User.query.filter_by(
        username="admin"
    ).first()

    if admin:

        admin.password = password_hash
        admin.role = "admin"
        admin.worker_id = None

    else:

        admin = User(
            username="admin",
            password=password_hash,
            role="admin",
            worker_id=None
        )

        db.session.add(admin)

    db.session.commit()

    return "Admin password has been set to admin123"


@app.route("/admin_staff")
def admin_staff():

    users = User.query.order_by(
        User.created_at.desc()
    ).all()

    workers = Worker.query.order_by(
        Worker.name.asc()
    ).all()

    return render_template(
        "admin_staff.html",
        users=users,
        workers=workers
    )

@app.route(
    "/admin/staff/<int:user_id>/reset-password",
    methods=["POST"]
)
def admin_reset_staff_password(user_id):

    # Check admin
    if not session.get("logged_in"):
        return redirect(url_for("login"))

    if session.get("role") != "admin":
        return "Unauthorized", 403

    user = User.query.get_or_404(user_id)

    new_password = request.form.get(
        "new_password",
        ""
    ).strip()

    if len(new_password) < 6:

        flash(
            "Password must contain at least 6 characters.",
            "error"
        )

        return redirect(
            url_for("admin_staff")
        )

    user.password = generate_password_hash(
        new_password
    )

    # Force staff member to change it
    # after logging in.
    user.must_change_password = True

    db.session.commit()

    flash(
        "Password reset successfully.",
        "success"
    )

    return redirect(
        url_for("admin_staff")
    )



# =========================================================
# ADMIN - EDIT STAFF
# =========================================================

@app.route("/admin_edit_staff/<int:user_id>", methods=["GET", "POST"])
def admin_edit_staff(user_id):

    # Check login
    if not session.get("logged_in"):
        return redirect(url_for("login"))

    # Only admin
    if session.get("role") != "admin":
        return "Unauthorized", 403

    user = User.query.get_or_404(user_id)

    workers = Worker.query.order_by(
        Worker.name.asc()
    ).all()

    if request.method == "POST":

        username = request.form.get("username", "").strip()
        role = request.form.get("role", "").strip()
        worker_id = request.form.get("worker_id")

        # ---------------------------------------------
        # Validate username
        # ---------------------------------------------

        if not username:

            flash(
                "Username is required.",
                "error"
            )

            return redirect(
                url_for(
                    "admin_edit_staff",
                    user_id=user.id
                )
            )

        # ---------------------------------------------
        # Check duplicate username
        # ---------------------------------------------

        existing_user = User.query.filter(
            User.username == username,
            User.id != user.id
        ).first()

        if existing_user:

            flash(
                "This username already exists.",
                "error"
            )

            return redirect(
                url_for(
                    "admin_edit_staff",
                    user_id=user.id
                )
            )

        # ---------------------------------------------
        # Validate role
        # ---------------------------------------------

        allowed_roles = [
            "admin",
            "supervisor",
            "technician"
        ]

        if role not in allowed_roles:

            flash(
                "Invalid role.",
                "error"
            )

            return redirect(
                url_for(
                    "admin_edit_staff",
                    user_id=user.id
                )
            )

        # ---------------------------------------------
        # Worker connection
        # ---------------------------------------------

        if role == "technician":

            if worker_id:

                try:
                    worker_id = int(worker_id)
                except ValueError:

                    flash(
                        "Invalid worker.",
                        "error"
                    )

                    return redirect(
                        url_for(
                            "admin_edit_staff",
                            user_id=user.id
                        )
                    )

                worker = Worker.query.get(worker_id)

                if not worker:

                    flash(
                        "Worker not found.",
                        "error"
                    )

                    return redirect(
                        url_for(
                            "admin_edit_staff",
                            user_id=user.id
                        )
                    )

                user.worker_id = worker.id

            else:

                flash(
                    "A technician must be connected to a worker.",
                    "error"
                )

                return redirect(
                    url_for(
                        "admin_edit_staff",
                        user_id=user.id
                    )
                )

        else:

            # Admin and supervisor don't need worker
            user.worker_id = None

        # ---------------------------------------------
        # Update account
        # ---------------------------------------------

        user.username = username
        user.role = role

        db.session.commit()

        # Keep current session correct
        if session.get("username") == user.username:
            session["username"] = username

        flash(
            "Staff account updated successfully.",
            "success"
        )

        return redirect(
            url_for("admin_staff")
        )

    return render_template(
        "admin_edit_staff.html",
        user=user,
        workers=workers
    )



# =========================================================
# ADMIN - CHANGE STAFF PASSWORD
# =========================================================

@app.route(
    "/admin_change_staff_password/<int:user_id>",
    methods=["POST"]
)
def admin_change_staff_password(user_id):

    # Check login
    if not session.get("logged_in"):
        return redirect(url_for("login"))

    # Only admin
    if session.get("role") != "admin":
        return "Unauthorized", 403

    user = User.query.get_or_404(user_id)

    new_password = request.form.get(
        "new_password",
        ""
    ).strip()

    confirm_password = request.form.get(
        "confirm_password",
        ""
    ).strip()

    # ---------------------------------------------
    # Check fields
    # ---------------------------------------------

    if not new_password or not confirm_password:

        flash(
            "Please fill in both password fields.",
            "error"
        )

        return redirect(
            url_for(
                "admin_edit_staff",
                user_id=user.id
            )
        )

    # ---------------------------------------------
    # Confirm password
    # ---------------------------------------------

    if new_password != confirm_password:

        flash(
            "Passwords do not match.",
            "error"
        )

        return redirect(
            url_for(
                "admin_edit_staff",
                user_id=user.id
            )
        )

    # ---------------------------------------------
    # Minimum length
    # ---------------------------------------------

    if len(new_password) < 6:

        flash(
            "Password must contain at least 6 characters.",
            "error"
        )

        return redirect(
            url_for(
                "admin_edit_staff",
                user_id=user.id
            )
        )

    # ---------------------------------------------
    # Hash password
    # ---------------------------------------------

    user.password = generate_password_hash(
        new_password
    )

    db.session.commit()

    flash(
        "Password changed successfully.",
        "success"
    )

    return redirect(
        url_for(
            "admin_edit_staff",
            user_id=user.id
        )
    )




@app.route("/admin/delete_staff/<int:user_id>", methods=["POST"])
def admin_delete_staff(user_id):

    user = User.query.get_or_404(user_id)

    # Don't allow deleting the main admin account
    if user.username == "admin":
        flash("The main admin account cannot be deleted.", "error")
        return redirect(url_for("admin_staff"))

    db.session.delete(user)
    db.session.commit()

    flash("Staff account deleted successfully.", "success")

    return redirect(url_for("admin_staff"))


@app.route(
    "/admin_account",
    methods=["GET", "POST"]
)
def admin_account():

    if not session.get("logged_in"):
        return redirect(url_for("login"))

    if session.get("role") != "admin":
        return "Unauthorized", 403

    username = session.get("username")

    user = User.query.filter_by(
        username=username
    ).first()

    if not user:
        return "Admin account not found.", 404

    if request.method == "POST":

        new_username = request.form.get(
            "username",
            ""
        ).strip()

        current_password = request.form.get(
            "current_password",
            ""
        )

        new_password = request.form.get(
            "new_password",
            ""
        )

        confirm_password = request.form.get(
            "confirm_password",
            ""
        )

        # ---------------------------------------------
        # Username
        # ---------------------------------------------

        if not new_username:

            flash(
                "Username cannot be empty.",
                "error"
            )

            return redirect(
                url_for("admin_account")
            )

        # ---------------------------------------------
        # Check username duplicate
        # ---------------------------------------------

        existing_user = User.query.filter(
            User.username == new_username,
            User.id != user.id
        ).first()

        if existing_user:

            flash(
                "This username is already used.",
                "error"
            )

            return redirect(
                url_for("admin_account")
            )

        # ---------------------------------------------
        # Password change
        # ---------------------------------------------

        if new_password:

            if not current_password:

                flash(
                    "Enter your current password.",
                    "error"
                )

                return redirect(
                    url_for("admin_account")
                )

            if not check_password_hash(
                user.password,
                current_password
            ):

                flash(
                    "Current password is incorrect.",
                    "error"
                )

                return redirect(
                    url_for("admin_account")
                )

            if len(new_password) < 6:

                flash(
                    "New password must contain at least 6 characters.",
                    "error"
                )

                return redirect(
                    url_for("admin_account")
                )

            if new_password != confirm_password:

                flash(
                    "New passwords do not match.",
                    "error"
                )

                return redirect(
                    url_for("admin_account")
                )

            user.password = generate_password_hash(
                new_password
            )

        # ---------------------------------------------
        # Update username
        # ---------------------------------------------

        user.username = new_username

        db.session.commit()

        # IMPORTANT:
        # Update session after changing username

        session["username"] = new_username

        flash(
            "Admin account updated successfully.",
            "success"
        )

        return redirect(
            url_for("admin_account")
        )

    return render_template(
        "admin_account.html",
        user=user
    )




# =========================================================
# TECHNICIAN
# =========================================================

@app.route("/Technician_Dashboard")
def technician_dashboard():

    # Check login
    if not session.get("logged_in"):
        return redirect(url_for("login"))

    # Check role
    if session.get("role") != "technician":
        return redirect(url_for("login"))

    username = session.get("username")

    # Find technician account
    user = User.query.filter_by(
        username=username
    ).first()

    if not user:
        return "Technician account not found.", 404

    # Technician must be connected to a worker
    if not user.worker_id:
        return "This technician account is not connected to a worker.", 400

    # Find worker
    worker = Worker.query.get(user.worker_id)

    if not worker:
        return "Worker record not found.", 404

    # Get this technician's tasks
    tasks = MaintenanceReport.query.filter_by(
        assigned_worker=worker.name
    ).order_by(
        MaintenanceReport.created_at.desc()
    ).all()

    # Statistics
    total_tasks = len(tasks)

    pending_tasks = sum(
        1 for task in tasks
        if task.status == "Pending"
    )

    active_tasks = sum(
        1 for task in tasks
        if task.status in ["Assigned", "In Progress"]
    )

    completed_tasks = sum(
        1 for task in tasks
        if task.status == "Completed"
    )

    urgent_tasks = sum(
        1 for task in tasks
        if task.priority == "Urgent"
    )

    return render_template(
        "Technician_Dashboard.html",

        worker=worker,
        tasks=tasks,

        total_tasks=total_tasks,
        pending_tasks=pending_tasks,
        active_tasks=active_tasks,
        completed_tasks=completed_tasks,
        urgent_tasks=urgent_tasks
    )








# =========================================================
# TECHNICIAN HISTORY
# =========================================================

@app.route("/technician_history")
def technician_history():

    # Check login
    if not session.get("logged_in"):
        return redirect(url_for("login"))

    # Only technicians
    if session.get("role") != "technician":
        return redirect(url_for("login"))

    # Get logged-in user
    user_id = session.get("user_id")

    user = User.query.get(user_id)

    if not user:
        return redirect(url_for("login"))

    # Technician must be connected to a worker
    if not user.worker_id:
        return "Technician account is not connected to a worker.", 400

    # Get worker
    worker = Worker.query.get(user.worker_id)

    if not worker:
        return "Worker record not found.", 404

    # Get completed tasks using WORKER NAME
    tasks = MaintenanceReport.query.filter(
        MaintenanceReport.assigned_worker == worker.name,
        MaintenanceReport.status == "Completed"
    ).order_by(
        MaintenanceReport.created_at.desc()
    ).all()

    return render_template(
        "technician_history.html",
        tasks=tasks,
        worker=worker,
        username=user.username
    )



# =========================================================
# TECHNICIAN NOTIFICATIONS
# =========================================================

@app.route("/technician_notifications")
def technician_notifications():

    # Check login
    if not session.get("logged_in"):
        return redirect(url_for("login"))

    # Only technicians
    if session.get("role") != "technician":
        return redirect(url_for("login"))

    # Get logged-in user
    user_id = session.get("user_id")

    user = User.query.get(user_id)

    if not user:
        return redirect(url_for("login"))

    # Technician must be connected to a worker
    if not user.worker_id:
        return "Technician account is not connected to a worker.", 400

    # Get worker
    worker = Worker.query.get(user.worker_id)

    if not worker:
        return "Worker record not found.", 404

    # Get tasks assigned to this worker
    assigned_tasks = MaintenanceReport.query.filter(
        MaintenanceReport.assigned_worker == worker.name
    ).order_by(
        MaintenanceReport.created_at.desc()
    ).all()

    notifications = []

    for task in assigned_tasks:

        # =========================
        # ASSIGNED
        # =========================

        if task.status == "Assigned":

            notifications.append({
                "title": "New Task Assigned",

                "message": (
                    f"You have been assigned a maintenance "
                    f"task for Room {task.room_number}: "
                    f"{task.issue_description}"
                ),

                "time": (
                    task.created_at.strftime("%d/%m/%Y %H:%M")
                    if task.created_at
                    else "Recently"
                ),

                "type": "assigned",

                "unread": True
            })

        # =========================
        # IN PROGRESS
        # =========================

        elif task.status == "In Progress":

            notifications.append({
                "title": "Task In Progress",

                "message": (
                    f"Task #{task.id} for Room "
                    f"{task.room_number} is currently "
                    f"in progress."
                ),

                "time": (
                    task.created_at.strftime("%d/%m/%Y %H:%M")
                    if task.created_at
                    else "Recently"
                ),

                "type": "info",

                "unread": False
            })

        # =========================
        # COMPLETED
        # =========================

        elif task.status == "Completed":

            notifications.append({
                "title": "Task Completed",

                "message": (
                    f"Task #{task.id} for Room "
                    f"{task.room_number} has been completed."
                ),

                "time": (
                    task.created_at.strftime("%d/%m/%Y %H:%M")
                    if task.created_at
                    else "Recently"
                ),

                "type": "completed",

                "unread": False
            })

    return render_template(
        "technician_notifications.html",
        notifications=notifications,
        worker=worker,
        username=user.username
    )


# =========================================================
# TECHNICIAN PROFILE
# =========================================================

@app.route("/technician_profile")
def technician_profile():

    username = session.get("username")

    if not username:
        return redirect(url_for("login"))

    worker = Worker.query.filter_by(
        name=username
    ).first()

    return render_template(
        "technician_profile.html",

        username=username,

        email=session.get("email", ""),

        phone=worker.phone if worker else "",

        department=worker.department if worker else "Maintenance"
    )


# =========================================================
# TECHNICIAN UPDATE PROFILE
# =========================================================

@app.route("/technician/update-profile", methods=["POST"])
def technician_update_profile():

    old_username = session.get("username")

    if not old_username:
        return redirect(url_for("login"))

    username = request.form.get("username", "").strip()
    email = request.form.get("email", "").strip()
    phone = request.form.get("phone", "").strip()

    worker = Worker.query.filter_by(
        name=old_username
    ).first()

    if worker:

        if username:
            worker.name = username

        worker.phone = phone

        db.session.commit()

    if username:
        session["username"] = username

    session["email"] = email

    flash(
        "Profile updated successfully.",
        "success"
    )

    return redirect(
        url_for("technician_profile")
    )


# =========================================================
# TECHNICIAN CHANGE PASSWORD
# =========================================================

@app.route(
    "/technician/change-password",
    methods=["POST"]
)
def technician_change_password():

    current_password = request.form.get(
        "current_password"
    )

    new_password = request.form.get(
        "new_password"
    )

    confirm_password = request.form.get(
        "confirm_password"
    )

    if not current_password or not new_password or not confirm_password:

        flash(
            "Please fill in all password fields.",
            "error"
        )

        return redirect(
            url_for("technician_profile")
        )

    if new_password != confirm_password:

        flash(
            "New passwords do not match.",
            "error"
        )

        return redirect(
            url_for("technician_profile")
        )

    if len(new_password) < 6:

        flash(
            "Password must contain at least 6 characters.",
            "error"
        )

        return redirect(
            url_for("technician_profile")
        )

    # -----------------------------------------------------
    # IMPORTANT
    # -----------------------------------------------------
    # Your current application does NOT have a User model.
    #
    # Therefore we cannot currently update a database
    # password here.
    #
    # We will connect this to the real authentication
    # system when we create the User table.
    # -----------------------------------------------------

    flash(
        "Password validation completed. "
        "Database authentication will be connected later.",
        "success"
    )

    return redirect(
        url_for("technician_profile")
    )


# =========================================================
# TECHNICIAN - START TASK
# =========================================================

@app.route("/technician/start-task/<int:report_id>", methods=["POST"])
def technician_start_task(report_id):

    # Check login
    if not session.get("logged_in"):
        return redirect(url_for("login"))

    # Only technicians
    if session.get("role") != "technician":
        return redirect(url_for("login"))

    # Get logged-in user
    username = session.get("username")

    user = User.query.filter_by(
        username=username
    ).first()

    if not user:
        return "Technician account not found.", 404

    # Get worker
    worker = Worker.query.get(user.worker_id)

    if not worker:
        return "Worker account is not connected to a worker.", 404

    # Get report
    task = MaintenanceReport.query.get_or_404(report_id)

    # Security check:
    # technician can only start their own task
    if not task.assigned_worker:
        return "This task is not assigned to you.", 403

    if task.assigned_worker.strip().lower() != worker.name.strip().lower():
        return "You are not authorized to start this task.", 403

    # Only Assigned tasks can be started
    if task.status != "Assigned":
        return redirect(
            url_for("technician_dashboard")
        )

    # Change status
    task.status = "In Progress"

    db.session.commit()

    return redirect(
        url_for("technician_dashboard")
    )


# =========================================================
# TECHNICIAN - COMPLETE TASK
# =========================================================

@app.route("/technician/complete-task/<int:report_id>", methods=["POST"])
def technician_complete_task(report_id):

    # Check login
    if not session.get("logged_in"):
        return redirect(url_for("login"))

    # Only technicians
    if session.get("role") != "technician":
        return redirect(url_for("login"))

    # Get logged-in technician
    username = session.get("username")

    user = User.query.filter_by(
        username=username
    ).first()

    if not user:
        return "Technician account not found.", 404

    # Get worker
    worker = Worker.query.get(user.worker_id)

    if not worker:
        return "Worker account is not connected to a worker.", 404

    # Get task
    task = MaintenanceReport.query.get_or_404(report_id)

    # Check assignment
    if not task.assigned_worker:
        return "This task is not assigned to you.", 403

    if task.assigned_worker.strip().lower() != worker.name.strip().lower():
        return "You are not authorized to complete this task.", 403

    # Task must be in progress
    if task.status != "In Progress":
        return redirect(
            url_for("technician_dashboard")
        )

    # Complete task
    task.status = "Completed"

    # Reduce worker's current task count
    if worker.current_tasks > 0:
        worker.current_tasks -= 1

    # Make worker available again
    worker.availability = "Available"

    db.session.commit()

    return redirect(
        url_for("technician_dashboard")
    )




# =========================================================
# SUPERVISOR
# =========================================================

@app.route("/supervisor_dashboard")
def supervisor_dashboard():

    # =====================================================
    # MAINTENANCE REPORTS
    # =====================================================

    reports = MaintenanceReport.query.order_by(
        MaintenanceReport.created_at.desc()
    ).limit(5).all()

    total_reports = MaintenanceReport.query.count()

    urgent_reports = MaintenanceReport.query.filter_by(
        priority="Urgent"
    ).count()

    active_tasks = MaintenanceReport.query.filter(
        MaintenanceReport.status.in_([
            "Assigned",
            "In Progress"
        ])
    ).count()

    completed_today = MaintenanceReport.query.filter_by(
        status="Completed"
    ).count()


    # =====================================================
    # WORKERS
    # =====================================================

    workers = Worker.query.order_by(
        Worker.name.asc()
    ).limit(5).all()

    total_workers = Worker.query.count()

    available_workers = Worker.query.filter_by(
        availability="Available"
    ).count()

    busy_workers = Worker.query.filter(
        Worker.availability.in_([
            "Busy",
            "Working"
        ])
    ).count()

    offline_workers = Worker.query.filter_by(
        availability="Offline"
    ).count()


    # =====================================================
    # NOTIFICATIONS
    # =====================================================

    notification_count = MaintenanceReport.query.filter(
        MaintenanceReport.status == "Pending"
    ).count()


    # =====================================================
    # AI ASSIGNMENT STATISTICS
    # =====================================================

    assignments_today = MaintenanceReport.query.filter(
        MaintenanceReport.assigned_worker.isnot(None)
    ).count()

    successful_assignments = MaintenanceReport.query.filter(
        MaintenanceReport.status == "Completed",
        MaintenanceReport.assigned_worker.isnot(None)
    ).count()

    waiting_for_review = MaintenanceReport.query.filter(
        MaintenanceReport.status == "Pending"
    ).count()


    # =====================================================
    # LATEST ASSIGNMENT
    # =====================================================

    latest_assignment = MaintenanceReport.query.filter(
        MaintenanceReport.assigned_worker.isnot(None)
    ).order_by(
        MaintenanceReport.id.desc()
    ).first()


    # =====================================================
    # SEND DATA TO TEMPLATE
    # =====================================================

    return render_template(
        "supervisor_dashboard.html",

        reports=reports,

        total_reports=total_reports,
        urgent_reports=urgent_reports,
        active_tasks=active_tasks,
        completed_today=completed_today,

        workers=workers,
        total_workers=total_workers,
        available_workers=available_workers,
        busy_workers=busy_workers,
        offline_workers=offline_workers,

        notification_count=notification_count,

        assignments_today=assignments_today,
        successful_assignments=successful_assignments,
        waiting_for_review=waiting_for_review,

        latest_assignment=latest_assignment
    )




# =========================================================
# SUPERVISOR TASKS
# =========================================================

@app.route("/supervisor_tasks")
def supervisor_tasks():

    selected_status = request.args.get(
        "status",
        "All"
    )

    selected_priority = request.args.get(
        "priority",
        "All"
    )

    # =====================================================
    # BASE QUERY
    # =====================================================

    query = MaintenanceReport.query

    # =====================================================
    # STATUS FILTER
    # =====================================================

    if selected_status != "All":

        query = query.filter(
            MaintenanceReport.status == selected_status
        )

    # =====================================================
    # PRIORITY FILTER
    # =====================================================

    if selected_priority != "All":

        query = query.filter(
            MaintenanceReport.priority == selected_priority
        )

    # =====================================================
    # GET TASKS
    # =====================================================

    tasks = query.order_by(
        MaintenanceReport.created_at.desc()
    ).all()

    # =====================================================
    # STATISTICS
    # =====================================================

    pending_tasks = MaintenanceReport.query.filter_by(
        status="Pending"
    ).count()

    assigned_tasks = MaintenanceReport.query.filter_by(
        status="Assigned"
    ).count()

    in_progress_tasks = MaintenanceReport.query.filter_by(
        status="In Progress"
    ).count()

    completed_tasks = MaintenanceReport.query.filter_by(
        status="Completed"
    ).count()

    # =====================================================
    # GET WORKERS
    # =====================================================

    workers = Worker.query.order_by(
        Worker.name.asc()
    ).all()

    # =====================================================
    # RENDER
    # =====================================================

    return render_template(
        "supervisor_tasks.html",

        tasks=tasks,

        workers=workers,

        pending_tasks=pending_tasks,
        assigned_tasks=assigned_tasks,
        in_progress_tasks=in_progress_tasks,
        completed_tasks=completed_tasks,

        selected_status=selected_status,
        selected_priority=selected_priority
    )


    # =====================================================
    # BASE QUERY
    # =====================================================

    query = MaintenanceReport.query


    # =====================================================
    # STATUS FILTER
    # =====================================================

    if selected_status != "All":

        query = query.filter(
            MaintenanceReport.status == selected_status
        )


    # =====================================================
    # PRIORITY FILTER
    # =====================================================

    if selected_priority != "All":

        query = query.filter(
            MaintenanceReport.priority == selected_priority
        )


    # =====================================================
    # GET TASKS
    # =====================================================

    tasks = query.order_by(
        MaintenanceReport.created_at.desc()
    ).all()


    # =====================================================
    # STATISTICS
    # =====================================================

    pending_tasks = MaintenanceReport.query.filter_by(
        status="Pending"
    ).count()


    assigned_tasks = MaintenanceReport.query.filter_by(
        status="Assigned"
    ).count()


    in_progress_tasks = MaintenanceReport.query.filter_by(
        status="In Progress"
    ).count()


    completed_tasks = MaintenanceReport.query.filter_by(
        status="Completed"
    ).count()


    # =====================================================
    # RENDER
    # =====================================================

    return render_template(
        "supervisor_tasks.html",

        tasks=tasks,

        pending_tasks=pending_tasks,
        assigned_tasks=assigned_tasks,
        in_progress_tasks=in_progress_tasks,
        completed_tasks=completed_tasks,

        selected_status=selected_status,
        selected_priority=selected_priority
    )





# =========================================================
# ASSIGN TASK TO WORKER
# =========================================================

@app.route(
    "/supervisor_tasks/assign/<int:report_id>",
    methods=["POST"]
)
def assign_task_to_worker(report_id):

    # Get the maintenance report
    task = MaintenanceReport.query.get_or_404(report_id)

    # Get selected worker ID
    worker_id = request.form.get("worker_id")

    if not worker_id:
        flash("Please select a worker.", "error")
        return redirect(url_for("supervisor_tasks"))

    # Find worker
    worker = Worker.query.get(worker_id)

    if not worker:
        flash("Worker not found.", "error")
        return redirect(url_for("supervisor_tasks"))

    # Don't assign a completed task
    if task.status == "Completed":
        flash("This task is already completed.", "error")
        return redirect(url_for("supervisor_tasks"))

    # If task was already assigned to another worker,
    # decrease the old worker's current task count
    if task.assigned_worker:

        old_worker = Worker.query.filter_by(
            name=task.assigned_worker
        ).first()

        if old_worker and old_worker.id != worker.id:

            if old_worker.current_tasks and old_worker.current_tasks > 0:
                old_worker.current_tasks -= 1

            # If old worker has no other tasks, make available
            if old_worker.current_tasks == 0:
                old_worker.availability = "Available"

    # Assign new worker
    task.assigned_worker = worker.name

    # Change task status
    task.status = "Assigned"

    # Increase worker task count
    worker.current_tasks = (worker.current_tasks or 0) + 1

    # Worker becomes working
    worker.availability = "Working"

    # Save
    db.session.commit()

    flash(
        f"Task #{task.id} assigned to {worker.name}.",
        "success"
    )

    return redirect(url_for("supervisor_tasks"))



# =========================================================
# UPDATE TASK STATUS
# =========================================================

@app.route(
    "/supervisor_tasks/update/<int:report_id>",
    methods=["POST"]
)
def update_task_status(report_id):

    task = MaintenanceReport.query.get_or_404(
        report_id
    )


    new_status = request.form.get(
        "status"
    )


    allowed_statuses = [
        "Pending",
        "Assigned",
        "In Progress",
        "Completed"
    ]


    if new_status not in allowed_statuses:

        return "Invalid task status.", 400


    task.status = new_status


    db.session.commit()


    return redirect(
        url_for("supervisor_tasks")
    )

# =========================================================
# SUPERVISOR PROFILE
# =========================================================
@app.route("/supervisor_profile", methods=["GET", "POST"])
def supervisor_profile():

    if not session.get("logged_in") or session.get("role") != "supervisor":
        return redirect(url_for("login"))

    user = User.query.filter_by(id=session.get("user_id")).first()
    if not user:
        session.clear()
        return redirect(url_for("login"))

    return render_template(
        "supervisor_profile.html",
        supervisor_name=user.worker.name if user.worker else user.username,
        supervisor_username=user.username,
        supervisor_email=session.get("email", "")
    )

@app.route("/supervisor_update_profile", methods=["POST"])
def supervisor_update_profile():

    if not session.get("logged_in") or session.get("role") != "supervisor":
        return redirect(url_for("login"))

    user = User.query.filter_by(id=session.get("user_id")).first()
    if not user:
        session.clear()
        return redirect(url_for("login"))

    username = request.form.get("username", "").strip()
    email = request.form.get("email", "").strip()
    phone = request.form.get("phone", "").strip()
    full_name = request.form.get("full_name", "").strip()

    if username and username != user.username:
        existing_user = User.query.filter(
            User.username == username,
            User.id != user.id
        ).first()
        if existing_user:
            flash("This username is already in use.", "error")
            return redirect(url_for("supervisor_profile"))
        user.username = username
        session["username"] = username

    if user.worker:
        user.worker.name = full_name or user.worker.name
        user.worker.phone = phone or user.worker.phone
    session["email"] = email
    session["phone"] = phone
    db.session.commit()

    flash("Profile updated successfully.", "success")

    return redirect(url_for("supervisor_profile"))


@app.route("/supervisor_change_password", methods=["POST"])
def supervisor_change_password():

    if not session.get("logged_in") or session.get("role") != "supervisor":
        return redirect(url_for("login"))

    user = User.query.filter_by(id=session.get("user_id")).first()
    if not user:
        session.clear()
        return redirect(url_for("login"))

    current_password = request.form.get("current_password", "").strip()
    new_password = request.form.get("new_password", "").strip()
    confirm_password = request.form.get("confirm_password", "").strip()

    # -----------------------------------------------------
    # Check login
    # -----------------------------------------------------

    if not session.get("logged_in"):
        return redirect(url_for("login"))

    # -----------------------------------------------------
    # Check fields
    # -----------------------------------------------------

    if not current_password or not new_password or not confirm_password:
        return redirect(
            url_for(
                "supervisor_profile",
                error="Please fill in all password fields."
            )
        )

    # -----------------------------------------------------
    # Check password confirmation
    # -----------------------------------------------------

    if new_password != confirm_password:
        return redirect(
            url_for(
                "supervisor_profile",
                error="New passwords do not match."
            )
        )

    # -----------------------------------------------------
    # Check password length
    # -----------------------------------------------------

    if len(new_password) < 6:
        return redirect(
            url_for(
                "supervisor_profile",
                error="Password must contain at least 6 characters."
            )
        )

    if not check_password_hash(
        user.password,
        current_password
    ):

        return redirect(
            url_for(
                "supervisor_profile",
                error="Current password is incorrect."
            )
        )

    # -----------------------------------------------------
    # Update password
    # -----------------------------------------------------

    user.password = generate_password_hash(
        new_password
    )
    user.must_change_password = False
    db.session.commit()

    return redirect(
        url_for(
            "supervisor_profile",
            success="Password changed successfully."
        )
    )



@app.route("/supervisor-reports/<int:report_id>")
def supervisor_report_details(report_id):

    report = MaintenanceReport.query.get_or_404(report_id)

    return render_template(
        "report_details.html",
        report=report
    )


@app.route("/supervisor_reports")
def supervisor_reports():

    reports = MaintenanceReport.query.order_by(
        MaintenanceReport.created_at.desc()
    ).all()

    return render_template(
        "supervisor_reports.html",
        reports=reports
    )





# =========================================================
# WORKERS
# =========================================================

@app.route("/workers")
def workers():

    workers_list = Worker.query.order_by(
        Worker.name.asc()
    ).all()

    total_workers = len(workers_list)

    available_workers = sum(
        1 for worker in workers_list
        if worker.availability == "Available"
    )

    working_workers = sum(
        1 for worker in workers_list
        if worker.availability == "Working"
    )

    offline_workers = sum(
        1 for worker in workers_list
        if worker.availability == "Offline"
    )

    return render_template(
        "workers.html",
        workers=workers_list,
        total_workers=total_workers,
        available_workers=available_workers,
        working_workers=working_workers,
        offline_workers=offline_workers
    )


# =========================================================
# WORKER DETAILS
# =========================================================

@app.route("/worker_details/<int:worker_id>")
def worker_details(worker_id):

    worker = Worker.query.get_or_404(worker_id)

    return render_template(
        "worker_details.html",
        worker=worker
    )

# =========================================================
# ADMIN DASHBOARD
# =========================================================

@app.route("/admin_dashboard")
def admin_dashboard():

    # =====================================================
    # WORKERS
    # =====================================================

    workers_list = Worker.query.all()

    total_workers = len(workers_list)

    available_workers = sum(
        1 for worker in workers_list
        if worker.availability == "Available"
    )

    working_workers = sum(
        1 for worker in workers_list
        if worker.availability == "Working"
    )

    offline_workers = sum(
        1 for worker in workers_list
        if worker.availability == "Offline"
    )


    # =====================================================
    # MAINTENANCE REPORTS
    # =====================================================

    reports = MaintenanceReport.query.all()

    total_reports = len(reports)

    pending_reports = sum(
        1 for report in reports
        if report.status == "Pending"
    )

    in_progress_reports = sum(
        1 for report in reports
        if report.status == "In Progress"
    )

    completed_reports = sum(
        1 for report in reports
        if report.status == "Completed"
    )


    # =====================================================
    # RECENT REPORTS
    # =====================================================

    recent_reports = MaintenanceReport.query.order_by(
        MaintenanceReport.id.desc()
    ).limit(5).all()


    # =====================================================
    # SEND DATA TO HTML
    # =====================================================

    return render_template(
        "admin_dashboard.html",

        total_reports=total_reports,
        pending_reports=pending_reports,
        in_progress_reports=in_progress_reports,
        completed_reports=completed_reports,

        total_workers=total_workers,
        available_workers=available_workers,
        working_workers=working_workers,
        offline_workers=offline_workers,

        recent_reports=recent_reports
    )




# =========================================================
# ADMIN - WORKERS
# =========================================================

@app.route("/admin_workers")
def admin_workers():

    workers_list = Worker.query.order_by(
        Worker.name.asc()
    ).all()

    total_workers = len(workers_list)

    available_workers = sum(
        1 for worker in workers_list
        if worker.availability == "Available"
    )

    working_workers = sum(
        1 for worker in workers_list
        if worker.availability == "Working"
    )

    busy_workers = sum(
        1 for worker in workers_list
        if worker.availability == "Busy"
    )

    offline_workers = sum(
        1 for worker in workers_list
        if worker.availability == "Offline"
    )

    return render_template(
        "admin_workers.html",
        workers=workers_list,
        total_workers=total_workers,
        available_workers=available_workers,
        working_workers=working_workers,
        busy_workers=busy_workers,
        offline_workers=offline_workers
    )


# =========================================================
# ADMIN - ADD WORKER PAGE
# =========================================================

@app.route("/admin_add_workers")
def admin_add_workers():

    return render_template(
        "admin_add_workers.html"
    )


# =========================================================
# ADMIN - SAVE NEW WORKER
# =========================================================

@app.route("/admin_add_worker", methods=["POST"])
def admin_add_worker():

    name = request.form.get("name")
    phone = request.form.get("phone")
    department = request.form.get("department")
    skills = request.form.get("skills")
    availability = request.form.get("availability")

    current_tasks = request.form.get("current_tasks")

    # Convert current tasks to integer
    if not current_tasks:
        current_tasks = 0
    else:
        current_tasks = int(current_tasks)


    # Create new worker
    new_worker = Worker(

        name=name,

        phone=phone,

        department=department,

        skills=skills,

        availability=availability,

        current_tasks=current_tasks

    )


    # Save worker to MySQL

    db.session.add(new_worker)

    db.session.commit()


    # Return to Admin Workers page

    return redirect(
        url_for("admin_workers")
    )

# =========================================================
# NOTIFICATIONS
# =========================================================

@app.route("/notifications")
def notifications():
    return render_template("notifications.html")


# =========================================================
# ADMIN SETUP
# =========================================================

@app.route("/admin_setup")
def admin_setup():

    hotel = Hotel.query.first()

    workers_list = Worker.query.order_by(
        Worker.name.asc()
    ).all()

    total_workers = len(workers_list)

    available_workers = sum(
        1 for worker in workers_list
        if worker.availability == "Available"
    )

    busy_workers = sum(
        1 for worker in workers_list
        if worker.availability == "Busy"
    )

    offline_workers = sum(
        1 for worker in workers_list
        if worker.availability == "Offline"
    )

    return render_template(
        "admin_setup.html",
        hotel=hotel,
        workers=workers_list,
        total_workers=total_workers,
        available_workers=available_workers,
        busy_workers=busy_workers,
        offline_workers=offline_workers
    )



#------------------
@app.route("/admin/save_hotel", methods=["POST"])
def save_hotel():

    hotel_name = request.form.get("hotel_name")
    address = request.form.get("address")
    floors = request.form.get("floors")
    rooms = request.form.get("rooms")

    if not hotel_name or not address or not floors or not rooms:
        return "Please fill in all fields."

    try:
        floors = int(floors)
        rooms = int(rooms)
    except ValueError:
        return "Floors and rooms must be numbers."

    hotel = Hotel.query.first()

    if hotel:

        hotel.hotel_name = hotel_name
        hotel.address = address
        hotel.floors = floors
        hotel.rooms = rooms

    else:

        hotel = Hotel(
            hotel_name=hotel_name,
            address=address,
            floors=floors,
            rooms=rooms
        )

        db.session.add(hotel)

    db.session.commit()

    return redirect(url_for("admin_setup"))



    # =====================================================
    # AADMIN ROOM
    # =====================================================
@app.route("/admin_rooms", methods=["GET", "POST"])
def admin_rooms():

    hotel = Hotel.query.first()

    # =====================================================
    # AUTOMATIC ROOM GENERATION
    # =====================================================

    if request.method == "POST":

        floors = request.form.get("floors")
        rooms_per_floor = request.form.get("rooms_per_floor")

        if not floors or not rooms_per_floor:
            return "Please enter floors and rooms per floor."

        try:
            floors = int(floors)
            rooms_per_floor = int(rooms_per_floor)

        except ValueError:
            return "Floors and rooms per floor must be numbers."

        # Generate rooms
        for floor in range(1, floors + 1):

            for room_index in range(1, rooms_per_floor + 1):

                room_number = f"{floor}{room_index:02d}"

                # Check if room already exists
                existing_room = Room.query.filter_by(
                    room_number=room_number
                ).first()

                if not existing_room:

                    new_room = Room(
                        room_number=room_number,
                        floor_number=floor,
                        room_type="Standard",
                        status="Available"
                    )

                    db.session.add(new_room)

        db.session.commit()

        return redirect(url_for("admin_rooms"))

    # =====================================================
    # GET ROOMS
    # =====================================================

    rooms = Room.query.order_by(
        Room.floor_number.asc(),
        Room.room_number.asc()
    ).all()

    # =====================================================
    # HOTEL INFORMATION
    # =====================================================

    if hotel:

        total_floors = hotel.floors

        # Total rooms configured for the hotel
        total_rooms = hotel.rooms

    else:

        total_floors = 0
        total_rooms = 0

    # =====================================================
    # REAL ROOM RECORDS
    # =====================================================

    configured_rooms = Room.query.count()

    available_rooms = Room.query.filter_by(
        status="Available"
    ).count()

    occupied_rooms = Room.query.filter_by(
        status="Occupied"
    ).count()

    maintenance_rooms = Room.query.filter_by(
        status="Maintenance"
    ).count()

    # =====================================================
    # ROOM CONFIGURATION PROGRESS
    # =====================================================

    if total_rooms > 0:

        room_percentage = round(
            (configured_rooms / total_rooms) * 100
        )

    else:

        room_percentage = 0

    # Prevent percentage from going above 100
    if room_percentage > 100:
        room_percentage = 100

    # Room configuration percentage
    if total_rooms > 0:
     room_percentage = round(
        (configured_rooms / total_rooms) * 100
    )
    else:
     room_percentage = 0

    # =====================================================
    # RENDER
    # =====================================================

    return render_template(
        "admin_rooms.html",

        hotel=hotel,
        rooms=rooms,

        total_floors=total_floors,
        total_rooms=total_rooms,

        configured_rooms=configured_rooms,
        available_rooms=available_rooms,

        occupied_rooms=occupied_rooms,
        maintenance_rooms=maintenance_rooms,

        room_percentage=room_percentage
    )

    # =====================================================
    # GET ROOMS FROM DATABASE
    # =====================================================

    rooms = Room.query.order_by(
        Room.floor_number.asc(),
        Room.room_number.asc()
    ).all()

    # =====================================================
    # HOTEL INFORMATION
    # =====================================================

    total_floors = hotel.floors if hotel else 0

    # =====================================================
    # REAL ROOM STATISTICS
    # =====================================================

    # Number of rooms actually stored in the rooms table
    total_rooms = Room.query.count()

    # Configured rooms = rooms actually created
    configured_rooms = Room.query.count()

    # Available rooms
    available_rooms = Room.query.filter_by(
        status="Available"
    ).count()

    # Occupied rooms
    occupied_rooms = Room.query.filter_by(
        status="Occupied"
    ).count()

    # Maintenance rooms
    maintenance_rooms = Room.query.filter_by(
        status="Maintenance"
    ).count()

    # =====================================================
    # ROOM CONFIGURATION PERCENTAGE
    # =====================================================

    if total_rooms > 0:
        room_percentage = round(
            (configured_rooms / total_rooms) * 100
        )
    else:
        room_percentage = 0

    # =====================================================
    # RENDER PAGE
    # =====================================================

    return render_template(
        "admin_rooms.html",

        hotel=hotel,
        rooms=rooms,

        total_floors=total_floors,
        total_rooms=total_rooms,
        configured_rooms=configured_rooms,
        available_rooms=available_rooms,

        occupied_rooms=occupied_rooms,
        maintenance_rooms=maintenance_rooms,

        room_percentage=room_percentage
    )
# =========================================================
# ADMIN - ADD ROOM PAGE
# =========================================================

@app.route("/admin_add_room", methods=["GET"])
def admin_add_room_page():
    return render_template("admin_add_room.html")



# =========================================================
# ADMIN - STAFF ACCOUNTS
# =========================================================

@app.route("/admin_staff_accounts")
def admin_staff_accounts():

    # Only logged-in admin can access
    if not session.get("logged_in"):
        return redirect(url_for("login"))

    if session.get("role") != "admin":
        return redirect(url_for("login"))

    users = User.query.order_by(
        User.created_at.desc()
    ).all()

    workers = Worker.query.order_by(
        Worker.name.asc()
    ).all()

    return render_template(
        "admin_staff_accounts.html",
        users=users,
        workers=workers
    )




# =========================================================
# ADMIN - CREATE STAFF ACCOUNT
# =========================================================

@app.route(
    "/admin_create_staff",
    methods=["POST"]
)
def admin_create_staff():

    # Check login
    if not session.get("logged_in"):
        return redirect(url_for("login"))

    # Only admin
    if session.get("role") != "admin":
        return redirect(url_for("login"))

    username = request.form.get(
        "username",
        ""
    ).strip()

    password = request.form.get(
        "password",
        ""
    )

    role = request.form.get(
        "role",
        ""
    ).strip().lower()

    worker_id = request.form.get(
        "worker_id"
    )

    # -----------------------------------------------------
    # VALIDATION
    # -----------------------------------------------------

    if not username or not password or not role:

        flash(
            "Please fill in all required fields.",
            "error"
        )

        return redirect(
            url_for("admin_staff_accounts")
        )

    # Only these roles can be created here
    if role not in ["technician", "supervisor"]:

        flash(
            "Invalid staff role.",
            "error"
        )

        return redirect(
            url_for("admin_staff_accounts")
        )

    # Password length
    if len(password) < 6:

        flash(
            "Password must contain at least 6 characters.",
            "error"
        )

        return redirect(
            url_for("admin_staff_accounts")
        )

    # -----------------------------------------------------
    # CHECK USERNAME
    # -----------------------------------------------------

    existing_user = User.query.filter_by(
        username=username
    ).first()

    if existing_user:

        flash(
            "This username already exists.",
            "error"
        )

        return redirect(
            url_for("admin_staff_accounts")
        )

    # -----------------------------------------------------
    # TECHNICIAN WORKER
    # -----------------------------------------------------

    selected_worker = None

    if role == "technician":

        if not worker_id:

            flash(
                "Please select a worker for the technician account.",
                "error"
            )

            return redirect(
                url_for("admin_staff_accounts")
            )

        try:
            worker_id = int(worker_id)

        except ValueError:

            flash(
                "Invalid worker selected.",
                "error"
            )

            return redirect(
                url_for("admin_staff_accounts")
            )

        selected_worker = Worker.query.get(
            worker_id
        )

        if not selected_worker:

            flash(
                "Selected worker does not exist.",
                "error"
            )

            return redirect(
                url_for("admin_staff_accounts")
            )

        # Check if worker already has an account
        existing_worker_account = User.query.filter_by(
            worker_id=worker_id
        ).first()

        if existing_worker_account:

            flash(
                "This worker already has a login account.",
                "error"
            )

            return redirect(
                url_for("admin_staff_accounts")
            )

    # -----------------------------------------------------
    # SUPERVISOR
    # -----------------------------------------------------

    if role == "supervisor":

        worker_id = None

    # -----------------------------------------------------
    # CREATE USER
    # -----------------------------------------------------

    hashed_password = generate_password_hash(
        password
    )

    new_user = User(
        username=username,
        password=hashed_password,
        role=role,
        worker_id=worker_id
    )

    db.session.add(new_user)

    db.session.commit()

    flash(
        f"Account '{username}' created successfully.",
        "success"
    )

    return redirect(
        url_for("admin_staff_accounts")
    )

# =========================================================
# ADMIN - SAVE NEW ROOM
# =========================================================

@app.route("/admin_add_room", methods=["POST"])
def admin_add_room():

    room_number = request.form.get("room_number")
    floor_number = request.form.get("floor_number")
    room_type = request.form.get("room_type")
    status = request.form.get("status")

    # Validate required fields
    if not room_number or not floor_number:
        return "Room number and floor number are required."

    try:
        floor_number = int(floor_number)
    except ValueError:
        return "Floor number must be a number."

    # Check duplicate room
    existing_room = Room.query.filter_by(
        room_number=room_number
    ).first()

    if existing_room:
        return "This room already exists."

    # Create room
    new_room = Room(
        room_number=room_number,
        floor_number=floor_number,
        room_type=room_type or "Standard",
        status=status or "Available"
    )

    db.session.add(new_room)
    db.session.commit()

    return redirect(url_for("admin_rooms"))


# =========================================================
# ADMIN IMPORT
# =========================================================

@app.route("/admin_import")
def admin_import():

    return render_template("admin_import.html")


# =========================================================
# IMPORT EXCEL DATA
# =========================================================

@app.route("/admin/import_excel", methods=["POST"])
def import_excel():

    # -----------------------------------------------------
    # Check file
    # -----------------------------------------------------

    if "excel_file" not in request.files:

        return "No Excel file was uploaded."

    file = request.files["excel_file"]

    if file.filename == "":

        return "Please select an Excel file."

    if not allowed_file(file.filename):

        return "Only .xlsx and .xls files are allowed."

    # -----------------------------------------------------
    # Save uploaded file
    # -----------------------------------------------------

    filename = secure_filename(file.filename)

    filepath = os.path.join(
        app.config["UPLOAD_FOLDER"],
        filename
    )

    file.save(filepath)

    try:

        # -------------------------------------------------
        # Read Excel workbook
        # -------------------------------------------------

        excel_file = pd.ExcelFile(filepath)

        imported_rooms = 0
        imported_workers = 0

        # =================================================
        # ROOMS SHEET
        # =================================================

        if "Rooms" in excel_file.sheet_names:

            rooms_df = pd.read_excel(
                filepath,
                sheet_name="Rooms"
            )

            # Remove empty rows
            rooms_df = rooms_df.dropna(how="all")

            for _, row in rooms_df.iterrows():

                room_number = str(
                    row.get("Room Number", "")
                ).strip()

                floor_number = row.get(
                    "Floor",
                    ""
                )

                room_type = str(
                    row.get(
                        "Room Type",
                        "Standard"
                    )
                ).strip()

                status = str(
                    row.get(
                        "Status",
                        "Available"
                    )
                ).strip()

                # Skip invalid rows
                if not room_number or not floor_number:

                    continue

                try:

                    floor_number = int(
                        float(floor_number)
                    )

                except:

                    continue

                # Check duplicate
                existing_room = Room.query.filter_by(
                    room_number=room_number
                ).first()

                if existing_room:

                    continue

                new_room = Room(

                    room_number=room_number,

                    floor_number=floor_number,

                    room_type=room_type,

                    status=status

                )

                db.session.add(new_room)

                imported_rooms += 1

        # =================================================
        # WORKERS SHEET
        # =================================================

        if "Workers" in excel_file.sheet_names:

            workers_df = pd.read_excel(
                filepath,
                sheet_name="Workers"
            )

            workers_df = workers_df.dropna(
                how="all"
            )

            for _, row in workers_df.iterrows():

                name = str(
                    row.get("Name", "")
                ).strip()

                phone = str(
                    row.get("Phone", "")
                ).strip()

                department = str(
                    row.get("Department", "")
                ).strip()

                skills = str(
                    row.get("Skills", "")
                ).strip()

                availability = str(
                    row.get(
                        "Availability",
                        "Available"
                    )
                ).strip()

                # Skip empty worker
                if not name:

                    continue

                # Check duplicate worker
                existing_worker = Worker.query.filter_by(
                    name=name
                ).first()

                if existing_worker:

                    continue

                new_worker = Worker(

                    name=name,

                    phone=phone,

                    department=department,

                    skills=skills,

                    availability=availability,

                    current_tasks=0

                )

                db.session.add(new_worker)

                imported_workers += 1

        # =================================================
        # SAVE EVERYTHING
        # =================================================

        db.session.commit()

        # -------------------------------------------------
        # Delete uploaded file
        # -------------------------------------------------

        try:

            os.remove(filepath)

        except:

            pass

        # -------------------------------------------------
        # Success
        # -------------------------------------------------

        return f"""
        <!DOCTYPE html>

        <html>

        <head>

            <title>Import Successful</title>

            <style>

                body {{
                    font-family: Arial, sans-serif;
                    background: #f5f7fb;
                    display: flex;
                    justify-content: center;
                    align-items: center;
                    height: 100vh;
                    margin: 0;
                }}

                .success {{
                    background: white;
                    padding: 40px;
                    border-radius: 16px;
                    box-shadow: 0 10px 30px rgba(0,0,0,0.08);
                    text-align: center;
                    max-width: 500px;
                }}

                h1 {{
                    color: #16a34a;
                }}

                .number {{
                    font-size: 20px;
                    margin: 15px 0;
                }}

                a {{
                    display: inline-block;
                    margin-top: 20px;
                    padding: 12px 24px;
                    background: #2563eb;
                    color: white;
                    text-decoration: none;
                    border-radius: 8px;
                }}

            </style>

        </head>

        <body>

            <div class="success">

                <h1>✓ Import Successful</h1>

                <div class="number">
                    🚪 Rooms imported:
                    <strong>{imported_rooms}</strong>
                </div>

                <div class="number">
                    👨‍🔧 Workers imported:
                    <strong>{imported_workers}</strong>
                </div>

                <a href="/admin_import">
                    Back to Import
                </a>

                <a href="/admin_rooms">
                    View Rooms
                </a>

            </div>

        </body>

        </html>
        """

    except Exception as e:

        # Rollback database
        db.session.rollback()

        # Remove uploaded file
        try:

            os.remove(filepath)

        except:

            pass

        return f"""
        <h2>Import Error</h2>

        <p>{str(e)}</p>

        <a href="/admin_import">
            Back to Import
        </a>
        """




# =========================================================
# ADMIN - RESET ENTIRE SYSTEM
# =========================================================

def reset_database():
    table_rows = db.session.execute(text("SHOW TABLES")).all()
    table_names = [row[0] for row in table_rows]

    auto_increment_rows = db.session.execute(
        text(
            "SELECT table_name "
            "FROM information_schema.tables "
            "WHERE table_schema = DATABASE() "
            "AND auto_increment IS NOT NULL"
        )
    ).all()
    auto_increment_tables = {row[0] for row in auto_increment_rows}

    try:
        db.session.execute(text("SET FOREIGN_KEY_CHECKS = 0"))

        for table_name in table_names:
            quoted_table = table_name.replace("`", "``")
            db.session.execute(text(f"DELETE FROM `{quoted_table}`"))

        for table_name in auto_increment_tables:
            quoted_table = table_name.replace("`", "``")
            db.session.execute(
                text(f"ALTER TABLE `{quoted_table}` AUTO_INCREMENT = 1")
            )

        default_admin = User(
            username="admin",
            password=generate_password_hash("admin123"),
            role="admin",
            worker_id=None,
            must_change_password=False
        )
        db.session.add(default_admin)
        db.session.commit()
        return default_admin
    except Exception:
        db.session.rollback()
        raise
    finally:
        db.session.execute(text("SET FOREIGN_KEY_CHECKS = 1"))

@app.route("/admin/reset-system", methods=["POST"])
def admin_reset_system():

    # Only logged-in admin
    if not session.get("logged_in"):
        return redirect(url_for("login"))

    if session.get("role") != "admin":
        return redirect(url_for("login"))

    try:
        admin = reset_database()

        session.clear()

        return redirect(
            url_for("login")
        )

    except Exception as e:

        db.session.rollback()

        print("\n====================================")
        print("SYSTEM RESET ERROR")
        print("====================================")
        print(e)
        print("====================================\n")

        return "System reset failed. Check the Flask console.", 500

# =========================================================
# SUPERVISOR AI ANALYSIS
# =========================================================

def supervisor_ai_analysis():

    reports = MaintenanceReport.query.all()
    workers = Worker.query.all()

    # =====================================================
    # REPORT STATISTICS
    # =====================================================

    total_reports = len(reports)

    pending_reports = [
        report for report in reports
        if report.status == "Pending"
    ]

    assigned_reports = [
        report for report in reports
        if report.status == "Assigned"
    ]

    in_progress_reports = [
        report for report in reports
        if report.status == "In Progress"
    ]

    completed_reports = [
        report for report in reports
        if report.status == "Completed"
    ]

    urgent_reports = [
        report for report in reports
        if report.priority == "Urgent"
    ]

    high_reports = [
        report for report in reports
        if report.priority == "High"
    ]

    # =====================================================
    # WORKER STATISTICS
    # =====================================================

    available_workers = [
        worker for worker in workers
        if (worker.availability or "").lower()
        == "available"
    ]

    working_workers = [
        worker for worker in workers
        if (worker.availability or "").lower()
        == "working"
    ]

    # =====================================================
    # OVERLOADED WORKERS
    # =====================================================

    overloaded_workers = [
        worker for worker in workers
        if (worker.current_tasks or 0) >= 3
    ]

    # =====================================================
    # AI RECOMMENDATIONS
    # =====================================================

    recommendations = []

    # -----------------------------------------------------
    # Urgent tasks
    # -----------------------------------------------------

    if urgent_reports:

        recommendations.append(
            f"⚠️ There are {len(urgent_reports)} "
            f"urgent maintenance task(s) requiring "
            f"immediate attention."
        )

    # -----------------------------------------------------
    # High priority pending
    # -----------------------------------------------------

    high_pending = [
        report for report in pending_reports
        if report.priority == "High"
    ]

    if high_pending:

        recommendations.append(
            f"🔴 {len(high_pending)} high-priority "
            f"task(s) are waiting for assignment."
        )

    # -----------------------------------------------------
    # No available workers
    # -----------------------------------------------------

    if pending_reports and not available_workers:

        recommendations.append(
            "⚠️ There are pending reports but no "
            "technicians are currently available."
        )

    # -----------------------------------------------------
    # Overloaded workers
    # -----------------------------------------------------

    if overloaded_workers:

        names = ", ".join(
            worker.name
            for worker in overloaded_workers
        )

        recommendations.append(
            f"⚠️ Workload is high for: {names}."
        )

    # -----------------------------------------------------
    # Normal situation
    # -----------------------------------------------------

    if not recommendations:

        recommendations.append(
            "✅ Maintenance operations are currently "
            "under control."
        )

    # =====================================================
    # RETURN ANALYSIS
    # =====================================================

    return {
        "total_reports": total_reports,

        "pending_reports": pending_reports,

        "assigned_reports": assigned_reports,

        "in_progress_reports": in_progress_reports,

        "completed_reports": completed_reports,

        "urgent_reports": urgent_reports,

        "high_reports": high_reports,

        "available_workers": available_workers,

        "working_workers": working_workers,

        "overloaded_workers": overloaded_workers,

        "recommendations": recommendations
    }



# =========================================================
# SUPERVISOR AI
# =========================================================

@app.route("/supervisor_ai")
def supervisor_ai():

    # -----------------------------------------------------
    # LOGIN CHECK
    # -----------------------------------------------------

    if not session.get("logged_in"):

        return redirect(
            url_for("login")
        )

    # -----------------------------------------------------
    # SUPERVISOR ONLY
    # -----------------------------------------------------

    if session.get("role") != "supervisor":

        return redirect(
            url_for("login")
        )

    # -----------------------------------------------------
    # RUN AI ANALYSIS
    # -----------------------------------------------------

    analysis = supervisor_ai_analysis()

    # -----------------------------------------------------
    # DISPLAY RESULTS
    # -----------------------------------------------------

    return render_template(
        "supervisor_ai.html",
        analysis=analysis
    )


@app.route("/report_details/<int:report_id>")
def report_details(report_id):

    report_details = MaintenanceReport.query.get_or_404(report_id)

    return render_template(
        "report_details.html",
        report=report_details
    )





@app.route("/download_template")
def download_template():

    wb = Workbook()

    # =========================
    # ROOMS SHEET
    # =========================

    rooms_sheet = wb.active
    rooms_sheet.title = "Rooms"

    rooms_sheet.append([
        "Floor",
        "Room Number",
        "Room Type",
        "Status"
    ])

    rooms_sheet.append([
        1,
        "101",
        "Standard",
        "Available"
    ])

    rooms_sheet.append([
        1,
        "102",
        "Deluxe",
        "Available"
    ])

    rooms_sheet.append([
        2,
        "201",
        "Suite",
        "Occupied"
    ])

    # =========================
    # WORKERS SHEET
    # =========================

    workers_sheet = wb.create_sheet("Workers")

    workers_sheet.append([
        "Name",
        "Phone",
        "Department",
        "Skills",
        "Availability"
    ])

    workers_sheet.append([
        "Ahmed",
        "0612345678",
        "Maintenance",
        "Plumbing",
        "Available"
    ])

    workers_sheet.append([
        "Youssef",
        "0623456789",
        "Electrical",
        "Electrical",
        "Working"
    ])

    # =========================
    # SAVE FILE
    # =========================

    file_path = "hotel_import_template.xlsx"

    wb.save(file_path)

    return send_file(
        file_path,
        as_attachment=True,
        download_name="hotel_import_template.xlsx"
    )



# =========================================================
# ADMIN WORKER DETAILS
# =========================================================

@app.route("/admin_workers_details/<int:worker_id>")
def admin_workers_details(worker_id):

    worker = Worker.query.get_or_404(worker_id)

    return render_template(
        "admin_workers_details.html",
        worker=worker
    )


# =====================================================
# ADMIN SETTINGS
# ========================


@app.route("/admin_settings")
def admin_settings():

    hotel = Hotel.query.first()

    total_rooms = Room.query.count()

    total_workers = Worker.query.count()

    total_reports = MaintenanceReport.query.count()

    settings = {
        "email": "",
        "new_reports": True,
        "urgent_reports": True,
        "worker_notifications": True,
        "ai_enabled": True,
        "assignment_strategy": "balanced",
        "default_room_status": "Available",
        "default_priority": "Pending"
    }

    return render_template(
        "admin_settings.html",
        hotel=hotel,
        total_rooms=total_rooms,
        total_workers=total_workers,
        total_reports=total_reports,
        settings=settings
    )



@app.route("/admin_settings/save", methods=["POST"])
def admin_save_settings():

    settings = session.get("admin_settings", {})

    section = request.form.get("section")


    # =====================================================
    # ACCOUNT
    # =====================================================

    if section == "account":

        username = request.form.get("username")
        email = request.form.get("email")

        if username:
            session["username"] = username

        settings["email"] = email or ""


    # =====================================================
    # NOTIFICATIONS
    # =====================================================

    elif section == "notifications":

        settings["new_reports"] = (
            "new_reports" in request.form
        )

        settings["urgent_reports"] = (
            "urgent_reports" in request.form
        )

        settings["worker_notifications"] = (
            "worker_notifications" in request.form
        )


    # =====================================================
    # AI SETTINGS
    # =====================================================

    elif section == "ai":

        settings["ai_enabled"] = (
            "ai_enabled" in request.form
        )

        settings["assignment_strategy"] = request.form.get(
            "assignment_strategy",
            "balanced"
        )


    # =====================================================
    # SYSTEM
    # =====================================================

    elif section == "system":

        settings["default_room_status"] = request.form.get(
            "default_room_status",
            "Available"
        )

        settings["default_priority"] = request.form.get(
            "default_priority",
            "Pending"
        )


    session["admin_settings"] = settings

    return redirect(
        url_for(
            "admin_settings",
            success="Settings saved successfully."
        )
    )





@app.route("/admin_settings/reset", methods=["POST"])
def admin_reset_settings():

    if not session.get("logged_in") or session.get("role") != "admin":
        return redirect(url_for("login"))

    try:
        reset_database()

    except Exception:
        db.session.rollback()
        return redirect(
            url_for(
                "admin_settings",
                error="Reset failed. Please check the database connection."
            )
        )

    for filename in os.listdir(UPLOAD_FOLDER):
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        if os.path.isfile(file_path):
            try:
                os.remove(file_path)
            except OSError:
                pass

    session.clear()
    return redirect(
        url_for(
            "login",
            success="System reset. Sign in with admin and admin123."
        )
    )

@app.route("/supervisor_workers")
def supervisor_workers():
    return render_template("supervisor_workers.html")

# =========================================================
# TEST DATABASE
# =========================================================

@app.route("/test")
def test():

    try:

        # Test database connection
        reports = MaintenanceReport.query.all()

        return f"""
            <h1>AI Hotel Maintenance System</h1>
            <p>Flask is working correctly.</p>
            <p>Database connection: SUCCESS</p>
            <p>Number of reports: {len(reports)}</p>
        """

    except Exception as e:

        return f"""
            <h1>Database Error</h1>
            <p>{e}</p>
        """



if __name__ == "__main__":

    app.run(debug=True)
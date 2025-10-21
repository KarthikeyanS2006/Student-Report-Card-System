import os
import sys
import csv
import sqlite3
import shutil
import smtplib
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, simpledialog
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors


class ReportCardSystem:
    def __init__(self):
        # Create directory structure
        self.data_dir = "student_records"
        self.reports_dir = os.path.join(self.data_dir, "report_cards_pdf")
        self.backup_dir = os.path.join(self.data_dir, "backups")
        self.db_file = os.path.join(self.data_dir, "students.db")
        self.csv_file = os.path.join(self.data_dir, "students.csv")
        
        # Create directories
        os.makedirs(self.data_dir, exist_ok=True)
        os.makedirs(self.reports_dir, exist_ok=True)
        os.makedirs(self.backup_dir, exist_ok=True)
        
        # Initialize database
        self.init_database()
        
        # Migrate CSV to SQLite if CSV exists
        if os.path.exists(self.csv_file):
            self.migrate_csv_to_sqlite()
    
    def init_database(self):
        """Initialize SQLite database"""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS students (
                student_id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                math REAL,
                science REAL,
                english REAL,
                social_studies REAL,
                computer REAL,
                total REAL,
                percentage REAL,
                grade TEXT,
                email TEXT,
                created_date TEXT
            )
        ''')
        
        conn.commit()
        conn.close()
    
    def migrate_csv_to_sqlite(self):
        """Migrate existing CSV data to SQLite"""
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            with open(self.csv_file, 'r') as file:
                reader = csv.DictReader(file)
                for row in reader:
                    try:
                        cursor.execute('''
                            INSERT OR REPLACE INTO students 
                            (student_id, name, math, science, english, social_studies, 
                             computer, total, percentage, grade, created_date)
                            VALUES (?,?,?,?,?,?,?,?,?,?,?)
                        ''', (row['Student_ID'], row['Name'], row['Math'], 
                              row['Science'], row['English'], row['Social_Studies'],
                              row['Computer'], row['Total'], row['Percentage'], 
                              row['Grade'], datetime.now().strftime('%Y-%m-%d')))
                    except:
                        continue
            
            conn.commit()
            conn.close()
            
            # Backup and remove old CSV
            backup_name = f"students_migrated_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            shutil.move(self.csv_file, os.path.join(self.backup_dir, backup_name))
            
        except Exception as e:
            print(f"Migration note: {e}")
    
    def calculate_grade(self, percentage):
        """Calculate grade based on percentage"""
        if percentage >= 90:
            return 'A+'
        elif percentage >= 80:
            return 'A'
        elif percentage >= 70:
            return 'B'
        elif percentage >= 60:
            return 'C'
        elif percentage >= 50:
            return 'D'
        else:
            return 'F'
    
    def add_student(self, student_id, name, math, science, english, social, computer, email=""):
        """Add new student with marks"""
        try:
            marks = [math, science, english, social, computer]
            if any(mark < 0 or mark > 100 for mark in marks):
                return False, "Marks must be between 0 and 100!"
            
            total = sum(marks)
            percentage = total / 5
            grade = self.calculate_grade(percentage)
            created_date = datetime.now().strftime('%Y-%m-%d')
            
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            cursor.execute('''
                INSERT INTO students VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
            ''', (student_id, name, math, science, english, social, computer, 
                  total, round(percentage, 2), grade, email, created_date))
            
            conn.commit()
            conn.close()
            
            return True, f"Student {name} added successfully!"
            
        except sqlite3.IntegrityError:
            return False, "Student ID already exists!"
        except Exception as e:
            return False, str(e)
    
    def get_all_students(self):
        """Return all students as list of dictionaries"""
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            cursor.execute("SELECT * FROM students ORDER BY name")
            columns = [description[0] for description in cursor.description]
            rows = cursor.fetchall()
            
            conn.close()
            
            students = []
            for row in rows:
                student_dict = dict(zip(columns, row))
                students.append(student_dict)
            
            return students
        except:
            return []
    
    def search_student(self, student_id):
        """Search for specific student by ID"""
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            cursor.execute("SELECT * FROM students WHERE student_id = ?", (student_id,))
            columns = [description[0] for description in cursor.description]
            row = cursor.fetchone()
            
            conn.close()
            
            if row:
                return dict(zip(columns, row))
            return None
        except:
            return None
    
    def search_by_name(self, name):
        """Search students by name"""
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            cursor.execute("SELECT * FROM students WHERE name LIKE ?", (f'%{name}%',))
            columns = [description[0] for description in cursor.description]
            rows = cursor.fetchall()
            
            conn.close()
            
            students = []
            for row in rows:
                students.append(dict(zip(columns, row)))
            
            return students
        except:
            return []
    
    def filter_by_grade(self, grade):
        """Filter students by grade"""
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            if grade == "All":
                cursor.execute("SELECT * FROM students ORDER BY percentage DESC")
            else:
                cursor.execute("SELECT * FROM students WHERE grade = ? ORDER BY percentage DESC", (grade,))
            
            columns = [description[0] for description in cursor.description]
            rows = cursor.fetchall()
            
            conn.close()
            
            students = []
            for row in rows:
                students.append(dict(zip(columns, row)))
            
            return students
        except:
            return []
    
    def get_statistics(self):
        """Get class statistics"""
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            # Total students
            cursor.execute("SELECT COUNT(*) FROM students")
            total = cursor.fetchone()[0]
            
            # Average percentage
            cursor.execute("SELECT AVG(percentage) FROM students")
            avg = cursor.fetchone()[0]
            
            # Grade distribution
            cursor.execute("SELECT grade, COUNT(*) FROM students GROUP BY grade")
            grade_dist = dict(cursor.fetchall())
            
            # Highest score
            cursor.execute("SELECT name, percentage FROM students ORDER BY percentage DESC LIMIT 1")
            top_student = cursor.fetchone()
            
            conn.close()
            
            return {
                'total_students': total,
                'average_percentage': round(avg, 2) if avg else 0,
                'grade_distribution': grade_dist,
                'top_student': top_student
            }
        except:
            return None
    
    def create_backup(self):
        """Create backup of database"""
        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            backup_file = os.path.join(self.backup_dir, f"students_backup_{timestamp}.db")
            
            shutil.copy2(self.db_file, backup_file)
            
            return True, f"Backup created: students_backup_{timestamp}.db"
        except Exception as e:
            return False, f"Backup failed: {str(e)}"
    
    def generate_pdf_report(self, student_id):
        """Generate PDF report card with colors"""
        student = self.search_student(student_id)
        
        if not student:
            return False, "Student not found!", None
        
        try:
            filename = f"ReportCard_{student_id}_{student['name'].replace(' ', '_')}.pdf"
            filepath = os.path.join(self.reports_dir, filename)
            
            c = canvas.Canvas(filepath, pagesize=A4)
            width, height = A4
            
            # Red header background
            c.setFillColor(colors.HexColor('#CC0000'))
            c.rect(0, height - 120, width, 120, fill=True, stroke=False)
            
            # College name
            c.setFillColor(colors.white)
            c.setFont("Helvetica-Bold", 24)
            c.drawCentredString(width/2, height - 50, "SETHUPATHY GOVERNMENT ARTS COLLEGE")
            
            c.setFont("Helvetica", 14)
            c.drawCentredString(width/2, height - 75, "RAMANATHAPURAM")
            
            c.setFont("Helvetica-Bold", 16)
            c.drawCentredString(width/2, height - 100, "STUDENT REPORT CARD")
            
            # Student details
            c.setFillColor(colors.black)
            y_position = height - 160
            
            c.setFont("Helvetica-Bold", 12)
            c.drawString(50, y_position, f"Student ID: {student['student_id']}")
            c.drawString(350, y_position, f"Date: {datetime.now().strftime('%d-%m-%Y')}")
            
            y_position -= 25
            c.drawString(50, y_position, f"Name: {student['name']}")
            
            # Line separator
            y_position -= 20
            c.setStrokeColor(colors.HexColor('#CC0000'))
            c.setLineWidth(2)
            c.line(50, y_position, width - 50, y_position)
            
            # Table header
            y_position -= 40
            c.setFillColor(colors.HexColor('#CC0000'))
            c.rect(50, y_position - 5, width - 100, 30, fill=True, stroke=False)
            
            c.setFillColor(colors.white)
            c.setFont("Helvetica-Bold", 11)
            c.drawString(70, y_position + 5, "Subject")
            c.drawString(250, y_position + 5, "Marks Obtained")
            c.drawString(380, y_position + 5, "Max Marks")
            c.drawString(490, y_position + 5, "Grade")
            
            # Subject marks
            subjects = [
                ("Mathematics", student['math']),
                ("Science", student['science']),
                ("English", student['english']),
                ("Social Studies", student['social_studies']),
                ("Computer Science", student['computer'])
            ]
            
            c.setFont("Helvetica", 11)
            y_position -= 30
            
            for i, (subject, marks) in enumerate(subjects):
                if i % 2 == 0:
                    c.setFillColor(colors.HexColor('#F0F0F0'))
                    c.rect(50, y_position - 5, width - 100, 25, fill=True, stroke=False)
                
                c.setFillColor(colors.black)
                c.drawString(70, y_position + 5, subject)
                c.drawString(280, y_position + 5, str(marks))
                c.drawString(410, y_position + 5, "100")
                
                subject_grade = self.calculate_grade(float(marks))
                c.drawString(510, y_position + 5, subject_grade)
                
                y_position -= 25
            
            # Summary section
            y_position -= 20
            c.setStrokeColor(colors.HexColor('#CC0000'))
            c.setLineWidth(2)
            c.line(50, y_position, width - 50, y_position)
            
            y_position -= 30
            c.setFont("Helvetica-Bold", 12)
            c.setFillColor(colors.HexColor('#CC0000'))
            c.drawString(70, y_position, f"Total Marks: {student['total']}/500")
            
            y_position -= 25
            c.drawString(70, y_position, f"Percentage: {student['percentage']}%")
            
            y_position -= 25
            c.drawString(70, y_position, f"Overall Grade: {student['grade']}")
            
            # Grade scale
            y_position -= 50
            c.setFillColor(colors.HexColor('#F0F0F0'))
            c.rect(50, y_position - 50, width - 100, 60, fill=True, stroke=True)
            
            c.setFillColor(colors.black)
            c.setFont("Helvetica-Bold", 11)
            c.drawString(70, y_position - 10, "GRADE SCALE:")
            
            c.setFont("Helvetica", 10)
            c.drawString(70, y_position - 25, "A+ (90-100)  |  A (80-89)  |  B (70-79)")
            c.drawString(70, y_position - 40, "C (60-69)  |  D (50-59)  |  F (Below 50)")
            
            # Footer
            c.setFont("Helvetica-Oblique", 9)
            c.setFillColor(colors.gray)
            c.drawCentredString(width/2, 50, "This is a computer-generated report card")
            
            c.save()
            
            return True, f"PDF Report card generated successfully!", filepath
            
        except Exception as e:
            return False, f"Error generating PDF: {str(e)}", None


class ReportCardGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Student Report Card Management System - Enhanced")
        self.root.geometry("1000x750")
        self.root.configure(bg="#f0f0f0")
        
        # Initialize system
        self.system = ReportCardSystem()
        
        # Email configuration
        self.sender_email = ""
        self.sender_password = ""
        
        # Setup window close protocol
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Create UI
        self.create_header()
        self.create_notebook()
    
    def on_closing(self):
        """Handle window close event"""
        if messagebox.askokcancel("Quit", "Do you want to exit the application?"):
            try:
                self.root.quit()
                self.root.destroy()
                sys.exit(0)
            except:
                sys.exit(0)
    
    def exit_application(self):
        """Exit application"""
        if messagebox.askokcancel("Exit", "Are you sure you want to exit?"):
            try:
                self.root.quit()
                self.root.destroy()
                sys.exit(0)
            except:
                sys.exit(0)
    
    def create_header(self):
        """Create college header"""
        header_frame = tk.Frame(self.root, bg="#CC0000", height=100)
        header_frame.pack(fill="x")
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text="SETHUPATHY GOVERNMENT ARTS COLLEGE", 
                font=("Arial", 20, "bold"), bg="#CC0000", fg="white").pack(pady=10)
        tk.Label(header_frame, text="RAMANATHAPURAM", 
                font=("Arial", 12), bg="#CC0000", fg="white").pack()
        tk.Label(header_frame, text="Student Report Card Management System", 
                font=("Arial", 11, "italic"), bg="#CC0000", fg="white").pack(pady=5)
        
        # Exit button
        exit_btn = tk.Button(header_frame, text="✕ Exit", font=("Arial", 10, "bold"),
                            bg="white", fg="#CC0000", padx=15, pady=5,
                            command=self.exit_application, cursor="hand2")
        exit_btn.place(relx=0.95, rely=0.5, anchor="e")
    
    def create_notebook(self):
        """Create tabbed interface"""
        style = ttk.Style()
        style.configure("TNotebook", background="#f0f0f0")
        style.configure("TFrame", background="#f0f0f0")
        
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Tabs
        self.add_tab = ttk.Frame(notebook)
        notebook.add(self.add_tab, text="  Add Student  ")
        self.create_add_student_tab()
        
        self.view_tab = ttk.Frame(notebook)
        notebook.add(self.view_tab, text="  View All Students  ")
        self.create_view_students_tab()
        
        self.search_tab = ttk.Frame(notebook)
        notebook.add(self.search_tab, text="  Search & Filter  ")
        self.create_search_tab()
        
        self.report_tab = ttk.Frame(notebook)
        notebook.add(self.report_tab, text="  Generate Report  ")
        self.create_generate_report_tab()
        
        self.email_tab = ttk.Frame(notebook)
        notebook.add(self.email_tab, text="  Email Reports  ")
        self.create_email_tab()
        
        self.analytics_tab = ttk.Frame(notebook)
        notebook.add(self.analytics_tab, text="  Analytics  ")
        self.create_analytics_tab()
        
        self.settings_tab = ttk.Frame(notebook)
        notebook.add(self.settings_tab, text="  Settings  ")
        self.create_settings_tab()
    
    def create_add_student_tab(self):
        """Create add student form"""
        form_frame = tk.Frame(self.add_tab, bg="#ffffff", relief="ridge", borderwidth=2)
        form_frame.pack(padx=20, pady=20, fill="both", expand=True)
        
        tk.Label(form_frame, text="Add New Student", font=("Arial", 16, "bold"), 
                bg="#ffffff", fg="#CC0000").pack(pady=15)
        
        fields_frame = tk.Frame(form_frame, bg="#ffffff")
        fields_frame.pack(padx=30, pady=10)
        
        # Student ID
        tk.Label(fields_frame, text="Student ID:", font=("Arial", 11), 
                bg="#ffffff").grid(row=0, column=0, sticky="w", pady=10)
        self.student_id_entry = tk.Entry(fields_frame, font=("Arial", 11), width=30)
        self.student_id_entry.grid(row=0, column=1, padx=20, pady=10)
        
        # Name
        tk.Label(fields_frame, text="Student Name:", font=("Arial", 11), 
                bg="#ffffff").grid(row=1, column=0, sticky="w", pady=10)
        self.name_entry = tk.Entry(fields_frame, font=("Arial", 11), width=30)
        self.name_entry.grid(row=1, column=1, padx=20, pady=10)
        
        # Email
        tk.Label(fields_frame, text="Email (Optional):", font=("Arial", 11), 
                bg="#ffffff").grid(row=2, column=0, sticky="w", pady=10)
        self.email_entry = tk.Entry(fields_frame, font=("Arial", 11), width=30)
        self.email_entry.grid(row=2, column=1, padx=20, pady=10)
        
        # Marks section
        tk.Label(fields_frame, text="Enter Marks (out of 100)", font=("Arial", 11, "bold"), 
                bg="#ffffff", fg="#CC0000").grid(row=3, column=0, columnspan=2, pady=15)
        
        # Mathematics
        tk.Label(fields_frame, text="Mathematics:", font=("Arial", 11), 
                bg="#ffffff").grid(row=4, column=0, sticky="w", pady=10)
        self.math_entry = tk.Entry(fields_frame, font=("Arial", 11), width=30)
        self.math_entry.grid(row=4, column=1, padx=20, pady=10)
        
        # Science
        tk.Label(fields_frame, text="Science:", font=("Arial", 11), 
                bg="#ffffff").grid(row=5, column=0, sticky="w", pady=10)
        self.science_entry = tk.Entry(fields_frame, font=("Arial", 11), width=30)
        self.science_entry.grid(row=5, column=1, padx=20, pady=10)
        
        # English
        tk.Label(fields_frame, text="English:", font=("Arial", 11), 
                bg="#ffffff").grid(row=6, column=0, sticky="w", pady=10)
        self.english_entry = tk.Entry(fields_frame, font=("Arial", 11), width=30)
        self.english_entry.grid(row=6, column=1, padx=20, pady=10)
        
        # Social Studies
        tk.Label(fields_frame, text="Social Studies:", font=("Arial", 11), 
                bg="#ffffff").grid(row=7, column=0, sticky="w", pady=10)
        self.social_entry = tk.Entry(fields_frame, font=("Arial", 11), width=30)
        self.social_entry.grid(row=7, column=1, padx=20, pady=10)
        
        # Computer
        tk.Label(fields_frame, text="Computer Science:", font=("Arial", 11), 
                bg="#ffffff").grid(row=8, column=0, sticky="w", pady=10)
        self.computer_entry = tk.Entry(fields_frame, font=("Arial", 11), width=30)
        self.computer_entry.grid(row=8, column=1, padx=20, pady=10)
        
        # Buttons
        button_frame = tk.Frame(form_frame, bg="#ffffff")
        button_frame.pack(pady=20)
        
        tk.Button(button_frame, text="Add Student", font=("Arial", 11, "bold"), 
                 bg="#CC0000", fg="white", width=15, command=self.add_student_action,
                 cursor="hand2").pack(side="left", padx=10)
        tk.Button(button_frame, text="Clear Form", font=("Arial", 11), 
                 bg="#666666", fg="white", width=15, command=self.clear_form,
                 cursor="hand2").pack(side="left", padx=10)
    
    def create_view_students_tab(self):
        """Create view students tab"""
        control_frame = tk.Frame(self.view_tab, bg="#f0f0f0")
        control_frame.pack(fill="x", padx=10, pady=10)
        
        tk.Button(control_frame, text="🔄 Refresh List", font=("Arial", 11, "bold"), 
                 bg="#CC0000", fg="white", command=self.refresh_students,
                 cursor="hand2").pack(side="left", padx=5)
        
        tree_frame = tk.Frame(self.view_tab, bg="#ffffff")
        tree_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        
        self.tree = ttk.Treeview(tree_frame, 
                                 columns=("ID", "Name", "Math", "Science", "English", "Social", "Computer", "Total", "Percentage", "Grade"),
                                 show="headings",
                                 yscrollcommand=vsb.set,
                                 xscrollcommand=hsb.set)
        
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)
        
        # Define columns
        self.tree.heading("ID", text="Student ID")
        self.tree.heading("Name", text="Name")
        self.tree.heading("Math", text="Math")
        self.tree.heading("Science", text="Science")
        self.tree.heading("English", text="English")
        self.tree.heading("Social", text="Social")
        self.tree.heading("Computer", text="Computer")
        self.tree.heading("Total", text="Total")
        self.tree.heading("Percentage", text="Percentage")
        self.tree.heading("Grade", text="Grade")
        
        for col in ("Math", "Science", "English", "Social", "Computer", "Grade"):
            self.tree.column(col, width=70)
        self.tree.column("ID", width=100)
        self.tree.column("Name", width=150)
        self.tree.column("Total", width=70)
        self.tree.column("Percentage", width=90)
        
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)
        
        self.refresh_students()
    
    def create_search_tab(self):
        """Create search and filter tab"""
        search_frame = tk.Frame(self.search_tab, bg="#ffffff", relief="ridge", borderwidth=2)
        search_frame.pack(padx=20, pady=20, fill="both", expand=True)
        
        tk.Label(search_frame, text="Search & Filter Students", font=("Arial", 16, "bold"), 
                bg="#ffffff", fg="#CC0000").pack(pady=15)
        
        input_frame = tk.Frame(search_frame, bg="#ffffff")
        input_frame.pack(pady=20)
        
        # Search by Name
        tk.Label(input_frame, text="Search by Name:", font=("Arial", 11), 
                bg="#ffffff").grid(row=0, column=0, sticky="w", padx=10, pady=10)
        self.search_name_entry = tk.Entry(input_frame, font=("Arial", 11), width=30)
        self.search_name_entry.grid(row=0, column=1, padx=10, pady=10)
        
        tk.Button(input_frame, text="🔍 Search", font=("Arial", 10, "bold"),
                 bg="#CC0000", fg="white", command=self.search_by_name_action,
                 cursor="hand2").grid(row=0, column=2, padx=10)
        
        # Filter by Grade
        tk.Label(input_frame, text="Filter by Grade:", font=("Arial", 11), 
                bg="#ffffff").grid(row=1, column=0, sticky="w", padx=10, pady=10)
        self.grade_filter = ttk.Combobox(input_frame, values=["All", "A+", "A", "B", "C", "D", "F"],
                                         font=("Arial", 11), width=28, state="readonly")
        self.grade_filter.set("All")
        self.grade_filter.grid(row=1, column=1, padx=10, pady=10)
        
        tk.Button(input_frame, text="Filter", font=("Arial", 10, "bold"),
                 bg="#CC0000", fg="white", command=self.filter_by_grade_action,
                 cursor="hand2").grid(row=1, column=2, padx=10)
        
        # Results treeview
        results_frame = tk.Frame(search_frame, bg="#ffffff")
        results_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        vsb = ttk.Scrollbar(results_frame, orient="vertical")
        hsb = ttk.Scrollbar(results_frame, orient="horizontal")
        
        self.search_tree = ttk.Treeview(results_frame,
                                        columns=("ID", "Name", "Total", "Percentage", "Grade"),
                                        show="headings",
                                        yscrollcommand=vsb.set,
                                        xscrollcommand=hsb.set)
        
        vsb.config(command=self.search_tree.yview)
        hsb.config(command=self.search_tree.xview)
        
        self.search_tree.heading("ID", text="Student ID")
        self.search_tree.heading("Name", text="Name")
        self.search_tree.heading("Total", text="Total Marks")
        self.search_tree.heading("Percentage", text="Percentage")
        self.search_tree.heading("Grade", text="Grade")
        
        self.search_tree.column("ID", width=100)
        self.search_tree.column("Name", width=200)
        self.search_tree.column("Total", width=100)
        self.search_tree.column("Percentage", width=100)
        self.search_tree.column("Grade", width=80)
        
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.search_tree.pack(fill="both", expand=True)
    
    def create_generate_report_tab(self):
        """Create generate report tab"""
        report_frame = tk.Frame(self.report_tab, bg="#ffffff", relief="ridge", borderwidth=2)
        report_frame.pack(padx=20, pady=20, fill="both", expand=True)
        
        tk.Label(report_frame, text="Generate PDF Report Card", font=("Arial", 16, "bold"), 
                bg="#ffffff", fg="#CC0000").pack(pady=20)
        
        input_frame = tk.Frame(report_frame, bg="#ffffff")
        input_frame.pack(pady=20)
        
        tk.Label(input_frame, text="Enter Student ID:", font=("Arial", 12), 
                bg="#ffffff").pack(side="left", padx=10)
        self.report_id_entry = tk.Entry(input_frame, font=("Arial", 12), width=20)
        self.report_id_entry.pack(side="left", padx=10)
        
        tk.Button(input_frame, text="📄 Generate PDF", font=("Arial", 11, "bold"), 
                 bg="#CC0000", fg="white", width=15, 
                 command=self.generate_report_action, cursor="hand2").pack(side="left", padx=10)
        
        self.info_text = scrolledtext.ScrolledText(report_frame, width=80, height=22, 
                                                    font=("Courier", 10), bg="#f9f9f9")
        self.info_text.pack(padx=20, pady=20)
    
    def create_email_tab(self):
        """Create email reports tab"""
        email_frame = tk.Frame(self.email_tab, bg="#ffffff", relief="ridge", borderwidth=2)
        email_frame.pack(padx=20, pady=20, fill="both", expand=True)
        
        tk.Label(email_frame, text="Email Report Cards", font=("Arial", 16, "bold"), 
                bg="#ffffff", fg="#CC0000").pack(pady=15)
        
        form_frame = tk.Frame(email_frame, bg="#ffffff")
        form_frame.pack(pady=20)
        
        # Student ID
        tk.Label(form_frame, text="Student ID:", font=("Arial", 11), 
                bg="#ffffff").grid(row=0, column=0, sticky="w", padx=10, pady=10)
        self.email_student_id = tk.Entry(form_frame, font=("Arial", 11), width=30)
        self.email_student_id.grid(row=0, column=1, padx=10, pady=10)
        
        # Recipient Email
        tk.Label(form_frame, text="Recipient Email:", font=("Arial", 11), 
                bg="#ffffff").grid(row=1, column=0, sticky="w", padx=10, pady=10)
        self.recipient_email = tk.Entry(form_frame, font=("Arial", 11), width=30)
        self.recipient_email.grid(row=1, column=1, padx=10, pady=10)
        
        tk.Button(form_frame, text="📧 Send Email", font=("Arial", 11, "bold"),
                 bg="#CC0000", fg="white", width=15, command=self.send_email_action,
                 cursor="hand2").grid(row=2, column=1, pady=20)
        
        # Email log
        tk.Label(email_frame, text="Email Log:", font=("Arial", 11, "bold"),
                bg="#ffffff").pack(pady=10)
        
        self.email_log = scrolledtext.ScrolledText(email_frame, width=80, height=15,
                                                   font=("Courier", 9), bg="#f9f9f9")
        self.email_log.pack(padx=20, pady=10)
    
    def create_analytics_tab(self):
        """Create analytics dashboard"""
        analytics_frame = tk.Frame(self.analytics_tab, bg="#ffffff", relief="ridge", borderwidth=2)
        analytics_frame.pack(padx=20, pady=20, fill="both", expand=True)
        
        tk.Label(analytics_frame, text="Class Analytics Dashboard", font=("Arial", 16, "bold"), 
                bg="#ffffff", fg="#CC0000").pack(pady=15)
        
        tk.Button(analytics_frame, text="🔄 Refresh Analytics", font=("Arial", 10, "bold"),
                 bg="#CC0000", fg="white", command=self.refresh_analytics,
                 cursor="hand2").pack(pady=10)
        
        self.analytics_display = tk.Frame(analytics_frame, bg="#ffffff")
        self.analytics_display.pack(fill="both", expand=True, padx=20, pady=10)
        
        self.refresh_analytics()
    
    def create_settings_tab(self):
        """Create settings tab"""
        settings_frame = tk.Frame(self.settings_tab, bg="#ffffff", relief="ridge", borderwidth=2)
        settings_frame.pack(padx=20, pady=20, fill="both", expand=True)
        
        tk.Label(settings_frame, text="System Settings", font=("Arial", 16, "bold"), 
                bg="#ffffff", fg="#CC0000").pack(pady=15)
        
        # Email configuration
        email_config_frame = tk.LabelFrame(settings_frame, text="Email Configuration", 
                                          font=("Arial", 12, "bold"), bg="#ffffff", padx=20, pady=20)
        email_config_frame.pack(padx=20, pady=10, fill="x")
        
        tk.Label(email_config_frame, text="Sender Email:", font=("Arial", 11),
                bg="#ffffff").grid(row=0, column=0, sticky="w", pady=10)
        self.sender_email_entry = tk.Entry(email_config_frame, font=("Arial", 11), width=35)
        self.sender_email_entry.grid(row=0, column=1, padx=10, pady=10)
        
        tk.Label(email_config_frame, text="App Password:", font=("Arial", 11),
                bg="#ffffff").grid(row=1, column=0, sticky="w", pady=10)
        self.sender_password_entry = tk.Entry(email_config_frame, font=("Arial", 11), width=35, show="*")
        self.sender_password_entry.grid(row=1, column=1, padx=10, pady=10)
        
        tk.Button(email_config_frame, text="Save Configuration", font=("Arial", 10, "bold"),
                 bg="#CC0000", fg="white", command=self.save_email_config,
                 cursor="hand2").grid(row=2, column=1, pady=10)
        
        # Backup section
        backup_frame = tk.LabelFrame(settings_frame, text="Database Backup",
                                     font=("Arial", 12, "bold"), bg="#ffffff", padx=20, pady=20)
        backup_frame.pack(padx=20, pady=10, fill="x")
        
        tk.Button(backup_frame, text="💾 Create Backup Now", font=("Arial", 11, "bold"),
                 bg="#28a745", fg="white", width=20, command=self.create_backup_action,
                 cursor="hand2").pack(pady=10)
        
        self.backup_status = tk.Label(backup_frame, text="No backup created yet", 
                                     font=("Arial", 10), bg="#ffffff", fg="gray")
        self.backup_status.pack(pady=5)
    
    # Action methods
    def add_student_action(self):
        """Handle add student"""
        try:
            student_id = self.student_id_entry.get().strip()
            name = self.name_entry.get().strip()
            email = self.email_entry.get().strip()
            math = float(self.math_entry.get())
            science = float(self.science_entry.get())
            english = float(self.english_entry.get())
            social = float(self.social_entry.get())
            computer = float(self.computer_entry.get())
            
            if not student_id or not name:
                messagebox.showerror("Error", "Please fill in Student ID and Name!")
                return
            
            success, message = self.system.add_student(student_id, name, math, science, 
                                                      english, social, computer, email)
            
            if success:
                messagebox.showinfo("Success", message)
                self.clear_form()
                self.refresh_students()
                self.refresh_analytics()
            else:
                messagebox.showerror("Error", message)
                
        except ValueError:
            messagebox.showerror("Error", "Please enter valid numeric marks!")
    
    def clear_form(self):
        """Clear all form fields"""
        self.student_id_entry.delete(0, tk.END)
        self.name_entry.delete(0, tk.END)
        self.email_entry.delete(0, tk.END)
        self.math_entry.delete(0, tk.END)
        self.science_entry.delete(0, tk.END)
        self.english_entry.delete(0, tk.END)
        self.social_entry.delete(0, tk.END)
        self.computer_entry.delete(0, tk.END)
    
    def refresh_students(self):
        """Refresh student list"""
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        students = self.system.get_all_students()
        for student in students:
            self.tree.insert("", "end", values=(
                student['student_id'],
                student['name'],
                student['math'],
                student['science'],
                student['english'],
                student['social_studies'],
                student['computer'],
                student['total'],
                f"{student['percentage']}%",
                student['grade']
            ))
    
    def search_by_name_action(self):
        """Search students by name"""
        name = self.search_name_entry.get().strip()
        
        if not name:
            messagebox.showwarning("Warning", "Please enter a name to search!")
            return
        
        for item in self.search_tree.get_children():
            self.search_tree.delete(item)
        
        students = self.system.search_by_name(name)
        
        if not students:
            messagebox.showinfo("No Results", f"No students found with name containing '{name}'")
            return
        
        for student in students:
            self.search_tree.insert("", "end", values=(
                student['student_id'],
                student['name'],
                student['total'],
                f"{student['percentage']}%",
                student['grade']
            ))
    
    def filter_by_grade_action(self):
        """Filter students by grade"""
        grade = self.grade_filter.get()
        
        for item in self.search_tree.get_children():
            self.search_tree.delete(item)
        
        students = self.system.filter_by_grade(grade)
        
        for student in students:
            self.search_tree.insert("", "end", values=(
                student['student_id'],
                student['name'],
                student['total'],
                f"{student['percentage']}%",
                student['grade']
            ))
    
    def generate_report_action(self):
        """Generate PDF report"""
        student_id = self.report_id_entry.get().strip()
        
        if not student_id:
            messagebox.showerror("Error", "Please enter Student ID!")
            return
        
        student = self.system.search_student(student_id)
        
        if not student:
            messagebox.showerror("Error", f"No student found with ID: {student_id}")
            return
        
        # Display info
        self.info_text.delete(1.0, tk.END)
        self.info_text.insert(tk.END, f"╔═══════════════════════════════════════════════════════════╗\n")
        self.info_text.insert(tk.END, f"║         STUDENT INFORMATION                               ║\n")
        self.info_text.insert(tk.END, f"╚═══════════════════════════════════════════════════════════╝\n\n")
        self.info_text.insert(tk.END, f"Student ID       : {student['student_id']}\n")
        self.info_text.insert(tk.END, f"Name             : {student['name']}\n")
        if student.get('email'):
            self.info_text.insert(tk.END, f"Email            : {student['email']}\n")
        self.info_text.insert(tk.END, f"\n{'─'*60}\n")
        self.info_text.insert(tk.END, f"MARKS:\n")
        self.info_text.insert(tk.END, f"{'─'*60}\n")
        self.info_text.insert(tk.END, f"  Mathematics       : {student['math']}/100\n")
        self.info_text.insert(tk.END, f"  Science           : {student['science']}/100\n")
        self.info_text.insert(tk.END, f"  English           : {student['english']}/100\n")
        self.info_text.insert(tk.END, f"  Social Studies    : {student['social_studies']}/100\n")
        self.info_text.insert(tk.END, f"  Computer Science  : {student['computer']}/100\n")
        self.info_text.insert(tk.END, f"\n{'─'*60}\n")
        self.info_text.insert(tk.END, f"SUMMARY:\n")
        self.info_text.insert(tk.END, f"{'─'*60}\n")
        self.info_text.insert(tk.END, f"  Total Marks       : {student['total']}/500\n")
        self.info_text.insert(tk.END, f"  Percentage        : {student['percentage']}%\n")
        self.info_text.insert(tk.END, f"  Overall Grade     : {student['grade']}\n")
        self.info_text.insert(tk.END, f"{'─'*60}\n")
        
        # Generate PDF
        success, message, filepath = self.system.generate_pdf_report(student_id)
        
        if success:
            self.info_text.insert(tk.END, f"\n✓ PDF Generated Successfully!\n")
            self.info_text.insert(tk.END, f"  Location: {filepath}\n")
            messagebox.showinfo("Success", message)
        else:
            messagebox.showerror("Error", message)
    
    def send_email_action(self):
        """Send report card via email"""
        student_id = self.email_student_id.get().strip()
        recipient = self.recipient_email.get().strip()
        
        if not student_id or not recipient:
            messagebox.showerror("Error", "Please fill in both Student ID and Email!")
            return
        
        if not self.sender_email or not self.sender_password:
            messagebox.showerror("Error", "Please configure email settings first in Settings tab!")
            return
        
        student = self.system.search_student(student_id)
        if not student:
            messagebox.showerror("Error", "Student not found!")
            return
        
        # Generate PDF first
        success, message, filepath = self.system.generate_pdf_report(student_id)
        
        if not success:
            messagebox.showerror("Error", "Failed to generate PDF!")
            return
        
        try:
            # Create message
            msg = MIMEMultipart()
            msg['From'] = self.sender_email
            msg['To'] = recipient
            msg['Subject'] = f"Report Card - {student['name']} ({student['student_id']})"
            
            body = f"""Dear Parent/Guardian,

Please find attached the report card for {student['name']}.

Academic Summary:
- Total Marks: {student['total']}/500
- Percentage: {student['percentage']}%
- Grade: {student['grade']}

Best Regards,
Sethupathy Government Arts College
Ramanathapuram
"""
            
            msg.attach(MIMEText(body, 'plain'))
            
            # Attach PDF
            with open(filepath, 'rb') as file:
                attach = MIMEApplication(file.read(), _subtype='pdf')
                attach.add_header('Content-Disposition', 'attachment', 
                                filename=f"ReportCard_{student_id}.pdf")
                msg.attach(attach)
            
            # Send email
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(self.sender_email, self.sender_password)
            server.send_message(msg)
            server.quit()
            
            # Log success
            log_entry = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Successfully sent report for {student['name']} to {recipient}\n"
            self.email_log.insert(tk.END, log_entry)
            
            messagebox.showinfo("Success", f"Email sent successfully to {recipient}!")
            
        except Exception as e:
            error_log = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ERROR: {str(e)}\n"
            self.email_log.insert(tk.END, error_log)
            messagebox.showerror("Error", f"Failed to send email: {str(e)}")
    
    def refresh_analytics(self):
        """Refresh analytics dashboard"""
        # Clear previous widgets
        for widget in self.analytics_display.winfo_children():
            widget.destroy()
        
        stats = self.system.get_statistics()
        
        if not stats or stats['total_students'] == 0:
            tk.Label(self.analytics_display, text="No data available", 
                    font=("Arial", 12), bg="#ffffff").pack(pady=20)
            return
        
        # Statistics cards
        cards_frame = tk.Frame(self.analytics_display, bg="#ffffff")
        cards_frame.pack(fill="x", pady=20)
        
        # Total Students Card
        card1 = tk.Frame(cards_frame, bg="#007bff", relief="raised", borderwidth=2)
        card1.pack(side="left", padx=20, pady=10, ipadx=30, ipady=20)
        tk.Label(card1, text="Total Students", font=("Arial", 12, "bold"),
                bg="#007bff", fg="white").pack()
        tk.Label(card1, text=str(stats['total_students']), font=("Arial", 24, "bold"),
                bg="#007bff", fg="white").pack()
        
        # Average Card
        card2 = tk.Frame(cards_frame, bg="#28a745", relief="raised", borderwidth=2)
        card2.pack(side="left", padx=20, pady=10, ipadx=30, ipady=20)
        tk.Label(card2, text="Class Average", font=("Arial", 12, "bold"),
                bg="#28a745", fg="white").pack()
        tk.Label(card2, text=f"{stats['average_percentage']}%", font=("Arial", 24, "bold"),
                bg="#28a745", fg="white").pack()
        
        # Top Student Card
        if stats['top_student']:
            card3 = tk.Frame(cards_frame, bg="#ffc107", relief="raised", borderwidth=2)
            card3.pack(side="left", padx=20, pady=10, ipadx=30, ipady=20)
            tk.Label(card3, text="Top Student", font=("Arial", 12, "bold"),
                    bg="#ffc107", fg="black").pack()
            tk.Label(card3, text=stats['top_student'][0], font=("Arial", 12, "bold"),
                    bg="#ffc107", fg="black").pack()
            tk.Label(card3, text=f"{stats['top_student'][1]}%", font=("Arial", 18, "bold"),
                    bg="#ffc107", fg="black").pack()
        
        # Grade Distribution
        dist_frame = tk.LabelFrame(self.analytics_display, text="Grade Distribution",
                                  font=("Arial", 13, "bold"), bg="#ffffff", padx=20, pady=15)
        dist_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        grade_colors = {
            'A+': '#28a745',
            'A': '#17a2b8',
            'B': '#007bff',
            'C': '#ffc107',
            'D': '#fd7e14',
            'F': '#dc3545'
        }
        
        for grade in ['A+', 'A', 'B', 'C', 'D', 'F']:
            count = stats['grade_distribution'].get(grade, 0)
            percentage = (count / stats['total_students'] * 100) if stats['total_students'] > 0 else 0
            
            grade_row = tk.Frame(dist_frame, bg="#ffffff")
            grade_row.pack(fill="x", pady=5)
            
            tk.Label(grade_row, text=f"Grade {grade}:", font=("Arial", 11, "bold"),
                    bg="#ffffff", width=10, anchor="w").pack(side="left")
            
            # Progress bar
            bar_frame = tk.Frame(grade_row, bg="#e9ecef", height=25, width=300)
            bar_frame.pack(side="left", padx=10)
            bar_frame.pack_propagate(False)
            
            if count > 0:
                bar_width = int(300 * percentage / 100)
                tk.Frame(bar_frame, bg=grade_colors.get(grade, '#007bff'),
                        width=bar_width).pack(side="left", fill="y")
            
            tk.Label(grade_row, text=f"{count} students ({percentage:.1f}%)",
                    font=("Arial", 10), bg="#ffffff").pack(side="left", padx=10)
    
    def save_email_config(self):
        """Save email configuration"""
        self.sender_email = self.sender_email_entry.get().strip()
        self.sender_password = self.sender_password_entry.get().strip()
        
        if self.sender_email and self.sender_password:
            messagebox.showinfo("Success", "Email configuration saved!")
        else:
            messagebox.showwarning("Warning", "Please fill in both email and password!")
    
    def create_backup_action(self):
        """Create database backup"""
        success, message = self.system.create_backup()
        
        if success:
            self.backup_status.config(text=f"Last backup: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                                     fg="green")
            messagebox.showinfo("Success", message)
        else:
            messagebox.showerror("Error", message)


# Main execution
if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = ReportCardGUI(root)
        
        print("="*60)
        print("Student Report Card System - Enhanced Edition")
        print("="*60)
        print("System started successfully!")
        print("Database: SQLite")
        print("Features: PDF Reports, Email, Analytics, Search & Backup")
        print("="*60)
        
        root.mainloop()
        
    except KeyboardInterrupt:
        print("\n\nApplication closed by user (Ctrl+C)")
        try:
            root.quit()
            root.destroy()
        except:
            pass
        sys.exit(0)
        
    except Exception as e:
        print(f"An error occurred: {e}")
        sys.exit(1)
    
    finally:
        print("\nThank you for using Student Report Card System!")
        sys.exit(0)

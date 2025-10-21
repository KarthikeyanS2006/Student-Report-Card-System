import os
import csv
import sys
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors

class ReportCardSystem:
    def __init__(self):
        # Create directory to store student data
        self.data_dir = "student_records"
        self.csv_file = os.path.join(self.data_dir, "students.csv")
        self.reports_dir = os.path.join(self.data_dir, "report_cards_pdf")
        
        # Create directories if they don't exist
        os.makedirs(self.data_dir, exist_ok=True)
        os.makedirs(self.reports_dir, exist_ok=True)
        
        # Initialize CSV file if it doesn't exist
        if not os.path.exists(self.csv_file):
            self.create_csv_file()
    
    def create_csv_file(self):
        """Create CSV file with headers"""
        headers = ['Student_ID', 'Name', 'Math', 'Science', 'English', 
                   'Social_Studies', 'Computer', 'Total', 'Percentage', 'Grade']
        with open(self.csv_file, 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(headers)
    
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
    
    def add_student(self, student_id, name, math, science, english, social, computer):
        """Add new student with marks"""
        try:
            # Validate marks
            marks = [math, science, english, social, computer]
            if any(mark < 0 or mark > 100 for mark in marks):
                return False, "Marks must be between 0 and 100!"
            
            # Calculate total and percentage
            total = sum(marks)
            percentage = total / 5
            grade = self.calculate_grade(percentage)
            
            # Save to CSV
            with open(self.csv_file, 'a', newline='') as file:
                writer = csv.writer(file)
                writer.writerow([student_id, name, math, science, english, 
                               social, computer, total, round(percentage, 2), grade])
            
            return True, f"Student {name} added successfully!"
            
        except Exception as e:
            return False, str(e)
    
    def get_all_students(self):
        """Return all students as list"""
        try:
            with open(self.csv_file, 'r') as file:
                reader = csv.DictReader(file)
                return list(reader)
        except:
            return []
    
    def search_student(self, student_id):
        """Search for specific student"""
        try:
            with open(self.csv_file, 'r') as file:
                reader = csv.DictReader(file)
                for row in reader:
                    if row['Student_ID'] == student_id:
                        return row
            return None
        except:
            return None
    
    def generate_pdf_report(self, student_id):
        """Generate PDF report card with colors"""
        student = self.search_student(student_id)
        
        if not student:
            return False, "Student not found!"
        
        try:
            # Create PDF filename
            filename = f"ReportCard_{student_id}_{student['Name'].replace(' ', '_')}.pdf"
            filepath = os.path.join(self.reports_dir, filename)
            
            # Create PDF
            c = canvas.Canvas(filepath, pagesize=A4)
            width, height = A4
            
            # Draw red header background
            c.setFillColor(colors.HexColor('#CC0000'))
            c.rect(0, height - 120, width, 120, fill=True, stroke=False)
            
            # College name in white
            c.setFillColor(colors.white)
            c.setFont("Helvetica-Bold", 24)
            c.drawCentredString(width/2, height - 50, "SETHUPATHY GOVERNMENT ARTS COLLEGE")
            
            c.setFont("Helvetica", 14)
            c.drawCentredString(width/2, height - 75, "RAMANATHAPURAM")
            
            c.setFont("Helvetica-Bold", 16)
            c.drawCentredString(width/2, height - 100, "STUDENT REPORT CARD")
            
            # Student details section
            c.setFillColor(colors.black)
            y_position = height - 160
            
            c.setFont("Helvetica-Bold", 12)
            c.drawString(50, y_position, f"Student ID: {student['Student_ID']}")
            c.drawString(350, y_position, f"Date: {datetime.now().strftime('%d-%m-%Y')}")
            
            y_position -= 25
            c.drawString(50, y_position, f"Name: {student['Name']}")
            
            # Draw line separator
            y_position -= 20
            c.setStrokeColor(colors.HexColor('#CC0000'))
            c.setLineWidth(2)
            c.line(50, y_position, width - 50, y_position)
            
            # Table header with red background
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
                ("Mathematics", student['Math']),
                ("Science", student['Science']),
                ("English", student['English']),
                ("Social Studies", student['Social_Studies']),
                ("Computer Science", student['Computer'])
            ]
            
            c.setFont("Helvetica", 11)
            y_position -= 30
            
            for i, (subject, marks) in enumerate(subjects):
                # Alternate row colors
                if i % 2 == 0:
                    c.setFillColor(colors.HexColor('#F0F0F0'))
                    c.rect(50, y_position - 5, width - 100, 25, fill=True, stroke=False)
                
                c.setFillColor(colors.black)
                c.drawString(70, y_position + 5, subject)
                c.drawString(280, y_position + 5, marks)
                c.drawString(410, y_position + 5, "100")
                
                subject_grade = self.calculate_grade(float(marks))
                c.drawString(510, y_position + 5, subject_grade)
                
                y_position -= 25
            
            # Total and summary section
            y_position -= 20
            c.setStrokeColor(colors.HexColor('#CC0000'))
            c.setLineWidth(2)
            c.line(50, y_position, width - 50, y_position)
            
            y_position -= 30
            c.setFont("Helvetica-Bold", 12)
            c.setFillColor(colors.HexColor('#CC0000'))
            c.drawString(70, y_position, f"Total Marks: {student['Total']}/500")
            
            y_position -= 25
            c.drawString(70, y_position, f"Percentage: {student['Percentage']}%")
            
            y_position -= 25
            c.drawString(70, y_position, f"Overall Grade: {student['Grade']}")
            
            # Grade scale box
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
            
            # Save PDF
            c.save()
            
            return True, f"PDF Report card generated successfully!\nSaved at: {filepath}"
            
        except Exception as e:
            return False, f"Error generating PDF: {str(e)}"


class ReportCardGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Student Report Card Management System")
        self.root.geometry("900x700")
        self.root.configure(bg="#f0f0f0")
        
        # Initialize system
        self.system = ReportCardSystem()
        
        # Setup window close protocol
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Create main UI
        self.create_header()
        self.create_notebook()
        
    def on_closing(self):
        """Handle window close event with confirmation"""
        if messagebox.askokcancel("Quit", "Do you want to exit the application?"):
            try:
                self.root.quit()
                self.root.destroy()
                sys.exit(0)
            except:
                sys.exit(0)
    
    def exit_application(self):
        """Exit application with confirmation"""
        if messagebox.askokcancel("Exit", "Are you sure you want to exit?"):
            try:
                self.root.quit()
                self.root.destroy()
                sys.exit(0)
            except:
                sys.exit(0)
    
    def create_header(self):
        """Create college header with exit button"""
        header_frame = tk.Frame(self.root, bg="#CC0000", height=100)
        header_frame.pack(fill="x")
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text="SETHUPATHY GOVERNMENT ARTS COLLEGE", 
                font=("Arial", 20, "bold"), bg="#CC0000", fg="white").pack(pady=10)
        tk.Label(header_frame, text="RAMANATHAPURAM", 
                font=("Arial", 12), bg="#CC0000", fg="white").pack()
        tk.Label(header_frame, text="Student Report Card Management System", 
                font=("Arial", 11, "italic"), bg="#CC0000", fg="white").pack(pady=5)
        
        # Add Exit button in header
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
        
        # Add Student Tab
        self.add_tab = ttk.Frame(notebook)
        notebook.add(self.add_tab, text="  Add Student  ")
        self.create_add_student_tab()
        
        # View Students Tab
        self.view_tab = ttk.Frame(notebook)
        notebook.add(self.view_tab, text="  View All Students  ")
        self.create_view_students_tab()
        
        # Generate Report Tab
        self.report_tab = ttk.Frame(notebook)
        notebook.add(self.report_tab, text="  Generate Report  ")
        self.create_generate_report_tab()
    
    def create_add_student_tab(self):
        """Create add student form"""
        # Main frame
        form_frame = tk.Frame(self.add_tab, bg="#ffffff", relief="ridge", borderwidth=2)
        form_frame.pack(padx=20, pady=20, fill="both", expand=True)
        
        tk.Label(form_frame, text="Add New Student", font=("Arial", 16, "bold"), 
                bg="#ffffff", fg="#CC0000").pack(pady=15)
        
        # Form fields
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
        
        # Marks section
        tk.Label(fields_frame, text="Enter Marks (out of 100)", font=("Arial", 11, "bold"), 
                bg="#ffffff", fg="#CC0000").grid(row=2, column=0, columnspan=2, pady=15)
        
        # Mathematics
        tk.Label(fields_frame, text="Mathematics:", font=("Arial", 11), 
                bg="#ffffff").grid(row=3, column=0, sticky="w", pady=10)
        self.math_entry = tk.Entry(fields_frame, font=("Arial", 11), width=30)
        self.math_entry.grid(row=3, column=1, padx=20, pady=10)
        
        # Science
        tk.Label(fields_frame, text="Science:", font=("Arial", 11), 
                bg="#ffffff").grid(row=4, column=0, sticky="w", pady=10)
        self.science_entry = tk.Entry(fields_frame, font=("Arial", 11), width=30)
        self.science_entry.grid(row=4, column=1, padx=20, pady=10)
        
        # English
        tk.Label(fields_frame, text="English:", font=("Arial", 11), 
                bg="#ffffff").grid(row=5, column=0, sticky="w", pady=10)
        self.english_entry = tk.Entry(fields_frame, font=("Arial", 11), width=30)
        self.english_entry.grid(row=5, column=1, padx=20, pady=10)
        
        # Social Studies
        tk.Label(fields_frame, text="Social Studies:", font=("Arial", 11), 
                bg="#ffffff").grid(row=6, column=0, sticky="w", pady=10)
        self.social_entry = tk.Entry(fields_frame, font=("Arial", 11), width=30)
        self.social_entry.grid(row=6, column=1, padx=20, pady=10)
        
        # Computer
        tk.Label(fields_frame, text="Computer Science:", font=("Arial", 11), 
                bg="#ffffff").grid(row=7, column=0, sticky="w", pady=10)
        self.computer_entry = tk.Entry(fields_frame, font=("Arial", 11), width=30)
        self.computer_entry.grid(row=7, column=1, padx=20, pady=10)
        
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
        # Control frame
        control_frame = tk.Frame(self.view_tab, bg="#f0f0f0")
        control_frame.pack(fill="x", padx=10, pady=10)
        
        tk.Button(control_frame, text="🔄 Refresh List", font=("Arial", 11, "bold"), 
                 bg="#CC0000", fg="white", command=self.refresh_students,
                 cursor="hand2").pack(side="left", padx=5)
        
        # Treeview
        tree_frame = tk.Frame(self.view_tab, bg="#ffffff")
        tree_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        
        # Create treeview
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
        
        # Column widths
        for col in ("Math", "Science", "English", "Social", "Computer", "Grade"):
            self.tree.column(col, width=70)
        self.tree.column("ID", width=100)
        self.tree.column("Name", width=150)
        self.tree.column("Total", width=70)
        self.tree.column("Percentage", width=90)
        
        # Pack elements
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)
        
        # Load initial data
        self.refresh_students()
    
    def create_generate_report_tab(self):
        """Create generate report tab"""
        # Main frame
        report_frame = tk.Frame(self.report_tab, bg="#ffffff", relief="ridge", borderwidth=2)
        report_frame.pack(padx=20, pady=20, fill="both", expand=True)
        
        tk.Label(report_frame, text="Generate PDF Report Card", font=("Arial", 16, "bold"), 
                bg="#ffffff", fg="#CC0000").pack(pady=20)
        
        # Input frame
        input_frame = tk.Frame(report_frame, bg="#ffffff")
        input_frame.pack(pady=20)
        
        tk.Label(input_frame, text="Enter Student ID:", font=("Arial", 12), 
                bg="#ffffff").pack(side="left", padx=10)
        self.report_id_entry = tk.Entry(input_frame, font=("Arial", 12), width=20)
        self.report_id_entry.pack(side="left", padx=10)
        
        tk.Button(input_frame, text="📄 Generate PDF", font=("Arial", 11, "bold"), 
                 bg="#CC0000", fg="white", width=15, 
                 command=self.generate_report_action, cursor="hand2").pack(side="left", padx=10)
        
        # Info display
        self.info_text = scrolledtext.ScrolledText(report_frame, width=70, height=20, 
                                                    font=("Courier", 10), bg="#f9f9f9")
        self.info_text.pack(padx=20, pady=20)
    
    def add_student_action(self):
        """Handle add student button click"""
        try:
            student_id = self.student_id_entry.get().strip()
            name = self.name_entry.get().strip()
            math = float(self.math_entry.get())
            science = float(self.science_entry.get())
            english = float(self.english_entry.get())
            social = float(self.social_entry.get())
            computer = float(self.computer_entry.get())
            
            if not student_id or not name:
                messagebox.showerror("Error", "Please fill in Student ID and Name!")
                return
            
            success, message = self.system.add_student(student_id, name, math, science, english, social, computer)
            
            if success:
                messagebox.showinfo("Success", message)
                self.clear_form()
                self.refresh_students()
            else:
                messagebox.showerror("Error", message)
                
        except ValueError:
            messagebox.showerror("Error", "Please enter valid numeric marks!")
    
    def clear_form(self):
        """Clear all form fields"""
        self.student_id_entry.delete(0, tk.END)
        self.name_entry.delete(0, tk.END)
        self.math_entry.delete(0, tk.END)
        self.science_entry.delete(0, tk.END)
        self.english_entry.delete(0, tk.END)
        self.social_entry.delete(0, tk.END)
        self.computer_entry.delete(0, tk.END)
    
    def refresh_students(self):
        """Refresh student list in treeview"""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Load students
        students = self.system.get_all_students()
        for student in students:
            self.tree.insert("", "end", values=(
                student['Student_ID'],
                student['Name'],
                student['Math'],
                student['Science'],
                student['English'],
                student['Social_Studies'],
                student['Computer'],
                student['Total'],
                f"{student['Percentage']}%",
                student['Grade']
            ))
    
    def generate_report_action(self):
        """Handle generate report button click"""
        student_id = self.report_id_entry.get().strip()
        
        if not student_id:
            messagebox.showerror("Error", "Please enter Student ID!")
            return
        
        # Get student info
        student = self.system.search_student(student_id)
        
        if not student:
            messagebox.showerror("Error", f"No student found with ID: {student_id}")
            return
        
        # Display info
        self.info_text.delete(1.0, tk.END)
        self.info_text.insert(tk.END, f"╔═══════════════════════════════════════════════════╗\n")
        self.info_text.insert(tk.END, f"║         STUDENT INFORMATION                       ║\n")
        self.info_text.insert(tk.END, f"╚═══════════════════════════════════════════════════╝\n\n")
        self.info_text.insert(tk.END, f"Student ID: {student['Student_ID']}\n")
        self.info_text.insert(tk.END, f"Name: {student['Name']}\n")
        self.info_text.insert(tk.END, f"\n{'─'*50}\n")
        self.info_text.insert(tk.END, f"MARKS:\n")
        self.info_text.insert(tk.END, f"{'─'*50}\n")
        self.info_text.insert(tk.END, f"  Mathematics       : {student['Math']}/100\n")
        self.info_text.insert(tk.END, f"  Science           : {student['Science']}/100\n")
        self.info_text.insert(tk.END, f"  English           : {student['English']}/100\n")
        self.info_text.insert(tk.END, f"  Social Studies    : {student['Social_Studies']}/100\n")
        self.info_text.insert(tk.END, f"  Computer Science  : {student['Computer']}/100\n")
        self.info_text.insert(tk.END, f"\n{'─'*50}\n")
        self.info_text.insert(tk.END, f"SUMMARY:\n")
        self.info_text.insert(tk.END, f"{'─'*50}\n")
        self.info_text.insert(tk.END, f"  Total Marks       : {student['Total']}/500\n")
        self.info_text.insert(tk.END, f"  Percentage        : {student['Percentage']}%\n")
        self.info_text.insert(tk.END, f"  Overall Grade     : {student['Grade']}\n")
        self.info_text.insert(tk.END, f"{'─'*50}\n")
        
        # Generate PDF
        success, message = self.system.generate_pdf_report(student_id)
        
        if success:
            messagebox.showinfo("Success", message)
            self.info_text.insert(tk.END, f"\n✓ PDF Generated Successfully!\n")
        else:
            messagebox.showerror("Error", message)


# Main execution with all exit handlers
if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = ReportCardGUI(root)
        
        # Start the application
        print("Student Report Card System started...")
        print("Press the Exit button or close the window to quit.")
        
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
        print("Thank you for using Student Report Card System!")
        sys.exit(0)

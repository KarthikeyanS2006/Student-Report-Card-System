import os
import csv
from datetime import datetime
from prettytable import PrettyTable

class ReportCardSystem:
    def __init__(self):
        # Create directory to store student data
        self.data_dir = "student_records"
        self.csv_file = os.path.join(self.data_dir, "students.csv")
        self.reports_dir = os.path.join(self.data_dir, "report_cards")
        
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
        print(f"✓ Created data file at {self.csv_file}")
    
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
    
    def add_student(self):
        """Add new student with marks"""
        print("\n" + "="*50)
        print("ADD NEW STUDENT")
        print("="*50)
        
        try:
            student_id = input("Enter Student ID: ").strip()
            name = input("Enter Student Name: ").strip()
            
            # Input marks for each subject
            print("\nEnter marks for each subject (out of 100):")
            math = float(input("Mathematics: "))
            science = float(input("Science: "))
            english = float(input("English: "))
            social = float(input("Social Studies: "))
            computer = float(input("Computer: "))
            
            # Validate marks
            marks = [math, science, english, social, computer]
            if any(mark < 0 or mark > 100 for mark in marks):
                print("❌ Error: Marks must be between 0 and 100!")
                return
            
            # Calculate total and percentage
            total = sum(marks)
            percentage = total / 5
            grade = self.calculate_grade(percentage)
            
            # Save to CSV
            with open(self.csv_file, 'a', newline='') as file:
                writer = csv.writer(file)
                writer.writerow([student_id, name, math, science, english, 
                               social, computer, total, round(percentage, 2), grade])
            
            print(f"\n✓ Student {name} added successfully!")
            print(f"  Total: {total}/500 | Percentage: {percentage:.2f}% | Grade: {grade}")
            
        except ValueError:
            print("❌ Error: Please enter valid numeric marks!")
        except Exception as e:
            print(f"❌ Error: {e}")
    
    def view_all_students(self):
        """Display all students in a formatted table"""
        print("\n" + "="*50)
        print("ALL STUDENTS RECORDS")
        print("="*50)
        
        try:
            with open(self.csv_file, 'r') as file:
                reader = csv.reader(file)
                data = list(reader)
                
                if len(data) <= 1:
                    print("No student records found!")
                    return
                
                # Create prettytable
                table = PrettyTable()
                table.field_names = data[0]
                
                for row in data[1:]:
                    table.add_row(row)
                
                print(table)
                
        except Exception as e:
            print(f"❌ Error: {e}")
    
    def search_student(self):
        """Search and display specific student"""
        print("\n" + "="*50)
        print("SEARCH STUDENT")
        print("="*50)
        
        student_id = input("Enter Student ID to search: ").strip()
        
        try:
            with open(self.csv_file, 'r') as file:
                reader = csv.DictReader(file)
                found = False
                
                for row in reader:
                    if row['Student_ID'] == student_id:
                        found = True
                        print(f"\n✓ Student Found!")
                        print(f"ID: {row['Student_ID']}")
                        print(f"Name: {row['Name']}")
                        print(f"Math: {row['Math']}")
                        print(f"Science: {row['Science']}")
                        print(f"English: {row['English']}")
                        print(f"Social Studies: {row['Social_Studies']}")
                        print(f"Computer: {row['Computer']}")
                        print(f"Total: {row['Total']}/500")
                        print(f"Percentage: {row['Percentage']}%")
                        print(f"Grade: {row['Grade']}")
                        break
                
                if not found:
                    print(f"❌ No student found with ID: {student_id}")
                    
        except Exception as e:
            print(f"❌ Error: {e}")
    
    def generate_report_card(self):
        """Generate formatted report card for a student"""
        print("\n" + "="*50)
        print(""*15+"Sethupathy Government Arts College, Ramanathapuram")
        print("\n" + "="*50)
        print("GENERATE REPORT CARD")
        print("="*50)
        
        student_id = input("Enter Student ID: ").strip()
        
        try:
            with open(self.csv_file, 'r') as file:
                reader = csv.DictReader(file)
                found = False
                
                for row in reader:
                    if row['Student_ID'] == student_id:
                        found = True
                        
                        # Create report card
                        report = []
                        report.append("\n" + "="*60)
                        report.append(" " * 15 + "Sethupathy Government Arts College, Ramanathapuram")
                        report.append("\n" + "="*60)
                        report.append(" " * 15 + "STUDENT REPORT CARD")
                        report.append("="*60)
                        report.append(f"\nStudent ID: {row['Student_ID']}")
                        report.append(f"Name: {row['Name']}")
                        report.append(f"Date: {datetime.now().strftime('%Y-%m-%d')}")
                        report.append("\n" + "-"*60)
                        
                        # Create subject marks table
                        table = PrettyTable()
                        table.field_names = ["Subject", "Marks Obtained", "Max Marks", "Grade"]
                        
                        subjects = [
                            ("Mathematics", row['Math']),
                            ("Science", row['Science']),
                            ("English", row['English']),
                            ("Social Studies", row['Social_Studies']),
                            ("Computer", row['Computer'])
                        ]
                        
                        for subject, marks in subjects:
                            marks_float = float(marks)
                            subject_grade = self.calculate_grade(marks_float)
                            table.add_row([subject, marks, "100", subject_grade])
                        
                        report.append(str(table))
                        report.append("\n" + "-"*60)
                        report.append(f"Total Marks: {row['Total']}/500")
                        report.append(f"Percentage: {row['Percentage']}%")
                        report.append(f"Overall Grade: {row['Grade']}")
                        report.append("="*60)
                        
                        # Grade scale
                        report.append("\nGRADE SCALE:")
                        report.append("A+ (90-100) | A (80-89) | B (70-79)")
                        report.append("C (60-69) | D (50-59) | F (Below 50)")
                        report.append("="*60)
                        
                        # Print to console
                        report_text = "\n".join(report)
                        print(report_text)
                        
                        # Save to file
                        filename = f"report_card_{student_id}_{row['Name'].replace(' ', '_')}.txt"
                        filepath = os.path.join(self.reports_dir, filename)
                        
                        with open(filepath, 'w') as report_file:
                            report_file.write(report_text)
                        
                        print(f"\n✓ Report card saved to: {filepath}")
                        break
                
                if not found:
                    print(f"❌ No student found with ID: {student_id}")
                    
        except Exception as e:
            print(f"❌ Error: {e}")
    
    def display_menu(self):
        """Display main menu"""
        print("\n" + "="*50)
        print(" " * 12 + "STUDENT REPORT CARD SYSTEM")
        print("="*50)
        print("1. Add New Student")
        print("2. View All Students")
        print("3. Search Student by ID")
        print("4. Generate Report Card")
        print("5. Exit")
        print("="*50)
    
    def run(self):
        """Main program loop"""
        print("\n" + "*"*50)
        print(" " * 8 + "Welcome to Report Card System!")
        print("*"*50)
        
        while True:
            self.display_menu()
            choice = input("\nEnter your choice (1-5): ").strip()
            
            if choice == '1':
                self.add_student()
            elif choice == '2':
                self.view_all_students()
            elif choice == '3':
                self.search_student()
            elif choice == '4':
                self.generate_report_card()
            elif choice == '5':
                print("\n" + "="*50)
                print(" " * 12 + "Thank you for using!")
                print("="*50)
                break
            else:
                print("❌ Invalid choice! Please enter 1-5.")
            
            input("\nPress Enter to continue...")

# Run the program
if __name__ == "__main__":
    system = ReportCardSystem()
    system.run()

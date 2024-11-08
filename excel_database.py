import pandas as pd
from datetime import datetime
import re

class ExcelStudentDatabase:
    def __init__(self, excel_file='student_database.xlsx'):
        """Initialize Excel database connection"""
        self.excel_file = excel_file
        try:
            self.df = pd.read_excel(self.excel_file)
        except FileNotFoundError:
            # Create new Excel file if it doesn't exist
            self.df = pd.DataFrame(columns=[
                'student_id', 'first_name', 'last_name', 'date_of_birth',
                'email', 'phone', 'address', 'enrollment_date', 'status'
            ])
            self.save_database()

    def save_database(self):
        """Save changes to Excel file"""
        self.df.to_excel(self.excel_file, index=False)

    def add_student(self, first_name, last_name, date_of_birth, email, phone, address):
        """Add a new student to the database"""
        try:
            new_id = 1 if self.df.empty else self.df['student_id'].max() + 1
            new_student = {
                'student_id': new_id,
                'first_name': first_name,
                'last_name': last_name,
                'date_of_birth': date_of_birth,
                'email': email,
                'phone': phone,
                'address': address,
                'enrollment_date': datetime.now().strftime('%Y-%m-%d'),
                'status': 'Active'
            }
            self.df = pd.concat([self.df, pd.DataFrame([new_student])], ignore_index=True)
            self.save_database()
            return new_id
        except Exception as e:
            print(f"Error adding student: {e}")
            return None

    def get_all_students(self):
        """Get all students from database"""
        return self.df.to_dict('records')

    def delete_student(self, student_id):
        """Delete student from database"""
        try:
            self.df = self.df[self.df['student_id'] != student_id]
            self.save_database()
            return True
        except Exception as e:
            print(f"Error deleting student: {e}")
            return False

    def search_students(self, search_term):
        """Search students by name or email"""
        try:
            return self.df[
                self.df['first_name'].str.contains(search_term, case=False, na=False) |
                self.df['last_name'].str.contains(search_term, case=False, na=False) |
                self.df['email'].str.contains(search_term, case=False, na=False)
            ].to_dict('records')
        except Exception as e:
            print(f"Error searching students: {e}")
            return []

    def backup_database(self, backup_filename):
        """Create a backup of the database"""
        try:
            self.df.to_excel(backup_filename, index=False)
            return True
        except Exception as e:
            print(f"Backup error: {e}")
            return False

    def restore_database(self, backup_filename):
        """Restore database from backup"""
        try:
            self.df = pd.read_excel(backup_filename)
            self.save_database()
            return True
        except Exception as e:
            print(f"Restore error: {e}")
            return False
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
                'student_id', 'first_name', 'last_name', 'gender',
                'date_of_birth', 'email', 'phone', 'address',
                'enrollment_date', 'status'
            ])
            self.save_database()

    def save_database(self):
        """Save changes to Excel file"""
        self.df.to_excel(self.excel_file, index=False)

    def add_student(self, student_id, first_name, last_name, gender,
                    date_of_birth, email, phone, address, status='Active'):
        """Add a new student to the database"""
        try:
            # If no student_id provided, generate new one
            if student_id is None:
                student_id = 1 if self.df.empty else self.df['student_id'].max() + 1

            # Check if ID already exists
            if not self.df.empty and student_id in self.df['student_id'].values:
                raise ValueError("Student ID already exists")

            new_student = {
                'student_id': student_id,
                'first_name': first_name,
                'last_name': last_name,
                'gender': gender,
                'date_of_birth': date_of_birth,
                'email': email,
                'phone': phone,
                'address': address,
                'enrollment_date': datetime.now().strftime('%Y-%m-%d'),
                'status': status
            }

            # Convert all values to string before adding to DataFrame
            new_student = {k: str(v) if pd.notnull(v) else '' for k, v in new_student.items()}
            self.df = pd.concat([self.df, pd.DataFrame([new_student])], ignore_index=True)
            self.save_database()
            return True
        except Exception as e:
            print(f"Error adding student: {e}")
            return False

    def update_student(self, student_id, **kwargs):
        """Update existing student information"""
        try:
            student_mask = self.df['student_id'] == int(student_id)
            if not any(student_mask):
                raise ValueError("Student not found")

            for key, value in kwargs.items():
                if key in self.df.columns:
                    self.df.loc[student_mask, key] = str(value)

            self.save_database()
            return True
        except Exception as e:
            print(f"Error updating student: {e}")
            return False

    def get_all_students(self):
        """Get all students from database"""
        # Convert all values to string before returning
        df_str = self.df.astype(str)
        return df_str.to_dict('records')

    def delete_student(self, student_id):
        """Delete student from database"""
        try:
            self.df = self.df[self.df['student_id'] != student_id]
            self.save_database()
            return True
        except Exception as e:
            print(f"Error deleting student: {e}")
            return False

    def search_students(self, search_term, column=None):
        """Search students by any column"""
        try:
            # Convert search term to string
            search_term = str(search_term).lower()

            if column and column in self.df.columns:
                # Search in specific column
                mask = self.df[column].astype(str).str.lower().str.contains(search_term, na=False)
            else:
                # Search in all columns
                mask = pd.Series(False, index=self.df.index)
                for col in self.df.columns:
                    mask |= self.df[col].astype(str).str.lower().str.contains(search_term, na=False)

            # Convert results to string before returning
            result_df = self.df[mask].astype(str)
            return result_df.to_dict('records')
        except Exception as e:
            print(f"Error searching students: {e}")
            return []

    def validate_email(self, email):
        """Validate email format"""
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return bool(re.match(pattern, email))

    def validate_phone(self, phone):
        """Validate phone number format"""
        pattern = r'^\+?1?\d{9,15}$'
        return bool(re.match(pattern, phone))

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
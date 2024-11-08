import pandas as pd
from datetime import datetime
import re


class ExcelStudentDatabase:
    def __init__(self, excel_file='student_database.xlsx'):
        """Initialize Excel database connection"""
        self.excel_file = excel_file
        try:
            self.df = pd.read_excel(self.excel_file)
            # Convert student_id to integer type
            self.df['student_id'] = self.df['student_id'].astype(int)
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
        # Ensure student_id is integer before saving
        if not self.df.empty:
            self.df['student_id'] = self.df['student_id'].astype(int)
        self.df.to_excel(self.excel_file, index=False)

    def add_student(self, student_id, first_name, last_name, gender,
                    date_of_birth, email, phone, address, status='Active'):
        """Add a new student to the database"""
        try:
            # Handle student_id
            if student_id is None:
                student_id = 1 if self.df.empty else int(self.df['student_id'].max()) + 1
            else:
                student_id = int(student_id)

            # Check for duplicate ID
            if not self.df.empty and student_id in self.df['student_id'].values:
                raise ValueError("Student ID already exists")

            new_student = {
                'student_id': student_id,
                'first_name': str(first_name),
                'last_name': str(last_name),
                'gender': str(gender),
                'date_of_birth': date_of_birth,
                'email': str(email),
                'phone': str(phone),
                'address': str(address),
                'enrollment_date': datetime.now().strftime('%Y-%m-%d'),
                'status': str(status)
            }

            # Convert to DataFrame and ensure student_id is integer
            new_df = pd.DataFrame([new_student])
            new_df['student_id'] = new_df['student_id'].astype(int)

            self.df = pd.concat([self.df, new_df], ignore_index=True)
            self.save_database()
            return True
        except Exception as e:
            print(f"Error adding student: {e}")
            return False

    def get_all_students(self):
        """Get all students from database"""
        try:
            # Ensure student_id is integer
            if not self.df.empty:
                self.df['student_id'] = self.df['student_id'].astype(int)
            # Convert all values to string for display
            df_display = self.df.astype(str)
            return df_display.to_dict('records')
        except Exception as e:
            print(f"Error getting students: {e}")
            return []

    def delete_student(self, student_id):
        """Delete student from database"""
        try:
            # Convert student_id to integer for comparison
            student_id = int(student_id)
            # Convert DataFrame student_id to integer
            self.df['student_id'] = self.df['student_id'].astype(int)

            if student_id in self.df['student_id'].values:
                self.df = self.df[self.df['student_id'] != student_id]
                self.save_database()
                return True
            return False
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

            # Ensure student_id is integer in results
            result_df = self.df[mask].copy()
            result_df['student_id'] = result_df['student_id'].astype(int)
            # Convert to string for display
            result_df = result_df.astype(str)
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
            # Ensure student_id is integer after restore
            if not self.df.empty:
                self.df['student_id'] = self.df['student_id'].astype(int)
            self.save_database()
            return True
        except Exception as e:
            print(f"Restore error: {e}")
            return False
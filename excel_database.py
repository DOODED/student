import pandas as pd
from datetime import datetime
import re

class ExcelStudentDatabase:
    def __init__(self, excel_file='student_database.xlsx'):
        self.excel_file = excel_file
        try:
            self.df = pd.read_excel(self.excel_file)
        except FileNotFoundError:
            self.df = pd.DataFrame(columns=[
                'student_id', 'first_name', 'last_name', 'gender',
                'date_of_birth', 'email', 'phone', 'address',
                'enrollment_date', 'status'
            ])
            self.save_database()

    def save_database(self):
        self.df.to_excel(self.excel_file, index=False)

    def add_student(self, first_name, last_name, gender, date_of_birth,
                   email, phone, address):
        try:
            new_id = 1 if self.df.empty else self.df['student_id'].max() + 1
            new_student = {
                'student_id': new_id,
                'first_name': first_name,
                'last_name': last_name,
                'gender': gender,
                'date_of_birth': date_of_birth,
                'email': email,
                'phone': phone,
                'address': address,
                'enrollment_date': datetime.now().strftime('%Y-%m-%d'),
                'status': 'Active'
            }
            self.df = pd.concat([self.df, pd.DataFrame([new_student])],
                              ignore_index=True)
            self.save_database()
            return True
        except Exception as e:
            print(f"Error adding student: {e}")
            return False

    def get_all_students(self):
        return self.df.to_dict('records')

    def delete_student(self, student_id):
        try:
            self.df = self.df[self.df['student_id'] != student_id]
            self.save_database()
            return True
        except Exception as e:
            print(f"Error deleting student: {e}")
            return False

    def search_students(self, search_term, column=None):
        try:
            if column and column in self.df.columns:
                return self.df[
                    self.df[column].astype(str).str.contains(
                        str(search_term), case=False, na=False
                    )
                ].to_dict('records')
            else:
                mask = pd.Series(False, index=self.df.index)
                for col in self.df.columns:
                    mask |= self.df[col].astype(str).str.contains(
                        str(search_term), case=False, na=False
                    )
                return self.df[mask].to_dict('records')
        except Exception as e:
            print(f"Error searching students: {e}")
            return []

    def validate_email(self, email):
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return bool(re.match(pattern, email))

    def validate_phone(self, phone):
        pattern = r'^\+?1?\d{9,15}$'
        return bool(re.match(pattern, phone))
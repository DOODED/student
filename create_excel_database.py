# main.py
import sys
from PyQt5 import QtWidgets, uic
from PyQt5.QtWidgets import QMessageBox, QTableWidgetItem, QMainWindow, QDialog
from PyQt5.QtCore import Qt
import pandas as pd
from datetime import datetime
from excel_database import ExcelStudentDatabase


class AddStudentDialog(QDialog):
    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self.setupUi()

    def setupUi(self):
        self.setWindowTitle("Add New Student")
        self.setFixedSize(400, 500)

        layout = QtWidgets.QVBoxLayout()
        form_layout = QtWidgets.QFormLayout()

        # Create input fields
        self.first_name = QtWidgets.QLineEdit()
        self.last_name = QtWidgets.QLineEdit()
        self.gender = QtWidgets.QComboBox()
        self.gender.addItems(["Male", "Female"])
        self.dob = QtWidgets.QDateEdit()
        self.email = QtWidgets.QLineEdit()
        self.phone = QtWidgets.QLineEdit()
        self.address = QtWidgets.QTextEdit()

        # Add fields to form layout
        form_layout.addRow("First Name:", self.first_name)
        form_layout.addRow("Last Name:", self.last_name)
        form_layout.addRow("Gender:", self.gender)
        form_layout.addRow("Date of Birth:", self.dob)
        form_layout.addRow("Email:", self.email)
        form_layout.addRow("Phone:", self.phone)
        form_layout.addRow("Address:", self.address)

        # Create buttons
        button_layout = QtWidgets.QHBoxLayout()
        self.save_button = QtWidgets.QPushButton("Save")
        self.cancel_button = QtWidgets.QPushButton("Cancel")

        self.save_button.clicked.connect(self.save_student)
        self.cancel_button.clicked.connect(self.reject)

        button_layout.addWidget(self.save_button)
        button_layout.addWidget(self.cancel_button)

        # Add layouts to main layout
        layout.addLayout(form_layout)
        layout.addLayout(button_layout)
        self.setLayout(layout)

    def save_student(self):
        try:
            student_data = {
                'first_name': self.first_name.text(),
                'last_name': self.last_name.text(),
                'gender': self.gender.currentText(),
                'date_of_birth': self.dob.date().toPyDate(),
                'email': self.email.text(),
                'phone': self.phone.text(),
                'address': self.address.toPlainText()
            }

            if self.db.add_student(**student_data):
                QMessageBox.information(self, "Success", "Student added successfully!")
                self.accept()
            else:
                QMessageBox.critical(self, "Error", "Failed to add student")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to add student: {str(e)}")


class StudentManagementSystem(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi()
        self.db = ExcelStudentDatabase()
        self.setup_connections()

    def setupUi(self):
        self.setWindowTitle("Student Management System")
        self.setFixedSize(1200, 800)

        # Create central widget and main layout
        central_widget = QtWidgets.QWidget()
        self.setCentralWidget(central_widget)
        layout = QtWidgets.QVBoxLayout(central_widget)

        # Create stacked widget for multiple pages
        self.stacked_widget = QtWidgets.QStackedWidget()
        layout.addWidget(self.stacked_widget)

        # Create login page
        login_page = QtWidgets.QWidget()
        login_layout = QtWidgets.QVBoxLayout(login_page)

        login_form = QtWidgets.QFormLayout()
        self.username_input = QtWidgets.QLineEdit()
        self.password_input = QtWidgets.QLineEdit()
        self.password_input.setEchoMode(QtWidgets.QLineEdit.Password)

        login_form.addRow("Username:", self.username_input)
        login_form.addRow("Password:", self.password_input)

        self.login_button = QtWidgets.QPushButton("Login")

        login_layout.addLayout(login_form)
        login_layout.addWidget(self.login_button)

        # Create main page
        main_page = QtWidgets.QWidget()
        main_layout = QtWidgets.QVBoxLayout(main_page)

        # Search area
        search_layout = QtWidgets.QHBoxLayout()
        self.search_combo = QtWidgets.QComboBox()
        self.search_combo.addItems(["ID", "Gender", "Name"])
        self.search_input = QtWidgets.QLineEdit()
        self.search_button = QtWidgets.QPushButton("Search")

        search_layout.addWidget(self.search_combo)
        search_layout.addWidget(self.search_input)
        search_layout.addWidget(self.search_button)

        # Table
        self.student_table = QtWidgets.QTableWidget()
        self.student_table.setColumnCount(8)
        self.student_table.setHorizontalHeaderLabels([
            "ID", "First Name", "Last Name", "Gender",
            "DOB", "Email", "Phone", "Status"
        ])

        # Buttons
        button_layout = QtWidgets.QHBoxLayout()
        self.add_button = QtWidgets.QPushButton("Add Student")
        self.delete_button = QtWidgets.QPushButton("Delete Student")
        self.refresh_button = QtWidgets.QPushButton("Refresh")

        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.delete_button)
        button_layout.addWidget(self.refresh_button)

        # Add widgets to main layout
        main_layout.addLayout(search_layout)
        main_layout.addWidget(self.student_table)
        main_layout.addLayout(button_layout)

        # Add pages to stacked widget
        self.stacked_widget.addWidget(login_page)
        self.stacked_widget.addWidget(main_page)

        # Start with login page
        self.stacked_widget.setCurrentIndex(0)

    def setup_connections(self):
        self.login_button.clicked.connect(self.handle_login)
        self.add_button.clicked.connect(self.show_add_dialog)
        self.delete_button.clicked.connect(self.delete_student)
        self.search_button.clicked.connect(self.search_students)
        self.refresh_button.clicked.connect(self.load_student_data)

    def handle_login(self):
        username = self.username_input.text()
        password = self.password_input.text()

        if username == "admin" and password == "admin123":
            self.stacked_widget.setCurrentIndex(1)
            self.load_student_data()
        else:
            QMessageBox.warning(self, "Error", "Invalid credentials!")

    def load_student_data(self):
        try:
            students = self.db.get_all_students()
            self.student_table.setRowCount(len(students))

            for i, student in enumerate(students):
                self.student_table.setItem(i, 0, QTableWidgetItem(str(student['student_id'])))
                self.student_table.setItem(i, 1, QTableWidgetItem(student['first_name']))
                self.student_table.setItem(i, 2, QTableWidgetItem(student['last_name']))
                self.student_table.setItem(i, 3, QTableWidgetItem(student.get('gender', '')))
                self.student_table.setItem(i, 4, QTableWidgetItem(str(student['date_of_birth'])))
                self.student_table.setItem(i, 5, QTableWidgetItem(student['email']))
                self.student_table.setItem(i, 6, QTableWidgetItem(student['phone']))
                self.student_table.setItem(i, 7, QTableWidgetItem(student['status']))
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load data: {str(e)}")

    def show_add_dialog(self):
        dialog = AddStudentDialog(self.db, self)
        if dialog.exec_() == QDialog.Accepted:
            self.load_student_data()

    def delete_student(self):
        selected = self.student_table.selectedItems()
        if not selected:
            QMessageBox.warning(self, "Warning", "Please select a student to delete")
            return

        row = selected[0].row()
        student_id = int(self.student_table.item(row, 0).text())

        reply = QMessageBox.question(self, "Confirm",
                                     "Are you sure you want to delete this student?",
                                     QMessageBox.Yes | QMessageBox.No)

        if reply == QMessageBox.Yes:
            if self.db.delete_student(student_id):
                self.load_student_data()
                QMessageBox.information(self, "Success", "Student deleted successfully!")
            else:
                QMessageBox.critical(self, "Error", "Failed to delete student")

    def search_students(self):
        search_text = self.search_input.text().lower()
        search_column = self.search_combo.currentText().lower()

        for row in range(self.student_table.rowCount()):
            match = False
            item = self.student_table.item(row, self.get_column_index(search_column))

            if item and search_text in item.text().lower():
                match = True

            self.student_table.setRowHidden(row, not match)

    def get_column_index(self, column_name):
        column_mapping = {
            'id': 0,
            'gender': 3,
            'name': 1
        }
        return column_mapping.get(column_name, 0)


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = StudentManagementSystem()
    window.show()
    sys.exit(app.exec_())

import pandas as pd
import os
from datetime import datetime
import calendar
from PySide6.QtWidgets import QApplication, QWidget, QTableWidget, QTableWidgetItem, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit, QHeaderView, QLabel, QCalendarWidget, QCheckBox, QMessageBox
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QDateEdit
from PySide6.QtCore import QDate
from PySide6.QtGui import QPixmap  # Import QPixmap for handling images
from num2words import num2words  # We will use this package to convert numbers to words
from openpyxl import load_workbook  # Import openpyxl for handling templates
from PySide6.QtWidgets import QApplication, QWidget, QTableWidget, QTableWidgetItem, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit, QHeaderView, QLabel, QCalendarWidget, QCheckBox, QMessageBox, QGroupBox
from PySide6.QtCore import Qt, QDate
from PySide6.QtGui import QPixmap, QColor
import pyperclip  # Required for clipboard functionality
import time




import extras
import my_gui
import main_processing
import mailing

# Global variable for button width size
# BUTTON_WIDTH = 10  # Set the button width globally






class Design:
    """Class to centralize UI layout settings for easy management."""
    def __init__(self):
        self.button_width = 150
        self.logo_width = 150
        self.logo_height = 100
        self.table_column_widths = [100]
        self.table_margin_left = 100
        self.table_margin_right = 100
        self.date_picker_width = 120
        self.spacing_between_title_and_content = 20
        self.spacing_between_components = 10
        self.bullet_style = "font-size: 16px; color: gray;"  # Initial bullet style (gray)


    def get_button_width(self):
        return self.button_width
    
    def get_table_column_widths(self):
        return self.table_column_widths
    
    def get_logo_dimensions(self):
        return (self.logo_width, self.logo_height)
    
    def get_table_margins(self):
        return (self.table_margin_left, self.table_margin_right)
    
    def get_date_picker_width(self):
        return self.date_picker_width
    
    def get_spacing_between_title_and_content(self):
        return self.spacing_between_title_and_content
    
    def get_spacing_between_components(self):
        return self.spacing_between_components

    def get_bullet_style(self):
        return self.bullet_style
    
# def update_bullet_point(bullet_label, is_success, step_name):
#     """Update bullet point color and style"""
#     if is_success:
#         bullet_label.setStyleSheet("font-size: 16px; color: green; font-weight: bold;")
#         bullet_label.setText(f"✔ {step_name}")
#     else:
#         bullet_label.setStyleSheet("font-size: 16px; color: red; font-weight: bold;")
#         bullet_label.setText(f"✘ {step_name}")




# def show_message_box(success, message):
#     """Show a message box with success or failure message"""
#     msg = QMessageBox()
#     msg.setWindowTitle("Status")
#     msg.setText(message)
#     # Load a custom green checkmark icon
#     if success:
#         pixmap = QPixmap('../Data/greencheck.jpg')  # Replace with your image path
#         pixmap = pixmap.scaled(50, 50)  # Set width and height (50x50 in this case)

#         msg.setIconPixmap(pixmap)
#     else:
#         msg.setIcon(QMessageBox.Critical)    
#     msg.exec()





def init_gui(employee_data, design):
    try:
        if employee_data is None:
            extras.print_colored("No valid employee data found. Cannot initialize the GUI.", "red")

            # print("No valid employee data found. Cannot initialize the GUI.")
            return None

        # Create the main window
        window = QWidget()
        window.setWindowTitle("Invoice Management System")
        window.setGeometry(100, 100, 1200, 600)  # Adjusted window size

        window.setWindowState(Qt.WindowMaximized)


        # Title Layout with increased font size for the label
        title_layout = QHBoxLayout()
        title_label = QLabel("Invoice Generation System")
        # title_label.setAlignment(Qt.AlignRight)
        title_label.setStyleSheet("""
            QLabel {
                font-size: 24px; /* Font size */
                font-weight: bold; /* Bold text */
                margin-left:600px;
            }
        """)
        title_layout.addWidget(title_label)


        # Create the logo with specified dimensions from Design class
        logo_label = QLabel()  # Placeholder for logo
        logo_pixmap = QPixmap("../Data/logo.jpg")  # Load the logo image
        scaled_logo = logo_pixmap.scaled(design.get_logo_dimensions()[0], design.get_logo_dimensions()[1], Qt.AspectRatioMode.KeepAspectRatio)
        logo_label.setPixmap(scaled_logo)
        title_layout.addWidget(logo_label, alignment=Qt.AlignRight)


        # Get the default invoice date (last day of previous month)
        default_invoice_date = extras.get_last_day_of_previous_month()
        

        # Create a small date picker (QDateEdit) for invoice date
        date_picker = QDateEdit()
        date_picker.setDisplayFormat("yyyy-MM-dd")
        date_picker.setDate(default_invoice_date)  # Set the default date to the last day of the previous month
        date_picker.setCalendarPopup(True)  # Enable the calendar pop-up on click
        date_picker.setMaximumWidth(380)


        # Styles for the QDateEdit (date_picker)
        date_picker.setStyleSheet("""
            QDateEdit {
                margin-left:265px;
                width:10px;
            }
        """)



        # Create the table
        table = QTableWidget()
        table.setColumnCount(7)  # 7 + 1 for the new "Email Remark" column
        table.setHorizontalHeaderLabels(["Select", "Name", "Invoice No", "Payable Days", "Food Allowance Days", "Calculated Salary", "Email Remark"])

        # font-size: 14px;
        # font-family: Arial, sans-serif;

        table.setStyleSheet("""
            /* General table styling */
            QTableWidget {
                background-color: #FFFFFF;
                border: 5px solid #CCCCCC;
                gridline-color: #E0E0E0;
                alternate-background-color: #F9F9F9;
                border-radius:15px;
                width:50%;
                margin-left:200px;
                margin-right:200px;

            }

            /* Header styling */
            QTableWidget::horizontalHeader {
                background-color: #EFEFEF;
                border: none;
                font-weight: bold;
                font-size: 16px;
                color: #333;
                                            border-radius:15px;

            }

            /* Header section (for each column) */
            QTableWidget::horizontalHeader::section {
                padding: 10px;
                border-right: 1px solid #CCCCCC;
                background-color: #D7D7D7;
                text-align: center;
                                                                        border-radius:15px;

            }

            /* Row selection styling */
            QTableWidget::item:selected {
                background-color: #D6EAF8;
                color: #333;
                                                                        border-radius:15px;

            }

            /* Styling for individual table cells */
            # QTableWidget::item {
            #     padding: 8px;
            #     border: 1px solid #E0E0E0;
            #     color: #333;
                                                                        border-radius:15px;

            # }

            /* Alternating row colors */
            QTableWidget::item:alternate {
                background-color: #F9F9F9;
                                                                        border-radius:15px;

            }

            /* Styling for the checkbox column */
            # QTableWidget QCheckBox {
            #     padding: 5px;
            #     width:10px;
            #     margin-right:-50px;
            # }

            # /* Styling for input fields in the table (QLineEdit) */
            # QTableWidget QLineEdit {
            #     background-color: #F0F0F0;
            #     border: 1px solid #CCCCCC;
            #     padding: 5px;
                                                                        border-radius:15px;

            #     border-radius: 4px;
            # }

            /* Styling for the last column's "Email Remark" field */
            QTableWidget::item:nth-child(1) {
                background-color: #FFFFFF;
                border: 5px solid #CCCCCC;
                gridline-color: #E0E0E0;
                alternate-background-color: #F9F9F9;

                width:-50%;
                
            }
        """)


        for _, row in employee_data.iterrows():
            invoice_no = extras.generate_invoice_no(row["Last_Invoice_No"])
            row_position = table.rowCount()
            table.insertRow(row_position)

            # Add checkbox in the first column
            checkbox = QCheckBox()
            checkbox.setChecked(True)  # Default to checked
            table.setCellWidget(row_position, 0, checkbox)

            # Add Name and Invoice No
            table.setItem(row_position, 1, QTableWidgetItem(row["Employee_Name"]))
            table.setItem(row_position, 2, QTableWidgetItem(invoice_no))

            # Add Payable Days and Food Allowance Days as input fields (QLineEdit)
            payable_days_input = QLineEdit()
            payable_days_input.setText('')  # Default text
            table.setCellWidget(row_position, 3, payable_days_input)

            food_days_input = QLineEdit()
            food_days_input.setText('')  # Default text
            table.setCellWidget(row_position, 4, food_days_input)

            # Add Calculated Salary (initially empty)
            table.setItem(row_position, 5, QTableWidgetItem(''))  # Calculated Salary

            # Add "Email Remark" as input field
            email_remark_input = QLineEdit()
            table.setCellWidget(row_position, 6, email_remark_input)


            # Use helper function to set up listeners with correct row reference
            extras.setup_text_change_listeners(table, row_position, employee_data, payable_days_input, food_days_input, date_picker.date())

            # print("this is the table")
            # print(table)
            # print("this is the payable days input")
            # print(payable_days_input.text())


        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        # Create the Select/Deselect Toggle Button
        toggle_btn = QPushButton("Deselect All")
        toggle_btn.setFixedWidth(design.get_button_width())
        toggle_btn.clicked.connect(lambda: extras.toggle_select_all(table, toggle_btn))

        # toggle_btn.setStyleSheet("""
        #     QPushButton {
        #     margin-left:50px;
                                 
        #     }


        # """)





        calendar_label = QLabel("Invoice Date:")

        # Styles for the QLabel (calendar_label)
        calendar_label.setStyleSheet("""
            QLabel {
                font-size: 16px; /* Font size */
                font-weight: bold; /* Bold text */
                color: #333333; /* Dark gray text color */
                padding-right: 10px; /* Space between label and date picker */
                margin-left:270px;
            }
        """)




        # calendar_label.setAlignment(Qt.AlignLeft)
        
        # Create Submit and Email Buttons
        submit_btn = QPushButton("Generate Invoices")
        submit_btn.setFixedWidth(design.get_button_width())  # Apply global button width
        submit_btn.clicked.connect(lambda: main_processing.submit_data(table, employee_data, date_picker.date()))

        email_btn = QPushButton("Send Emails")
        email_btn.setFixedWidth(design.get_button_width())  # Apply global button width
        email_btn.clicked.connect(lambda: mailing.process_employee_data_for_email())

        # Create layout for the toggle button above the table
        button_layout = QHBoxLayout()
        button_layout.addWidget(toggle_btn)
        button_layout.addWidget(submit_btn)
        button_layout.addWidget(email_btn)
        
        # button_layout.addWidget(copy_btn)

        # Layout
        layout = QVBoxLayout()
        layout.addLayout(title_layout)  # This is assuming title_layout is a valid layout
        layout.addWidget(calendar_label)
        layout.addWidget(date_picker)  # Add the small date picker with margin
        layout.addLayout(button_layout)  # Use addLayout here instead of addWidget
        layout.addWidget(table)

        window.setLayout(layout)
        window.show()

        return window

    except Exception as e:
        extras.print_colored(f"Error initializing the GUI: {str(e)}", "red")
        return None


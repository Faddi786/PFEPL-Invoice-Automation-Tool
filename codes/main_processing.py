
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
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os


import extras
import my_gui
import main_processing






    
global store_employee_data
store_employee_data = pd.DataFrame()

def submit_data(table, employee_data, invoice_date):
    global store_employee_data
    try:
        # Call this function before starting any processing
        extras.show_command_prompt()
        extras.get_manager_info()

        # Collect the data from the table and process only checked rows
        for row in range(table.rowCount()):
            checkbox = table.cellWidget(row, 0)
            if checkbox.isChecked():
                employee_name = table.item(row, 1).text().strip()
                invoice_no = table.item(row, 2).text().strip()
                payable_days = table.cellWidget(row, 3).text().strip()
                food_days = table.cellWidget(row, 4).text().strip()
                # salary = table.item(row, 5).text().strip()

                email_remark = table.cellWidget(row, 6).text().strip()
                # email_remark == None or email_remark is None or 
                # if email_remark == "":
                #     email_remark = "-"
                # print("this is the email remark")
                # print(f"{email_remark}")

                if food_days is None or food_days == "":
                    food_days = "0"

                # print
                if not payable_days.isdigit() or not food_days.isdigit():
                    extras.print_colored("Error: Payable Days and Food Allowance Days must be numeric.", "red")

                    # print("Error: Payable Days and Food Allowance Days must be numeric.")
                    return

                payable_days = int(payable_days)
                food_days = int(food_days)


                # Check if the employee exists in the data
                matching_index = employee_data.loc[employee_data["Employee_Name"] == employee_name].index
                if matching_index.empty:
                    extras.print_colored(f"Error: Employee '{employee_name}' not found in data.", "red")
                    continue  # Skip this employee and move to the next row

                # Proceed with the data update
                employee_index = matching_index[0]
                employee_data.loc[employee_index,"Checked"] = "Yes"
                employee_data.loc[employee_index,"Last_Invoice_No"] = invoice_no
                employee_data.loc[employee_index,"Payable_Days"] = payable_days
                employee_data.loc[employee_index,"Food_Days"] =  food_days
                employee_data.loc[employee_index,"Email_Remark"] =  email_remark

                # print(employee_data)

                # Call to populate the template with the new data
                extras.populate_template(employee_data, employee_index, invoice_date)
                message = invoice_no
                column = "Last_Invoice_No"
                extras.update_email_status(employee_name,message,column)
        store_employee_data = employee_data
        extras.print_colored("\033[92m\n\nInvoices generated successfully!\033[0m", "green")

        # print("\033[92mInvoices generated successfully!\033[0m")

    except Exception as e:
        extras.print_colored(f"Submission Error: {str(e)}", "red")

        # print(f"Submission Error: {str(e)}")


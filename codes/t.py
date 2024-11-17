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



def main():
    try:
        
        # extras.strip_excel_whitespace()

        filepath = "../invoice_data.xlsx"
        employee_data = extras.load_employee_data(filepath)


        design = my_gui.Design()  # Create an instance of Design class

        app = QApplication([])

        window = my_gui.init_gui(employee_data, design)

        if window is None:
            extras.show_command_prompt()

            return


        extras.show_command_prompt()

        app.exec()


    except Exception as e:
        extras.print_colored(f"Application Error: {str(e)}", "red")
        extras.show_command_prompt()

        # print(f"Application Error: {str(e)}")

if __name__ == "__main__":
    main()



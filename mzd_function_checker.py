#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Author: Angelo Matthew Quilapio
License: Apache License 2.0
GitHub Portfolio Project
--------------------------------------------
This script scans a target directory for ZIP files
and compiles their filenames into an Excel file named
'Compiled_Function_Names.xlsx'.

Purpose:
- Demonstrates basic file handling (os module)
- Demonstrates Excel file creation using openpyxl
- Example project for GitHub portfolio
--------------------------------------------
"""

import os
import pandas as pd
from openpyxl import Workbook

# --------------------------------------------
# Global variable to store ZIP filenames
# --------------------------------------------
array_out = []

# --------------------------------------------
# Function: create_CSV
# Description:
#   Adds a found ZIP filename (without extension)
#   into an Excel file for record-keeping.
# --------------------------------------------
def create_CSV(target_zip):
    # Add the ZIP filename to the list
    array_out.append(target_zip)

    # Write each collected filename to the Excel sheet
    for ctr in range(len(array_out)):
        work_sheet.cell(row=ctr + 1, column=1, value=array_out[ctr])

    # Save the Excel file
    work_book.save('Compiled_Function_Names.xlsx')


# --------------------------------------------
# Configuration: Directory to search for ZIP files
# --------------------------------------------
directory = r'C:\Users\Lenovo\Documents'

# Create a new Excel workbook and select the active sheet
work_book = Workbook()
work_sheet = work_book.active

# --------------------------------------------
# Main logic: Scan directory and collect ZIP filenames
# --------------------------------------------
for target_zip in os.listdir(directory):
    full_path = os.path.join(directory, target_zip)

    # Check if it's a file (not a directory)
    if os.path.isfile(full_path):
        check_file = target_zip.split('.')

        # Ensure it has an extension and that it's '.zip'
        if len(check_file) > 1 and check_file[1].lower() == 'zip':
            # Add filename (without extension) to Excel
            create_CSV(check_file[0])

# --------------------------------------------
# End of Script
# --------------------------------------------
print("ZIP filenames have been compiled into 'Compiled_Function_Names.xlsx'")

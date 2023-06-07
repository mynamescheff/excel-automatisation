from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from transposer import ExcelTransposer
from wide import Wide
from case_checker import CaseList
import os
import pandas as pd
from datetime import date
from acc_checker import ExcelComparator

SOURCE_DIR = ("C:\\IT project\\30.05")
excel_files = list(Path(SOURCE_DIR).glob("*.xlsx"))

values_excel_files = {}
for excel_file in excel_files:
    wb = load_workbook(filename = excel_file)
    extra_cell_1 = wb["Sheet 1"]["B19"]
    extra_cell_2 = wb["Sheet 1"]["C19"]
    extra_cell_3 = wb["Sheet 1"]["C33"]
    extra_cell_4 = wb["Sheet 1"]["C34"]
    extra_cell_5 = wb["Sheet 1"]["C35"]
    
    rng = wb["Sheet 1"]["B16":"B17"]
    rng_values = []
    for cells in rng:
        for cell in cells:
            rng_values.append(cell.value)

    extra_cell_1_value = extra_cell_1.value
    extra_cell_2_value = extra_cell_2.value
    extra_cell_3_value = extra_cell_3.value
    extra_cell_4_value = extra_cell_4.value
    extra_cell_5_value = extra_cell_5.value

    # Add concatenated values to values_excel_files
    values_excel_files[excel_file.name] = rng_values + [extra_cell_1_value, extra_cell_2_value, extra_cell_3_value, extra_cell_4_value, extra_cell_5_value]

workbook = Workbook()

worksheet = workbook.active

filename = "C:\\IT project\\test\\combined2.xlsx"
workbook.save(filename)

wb = load_workbook(filename = "C:\\IT project\\test\\combined2.xlsx")

header_list = [
    "uni name",
    "candidate name",
    "case nr",
    "amount",
    "currency",
    "acc number",
    "iban number",
    "swift/bic"
]

ws = wb.active

# Write the header list in column A
for i, header in enumerate(header_list):
    ws[f"A{i+1}"] = header

# Append values to the "combined" excel file starting from column B
for i, excel_file in enumerate(values_excel_files):
    column_letter = get_column_letter(i+2)  # Convert column index to letter
    ws[f"{column_letter}1"] = excel_file  # Write excel file name in the first row of the column
    for j, value in enumerate(values_excel_files[excel_file]):
        # Check if the current value is from cell C34
        if j == len(values_excel_files[excel_file]) - 2:
            # Check if value is not None before replacing spaces
            if value is not None:
                value = value.replace(" ", "")
        ws[f"{column_letter}{j+2}"] = value

wb.save("C:\\IT project\\test\\combined2.xlsx")

filename = "C:\\IT project\\test\\combined2.xlsx"
transposer = ExcelTransposer(filename)
transposer.transpose_cells_to_table()

# Adjust column width
wide = Wide(filename, "Transposed")
wide.auto_adjust_column_width()


excel_folder = "C:\\IT project\\30.05"
list_folder = "C:\\IT project\\case_list"

case_list = CaseList(excel_folder, list_folder)
case_list.process_excel_files()

comparator = ExcelComparator("C:\\IT project\\test\\combined2.xlsx", "C:\\IT project\\sprawdzacz.xlsx")
comparator.compare_and_append()


import os
from openpyxl import load_workbook, Workbook
from pathlib import Path
import shutil
from typing import List
from utils.modules import Sheet
import pandas as pd
from copy import copy
from utils.Utils import get_file_name_without_extension_from_file_path
from utils.modules import Sheet
from openpyxl.utils.dataframe import dataframe_to_rows


def copy_worksheet(source_ws, target_ws, copy_value: bool = True):
    for row in source_ws.iter_rows():
        for cell in row:
            if copy_value:
                target_ws[cell.coordinate].value = cell.value

            target_ws[cell.coordinate].font = copy(cell.font)
            target_ws[cell.coordinate].border = copy(cell.border)
            target_ws[cell.coordinate].fill = copy(cell.fill)
            target_ws[cell.coordinate].number_format = cell.number_format
            target_ws[cell.coordinate].protection = copy(cell.protection)
            target_ws[cell.coordinate].alignment = copy(cell.alignment)
            target_ws[cell.coordinate].comment = cell.comment

    # Copy cell width and height
    for idx, rd in source_ws.row_dimensions.items():
        target_ws.row_dimensions[idx] = copy(rd)

    for idx, rd in source_ws.column_dimensions.items():
        target_ws.column_dimensions[idx] = copy(rd)

    for merged_cell in source_ws.merged_cells:
        target_ws.merge_cells(f"{merged_cell}")


def merge_multiple_excels_to_one_excel(input_files: List, output_file):
    output_wb = Workbook()

    for file_name in input_files:
        source_wb = load_workbook(file_name)
        number_of_sheets = len(source_wb.sheetnames)

        for sheet in source_wb:
            file_name = get_file_name_without_extension_from_file_path(file_name)
            sheet_name = f"{file_name}_{sheet.title}"

            # if only 1 sheet, use file name instead
            if number_of_sheets == 1:
                sheet_name = file_name

            output_ws = output_wb.create_sheet(sheet_name)
            copy_worksheet(sheet, output_ws)

    # Delete the default sheet
    del output_wb["Sheet"]
    output_wb.save(output_file)


# merge wb2 into wb1
def merge_two_workbook(wb1: Workbook, wb2: Workbook):
    source_wb = wb2

    for sheet in source_wb:
        sheet_name = sheet.title

        output_ws = wb1.create_sheet(sheet_name)
        copy_worksheet(sheet, output_ws)


def write_sheet_to_worksheet(wb: Workbook, sheets: List[Sheet]):
    for sheet in sheets:
        ws = wb.create_sheet(sheet.name)
        rows = dataframe_to_rows(sheet.data_frame, index=False)

        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)


def save_file_as_xlsm(workbook: Workbook, result_file, empty_macro_file):
    macro_wb = load_workbook(empty_macro_file, keep_vba=True)
    merge_two_workbook(macro_wb, workbook)
    macro_wb.save(result_file)

import xlwings as xw
import pandas as pd
import numpy as np
import os
import utils
from typing import Union

def display(df: pd.DataFrame) -> None:
    xw.view(df)

def open_book(file_path = os.path.join("EXCEL_REPO", "file.xlsx")):
    return xw.Book(file_path)

def close_book(file_path = os.path.join("EXCEL_REPO", "file.xlsx")):
    return xw.Book(file_path).close()

def delete_column(sheet: xw.Sheet, xl_col_idx: int):
    try:
        cell_name = chr((xl_col_idx-1)+65) # xl has 1 based indexing
        sheet.range(f"{cell_name}:{cell_name}").api.Delete()
    except Exception as err:
        utils.log_xl_err(err)

def delete_col_range(sheet: xw.Sheet, top_row: int, top_col: int, bottom_row: int, bottom_col: int):
    """Delete a subset of the column

    Args:
        sheet (xw.Sheet): _description_
        top_row (int): _description_
        top_col (int): _description_
        bottom_row (int): _description_
        bottom_col (int): _description_
    """
    try:
        sheet.range((top_row, top_col), (bottom_row, bottom_col)).api.Delete()
    except Exception as err:
        utils.log_xl_err(err)

def add_column(sheet: xw.Sheet, data: list[object], xl_col_idx: int):
    try:
        cell_name = chr((xl_col_idx-1)+65) # xl has 1 based indexing
        sheet.range(f"{cell_name}:{cell_name}").api.Insert()
        sheet.range(f"{cell_name}1").options(transpose = True).value = utils.parse_lst(data)
    except Exception as err:
        utils.log_xl_err(err)

def del_row(sheet: xw.Sheet, row_idx: int) -> bool:
    try:
        sheet.range((row_idx, 1), (row_idx, sheet.cells.last_cell.column)).api.Delete()
        return True
    except Exception as err:
        utils.log_xl_err(err)
        return False

def fix_dates(sheet: xw.Sheet, col_range: Union[str|list[str]]) -> bool:
    """Fix the date formatting

    Args:
        sheet (xw.Sheet): xlwings sheet object
        col_range (Union[str | list[str]]): name of the column (eg: B) or range (eg: B2:B100)

    Returns:
        bool: acknowledgement
    """
    try:
        dates = sheet.range(col_range).value
        fixed_dates = pd.to_datetime(dates, format = "mixed")
        sheet.range(col_range).value = fixed_dates.to_list()
        return True
    except Exception as err:
        utils.log_xl_err(err)
        return False

def fill_missing_values(sheet: xw.Sheet, col_range: str) -> bool:
    """Fill missing values with mean (only for numeric columns)

    Args:
        sheet (xw.Sheet): sheet object
        col_range (Union[str | list[str]]): name of the column (eg: B) or range (eg: B2:B100)

    Returns:
        bool: acknowledgement
    """
    try:
        last_row = sheet.range(f"{col_range}{sheet.cells.last_cell.row}").end("up").row
        data = sheet.range(f"{col_range}2:{col_range}{last_row}").value
        if not data:
            return False

        mean = utils.calc_arr_mean(data)
        updated = utils.replace_missing(data, mean)

        sheet.range(f"{col_range}2:{col_range}{last_row}").options(transpose = True).value = updated
        return True
    except Exception as err:
        utils.log_xl_err(err)
        return False
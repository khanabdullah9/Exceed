import xlwings as xw
import pandas as pd
import os
import utils

def display(df: pd.DataFrame) -> None:
    xw.view(df)

def open_book(file_path = os.path.join("EXCEL_REPO", "file.xlsx")):
    return xw.Book(file_path)

def delete_column(sheet: xw.Sheet, xl_col_idx: int):
    try:
        cell_name = chr((xl_col_idx-1)+65) # xl has 1 based indexing
        sheet.range(f"{cell_name}:{cell_name}").api.Delete()
    except Exception as err:
        pass

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
        utils.log_err(err)

def add_column(sheet: xw.Sheet, data: list[object], xl_col_idx: int):
    try:
        cell_name = chr((xl_col_idx-1)+65) # xl has 1 based indexing
        sheet.range(f"{cell_name}:{cell_name}").api.Insert()
        sheet.range(f"{cell_name}1").options(transpose = True).value = utils.parse_lst(data)
    except Exception as err:
        utils.log_err(err)

def del_row(sheet: xw.Sheet, row_idx: int) -> bool:
    try:
        sheet.range((row_idx, 1), (row_idx, sheet.cells.last_cell.column)).api.Delete()
        return True
    except Exception as err:
        utils.log_err(err)
        return False
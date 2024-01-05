import re

from openpyxl.utils import column_index_from_string, get_column_letter

from bb_openpyxl._utils import _duplicate


def _new_range(start_row, start_column, end_row, end_column, row_idx, col_idx, amount):
    """计算新的受增删行或列影响单元格的范围"""
    new_start_column = _new_index(start_column, col_idx, amount)
    new_end_column = _new_index(end_column, col_idx, amount)
    new_start_row = _new_index(start_row, row_idx, amount)
    new_end_row = _new_index(end_row, row_idx, amount)
    return new_end_column, new_end_row, new_start_column, new_start_row


def _new_index(start_or_end, idx, amount):
    """计算新的受增删行或列影响单元格的行或者列起点"""
    return start_or_end + amount if idx is not None and idx <= start_or_end else start_or_end


def _new_col_string(col, idx, amount=1):
    """计算新的受增删行或列影响单元格的列，入参和出参都是字符串"""
    current = column_index_from_string(col)
    return get_column_letter(_new_index(current, idx, amount))


def _new_row_string(row, idx, amount=1):
    """计算新的受增删行或列影响单元格的行，入参和出参都是字符串"""
    current = int(row)
    return str(_new_index(current, idx, amount))


def _get_all_cells(formula_or_range):
    """提取公式或者范围内中的单元格"""
    cell_pattern = r"(?<!\w)([A-Z]+\d+)(?!\()"
    return _duplicate(re.findall(cell_pattern, formula_or_range))


def _get_all_rows(formula):
    """提取公式中的行"""
    row_pattern = r"\$(\d+)"
    return _duplicate(re.findall(row_pattern, formula))


def _get_all_columns(formula):
    """提取公式中的列"""
    column_pattern = r"\$([A-Z]+)"
    return _duplicate(re.findall(column_pattern, formula))

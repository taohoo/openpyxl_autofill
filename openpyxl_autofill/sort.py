# -*- coding: utf-8 -*-
"""
@author: hubo
@project: python311
@file: sort.py
@time: 2023/12/18 19:38
@desc: 排序
"""
import warnings


def sort(worksheet, start_row, start_column, end_row, end_column, sort_column, reverse=False):
    """
    Sorts a worksheet by a specified column. If the column contains formulas, the Excel file must be opened with the `data_only=True` parameter.
    对worksheet进行排序，如果排序字段为公式，必须用data_only=True的方式打开excel。
    Args:
        worksheet (object): The worksheet to be sorted.
        start_row (int): The starting row for sorting.
        start_column (int): The starting column for sorting.
        end_row (int): The ending row for sorting.
        end_column (int): The ending column for sorting.
        sort_column (int): The column to be used as the sorting criteria.
        reverse (bool): Whether to sort in reverse order.
    Returns:
        None
    """
    _sort_f(worksheet, start_row, start_column, end_row, end_column, key=lambda x: x[sort_column - 1].value, reverse=reverse)


def _sort_f(worksheet, start_row, start_column, end_row, end_column, key, reverse=False):
    """
    用指定的func进行排序，灵活应用可以解决公式的问题
    除了func，参数都和sort一致
    :param func:
    :return:
    """
    import copy
    # 把数据全部读取出来
    data = []
    for row in worksheet.iter_rows(min_row=start_row, max_row=end_row):
        _d = []
        for c in row:
            _d.append(copy.copy(c))
        data.append(_d)
    sorted_data = sorted(data, key=key, reverse=reverse)
    # 数据用覆盖的方式写回去
    for r in range(start_row, end_row + 1):
        for c in range(start_column, end_column + 1):
            s_cell = sorted_data[r - start_row][c - start_column]
            if s_cell.data_type == 'f':  # 公式，写警告
                warnings.warn('排序的单元格中的包含公式，排序之后，可能导致公式异常，为防止被忽略的数据错误，该公式被清除。')
                s_cell.value = None
            worksheet.cell(row=r, column=c).value = s_cell.value
            worksheet.cell(row=r, column=c)._style = s_cell._style

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
    对worksheet进行排序，如果排序字段为公式，必须用data_only=True的方式打开excel。
    :param worksheet: 要排序的worksheet 
    :param start_row: 排序开始的行
    :param start_column: 排序开始的列
    :param end_row: 排序结束的行
    :param end_column: 排序结束的列
    :param sort_column: 用来做排序依据的列
    :param reverse: 是否倒序
    :return: 
    """
    sort_f(worksheet, start_row, start_column, end_row, end_column, key=lambda x: x[sort_column - 1].value, reverse=reverse)


def sort_f(worksheet, start_row, start_column, end_row, end_column, key, reverse=False):
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
                warnings.warn('排序的单元格中的包含公式，排序之后，可能导致公式异常')
            worksheet.cell(row=r, column=c).value = s_cell.value
            worksheet.cell(row=r, column=c)._style = s_cell._style

# -*- coding: utf-8 -*-
"""
@author: hubo
@project: bb-py
@file: delete.py.py
@time: 2023/12/22 16:09
@desc:
"""
from .insert import insert_rows, insert_cols


def delete_rows(worksheet, idx, amount=1):
    """
    Delete rows from a worksheet.

    Args:
        worksheet (Worksheet): The worksheet from which to delete rows.
        idx (int): The index of the first row to delete.
        amount (int, optional): The number of rows to delete. Defaults to 1.

    Returns:
        None
    """
    return insert_rows(worksheet, idx, amount=-amount)


def delete_cols(worksheet, idx, amount=1):
    return insert_cols(worksheet, idx, amount=-amount)

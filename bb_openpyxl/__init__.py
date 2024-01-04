# -*- coding: utf-8 -*-
"""
@author: hubo
@project: bb-py
@file: __init__.py.py
@time: 2023/12/22 14:00
@desc:
"""
from openpyxl.worksheet.worksheet import Worksheet

from .insert import insert_rows, insert_cols
from .delete import delete_rows, delete_cols
from .sort import sort


def patch_all():
    Worksheet.bb_patched = True
    # rename original methods
    Worksheet.insert_rows_ = Worksheet.insert_rows
    Worksheet.insert_cols_ = Worksheet.insert_cols
    Worksheet.delete_rows_ = Worksheet.delete_rows
    Worksheet.delete_cols_ = Worksheet.delete_cols
    # repalce methods
    Worksheet.insert_rows = insert_rows
    Worksheet.insert_cols = insert_cols
    Worksheet.delete_rows = delete_rows
    Worksheet.delete_cols = delete_cols
    Worksheet.sort = sort


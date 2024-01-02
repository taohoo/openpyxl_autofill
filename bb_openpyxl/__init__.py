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


Worksheet.insert_rows_b = insert_rows
Worksheet.insert_cols_b = insert_cols
Worksheet.delete_rows_b = delete_rows
Worksheet.delete_cols_b = delete_cols
Worksheet.sort_b = sort

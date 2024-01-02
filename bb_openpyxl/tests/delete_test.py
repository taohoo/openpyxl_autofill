# -*- coding: utf-8 -*-
"""
@author: hubo
@project: bb-py
@file: insert_test.py
@time: 2023/12/22 14:15
@desc:
"""
from openpyxl import load_workbook
from ..delete import delete_rows, delete_cols


def test_delete_cols():
    wb = load_workbook('bb_openpyxl/tests/test.xlsx')
    ws = wb.worksheets[0]
    delete_cols(ws, 2, amount=2)
    # delete_cols(ws, 5, amount=1)
    wb.save('bb_openpyxl/tests/out.xlsx')


def test_delete_rows():
    wb = load_workbook('bb_openpyxl/tests/test.xlsx')
    ws = wb.worksheets[0]
    # insert_rows(ws, 6)
    delete_rows(ws, 4)
    wb.save('bb_openpyxl/tests/out.xlsx')




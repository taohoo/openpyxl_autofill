# -*- coding: utf-8 -*-
"""
@author: hubo
@project: openpyxl_autofill
@file: test_sort.py.py
@time: 2023/12/22 14:01
@desc:
"""
from openpyxl import load_workbook
from .. import enable_all
enable_all()


def test_sort():
    wb = load_workbook('openpyxl_autofill/tests/test.xlsx')
    ws = wb.worksheets[0]
    # _sort_f(ws, 3, 1, 12, 7, key=lambda x: (x[4].value + x[5].value + x[6].value), reverse=True)
    ws.sort(3, 1, 12, 7, sort_column= 5, reverse=True)
    wb.save('openpyxl_autofill/tests/out.xlsx')

# -*- coding: utf-8 -*-
"""
@author: hubo
@project: bb-py
@file: test_sort.py.py
@time: 2023/12/22 14:01
@desc:
"""
from openpyxl import load_workbook
from ..sort import sort_f


def test_sort():
    wb = load_workbook('bb_openpyxl/tests/test.xlsx')
    ws = wb.worksheets[0]
    sort_f(ws, 3, 1, 12, 7, key=lambda x: (x[4].value + x[5].value + x[6].value), reverse=True)
    wb.save('bb_openpyxl/tests/out.xlsx')

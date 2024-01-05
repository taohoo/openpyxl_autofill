# -*- coding: utf-8 -*-
"""
@author: hubo
@project: bb-py
@file: test_sort.py.py
@time: 2023/12/22 14:01
@desc:
"""
from openpyxl import load_workbook
from .. import patch_all
patch_all()


def test_sort():
    wb = load_workbook('bb_openpyxl/tests/test.xlsx')
    ws = wb.worksheets[0]
    # _sort_f(ws, 3, 1, 12, 7, key=lambda x: (x[4].value + x[5].value + x[6].value), reverse=True)
    ws.sort(3, 1, 12, 7, sort_column= 5, reverse=True)
    wb.save('bb_openpyxl/tests/out.xlsx')

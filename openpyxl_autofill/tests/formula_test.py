# -*- coding: utf-8 -*-
"""
@author: hubo
@project: openpyxl_autofill
@file: formula_test.py
@time: 2023/12/23 14:07
@desc:
"""
from .._formula import _reset_formula
from .._range import _get_all_cells, _get_all_rows, _get_all_columns


def test_get():
    assert ['H8', 'I12'] == _get_all_cells(r"H8:I12")
    assert ['H8', 'I12'] == _get_all_cells(r"=ATAN2(H8,I12)")
    assert ['I6'] == _get_all_cells(r"=RANK(I6,$I$3:$I$12,0)")
    assert ['3', '12'] == _get_all_rows(r"=RANK(I6,$I$3:$I$12,0)")
    assert ['I'] == _get_all_columns(r"=RANK(I6,$I$3:$I$12,0)")


def test_reset_cell():
    assert 'H8:I14' == _reset_formula(r"H8:I12", row_idx=9, amount=2)
    assert 'A8:I14,AA8:II14' == _reset_formula(r"A8:I12,AA8:II12", row_idx=9, amount=2)
    # 增加行，第一个单元格不动
    assert r'=ATAN2(H8,I14)' == _reset_formula(r"=ATAN2(H8,I12)", row_idx=9, amount=2)
    # 增加行，第一个单元格移到第二个单元格的位置
    assert r'=ATAN2(H12,H16)' == _reset_formula(r"=ATAN2(H8,H12)", row_idx=7, amount=4)
    # 增加列，第一个单元格不动
    assert r'=ATAN2(D8,I14)' == _reset_formula(r"=ATAN2(D8,G14)", col_idx=5, amount=2)
    # 增加列，第一个单元格移动第二个单元格的位置
    assert r'=ATAN2(G8,J8)' == _reset_formula(r"=ATAN2(D8,G8)", col_idx=2, amount=3)
    assert r'=ATAN2(J8,G8)' == _reset_formula(r"=ATAN2(G8,D8)", col_idx=2, amount=3)


def test_reset_col():
    # 增加列，只改变列
    assert r"=RANK(A6,$D$3:$L$12,0)" == _reset_formula(r"=RANK(A6,$D$3:$I$12,0)", col_idx=7, amount=3)
    # 增加列，其中一列占了另一列的位置
    assert r"=RANK(A6,$G$3:$J$12,0)" == _reset_formula(r"=RANK(A6,$D$3:$G$12,0)", col_idx=2, amount=3)
    assert r"=RANK(A6,$J$3:$G$12,0)" == _reset_formula(r"=RANK(A6,$G$3:$D$12,0)", col_idx=2, amount=3)


def test_reset_row():
    # 增加行，改变一行
    assert r"=RANK(I6,$I$3:$I$15,0)" == _reset_formula(r"=RANK(I6,$I$3:$I$12,0)", row_idx=7, amount=3)
    # 增加行，第一行替代第二行的位置
    assert r"=RANK(I6,$I$12:$I$15,0)" == _reset_formula(r"=RANK(I6,$I$9:$I$12,0)", row_idx=7, amount=3)
    assert r"=RANK(I6,$I$15:$I$12,0)" == _reset_formula(r"=RANK(I6,$I$12:$I$9,0)", row_idx=7, amount=3)
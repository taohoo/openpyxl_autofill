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
    return insert_rows(worksheet, idx, amount=-amount)


def delete_cols(worksheet, idx, amount=1):
    return insert_cols(worksheet, idx, amount=-amount)

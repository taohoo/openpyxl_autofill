# -*- coding: utf-8 -*-
"""
@author: hubo
@project: bb-py
@file: insert.py
@time: 2023/12/22 14:13
@desc:
关于公式和宏，应该python代码代替复杂的公式和宏。目前支持当个sheet页内的公式迁移，跨sheet页的公式迁移会出错，可以通过公式名称管理一定程度上解决跨sheet页的公式问题
从excel使用角度：
    插入或者删除行或者列的时候，受影响的有公式、合并的单元格、列宽、行高、数据格式、数据验证、图表、宏代码
    其中宏代码不考虑。
    已经处理：公式、合并的单元格、列宽、行高
    待验证：数据格式、数据验证、图表、宏代码
从代码角度：
    已经处理：merged_cells
    未处理完整：row_dimensions，column_dimensions
    待处理：col_breaks, row_breaks, data_validations
    待验证： scenarios
    暂不处理：宏，defined_names，跨sheet页
"""
from ._insert import (_un_merge_cells_before_insert, _re_merge_cells_when_after_insert,
                      _get_all_columns_width, _get_all_rows_height, _reset_all_columns_width, _reset_all_rows_height,
                      _re_set_all_formulas, _warning_unsupported_formula, _re_set_tables)


def insert_rows(worksheet, idx, amount=1):
    """
    Inserts rows into a worksheet at a specified index.
    Auto adjust merged cells, formulas and tables.
    Args:
        worksheet (Worksheet): The worksheet to insert rows into.
        idx (int): The index at which to insert the rows.
        amount (int, optional): The number of rows to insert. Defaults to 1.

    Returns:
        None

    Raises:
        None
    """
    if not hasattr(worksheet, 'bb_patched'):
        # 未打补丁，直接调用的本方法
        worksheet.insert_rows_ = worksheet.insert_rows
        worksheet.delete_rows_ = worksheet.delete_rows
    unmerged_ranges = _un_merge_cells_before_insert(worksheet, row_idx=idx, amount=amount)
    heights = _get_all_rows_height(worksheet)
    if amount > 0:
        worksheet.insert_rows_(idx, amount=amount)
    else:
        worksheet.delete_rows_(idx, amount=-amount)
    _re_set_all_formulas(worksheet, row_idx=idx, amount=amount)
    _warning_unsupported_formula(worksheet.parent)
    _reset_all_rows_height(worksheet, heights, idx, amount=amount)
    _re_merge_cells_when_after_insert(worksheet, unmerged_ranges, row_idx=idx, amount=amount)
    _re_set_tables(worksheet, row_idx=idx, amount=amount)


def insert_cols(worksheet, idx, amount=1):
    """
    Inserts a specified number of columns at a given index in a worksheet.
    Auto adjust merged cells, formulas and tables.
    插入或者删除行或者列的时候，自动校正受影响的有公式、合并的单元格、表格
    Args:
        worksheet (object): The worksheet object on which the columns are inserted.
        idx (int): The index at which the columns are to be inserted.
        amount (int, optional): The number of columns to be inserted. Defaults to 1.

    Returns:
        None
    """
    if not hasattr(worksheet, 'bb_patched'):
        worksheet.insert_cols_ = worksheet.insert_cols
        worksheet.delete_cols_ = worksheet.delete_cols
    unmerged_ranges = _un_merge_cells_before_insert(worksheet, col_idx=idx, amount=amount)
    widths = _get_all_columns_width(worksheet)
    if amount > 0:  # 插入列
        worksheet.insert_cols_(idx, amount=amount)
    else:   # 删除列
        worksheet.delete_cols_(idx, amount=-amount)
    _re_set_all_formulas(worksheet, col_idx=idx, amount=amount)
    _warning_unsupported_formula(worksheet.parent)
    _reset_all_columns_width(worksheet, widths, idx, amount=amount)
    _re_merge_cells_when_after_insert(worksheet, unmerged_ranges, col_idx=idx, amount=amount)
    _re_set_tables(worksheet, col_idx=idx, amount=amount)

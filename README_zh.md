# 从pip安装
```
pip install bb_openpyxl
```
# 使用
```
from openpyxl import load_workbook
import bb_openpyxl
bb_openpyxl.patch_all()
wb = load_workbook(...)
ws = wb.active
ws.insert_rows(...)
...
```
# 针对openpyxl的各种扩展操作
## sort
对给定范围的单元格进行排序
## insert
包含insert_rows和insert_cols，在插入行和列的时候，自动调整原有excel中的公式，合并单元格，表格
## delete
包含delete_rows和delete_cols，在删除行和删除列的时候，自动调整原有excel中的公式，合并单元格，表格
# 下一步计划
增删行和列的时候支持跨sheet的公式，支持宏代码   
排序的时候支持对有公式的列进行调整，先对公式进行计算再排序


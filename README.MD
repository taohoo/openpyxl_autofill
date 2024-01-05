# Installation via pip
```
pip install bb_openpyxl
```
# Usage
```
from openpyxl import load_workbook
import bb_openpyxl
bb_openpyxl.patch_all()
wb = load_workbook(...)
ws = wb.active
ws.insert_rows(...)
...
```
# Extensions for openpyxl
## sort
Sorts the cells within the given range.
## insert
Insert rows and insert cols. Auto adjust merged cells, formulas and tables.
## delete
Delete rows and delete cols. Auto adjust merged cells, formulas and tables.


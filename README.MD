Many people who transition from Excel or VBA to using openpyxl often find themselves confused about the lack of automatic formula correction when inserting or deleting rows or columns in Excel.
This library is designed to automatically correct formulas, merged cells, and table styles when using the insert_rows, insert_cols, delete_rows, and delete_cols methods in openpyxl.
# Installation via pip
```
pip install openpyxl_autofill
```
# Usage
```
from openpyxl import load_workbook
import openpyxl_autofill
openpyxl_autofill.enable_all()
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
# TODO
Support formulas cross sheet,  perspective view.
Support formulas in sort cells.

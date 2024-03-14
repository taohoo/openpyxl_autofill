# Explanation
openpyxl does not provide support for many automation operations in Excel.
For example, when inserting or deleting rows and columns, formulas in Excel are affected.
openpyxl does not have methods like sorting, which are supported in Excel.
This library provides support for these operations.
# Example
## Before deleting columns
![img_2.png](img_2.png)
## Directly delete two columns using openpyxl
Formulas are not automatically adjusted, calculation may be incorrect.
![img.png](img.png)
## Delete two columns using openpyxl with openpyxl_autofill patch 
Formulas are automatically adjusted like Excel.
![img_1.png](img_1.png)
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
# Functions
## sort
Sorts cells in a given range
## insert
Includes insert_rows and insert_cols, which automatically adjust formulas, merged cells, and tables when inserting rows and columns
## delete
Includes delete_rows and delete_cols, which automatically adjust formulas, merged cells, and tables when deleting rows and columns
# Next steps
Support formulas across sheets when inserting or deleting rows and columns
Support adjustment of columns with formulas when sorting, by calculating formulas first before sorting

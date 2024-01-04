# Installation via pip
```
pip install bb_openpyxl
```
# Usage
```
import bb_openpyxl
bb_openpyxl.patch_all()
```
# Various extensions for openpyxl
## sort
Sorts the cells within the given range.
## insert
Contains insert_rows and insert_cols. When inserting rows and columns, it preserves the original formulas and keeps the data consistent for merged cells.
## delete
Contains delete_rows and delete_cols. When deleting rows and columns, it preserves the original formulas and keeps the data consistent for merged cells.


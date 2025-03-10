---
order: 4
title: Spreadsheet File Template - Functions
menu: Functions
toc: false
--- 

This is the list of the currently supported functions in the SpreadSheet template

| function name | Example | parameters | explanation |
|---|---|---|---|
| `ABS` | `{{=ABS(A2)}}` | cell region | returns the absolute value of the SUM of this region. Error, if it finds a value that is not a number.
| `AVERAGE` | `{{=AVERAGE(A2)}}` | cell region | returns the avarage value of this region. Error, if it finds a value that is not a number.
| `CONCAT` | `{{=CONCAT(A2)(S=',')}}` | cell region</br>separator | returns the avarage value of this region.
| `MAX` | `{{=MAX(A2,B2:B3)}}` | cell region | returns the maximum value found in this region. Error, if it finds a value that is not a number.
| `MIN` | `{{=MIN(A2,B2:B3)}}` | cell region | returns the minumum value found in this region. Error, if it finds a value that is not a number.
| `SUM` | `{{=SUM(A2,B2:B3)}}` | cell region | returns a sum of all values found in this region. Error, if it finds a value that is not a number.

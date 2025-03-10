---
order: 3
title: Document File Template - Supported Tag Parameters
menu: Tag Parameters
toc: false
--- 
## Supported in this template in paragraphs
The list of supported tag parameters with WvTextTemplate are as follows:

| template | explanation |
|---|---|---|---|
| `{{column_name}}` | no parameters, will return all values of this column in a concatenated list
| `{{column_name(S=',')}}` | will return all values of this column in a concatenated list separated by ","


## Supported in this template in tables
The list of supported tag parameters with WvSpreadsheetTemplate are as follows:

### Ranges
| template | explanation |
|---|---|---|---|
| `{{=function_name(A1)}}` | single cell reference
| `{{=function_name(A1:C1)}}` | region reference
| `{{=function_name(A1:C1,B2:C2)}}` | multiple regions reference
| `{{=function_name($A1)}}` | static column A reference
| `{{=function_name($A$1)}}` | static row and column

### Flow
| template | explanation |
|---|---|---|---|
| `{{column_name}}` | will expand values vertically or horizontally based on the parent context flow
| `{{column_name(F=V)}}` | will expand values vertically
| `{{column_name(F=D)}}` | will expand values vertically


### ParentContext
The general principal of selecting *default parent context* is:
1. If it expand values vertically the left adjacent cell will be its parent context.
1. If it expands values horizontally the top adjacent cell will be its context.

| template | explanation |
|---|---|---|---|
| `{{column_name}}` | default parent context
| `{{column_name(C=A1)}}` | the first contexts that is found to include *A1* address in its expanded result region will be referred as parent.
| `{{column_name(C=null)}}` | it will not have parent context to influence its expansion direction
| `{{column_name(C=none)}}` | it will not have parent context its expansion direction


### Separator
| template | explanation |
|---|---|---|---|
| `{{column_name(S=',')}}` | values will be rendered in one cell, separated by ","
| `{{column_name(S='$rn')}}` | values will be rendered in one cell, each value on a new line
---
order: 2
title: Document File Template - Cheat sheet
menu: Cheat sheet
toc: false
--- 

## In paragraph text

| template | DataTable | result | explanation |
|---|---|---|---|
| `{{column_name}}` | *column_name* is not present in the data table | `String.Empty` | data is considered missing and template is replaced with an empty result
| `{{column_name}}` | *one row* in the data table with the value *(string)test* | `test` | template is replaced with the value
| `{{column_name}}` | *one row* in the data table with the value *(int)1* | `1` as string | template is replaced with the value cast to string
| `{{column_name}}` | *multiple rows* in the data table with values *(string)test* and *(string)test2* | `testtest2` | template is replaced by joining all values found 
| `{{column_name[0]}}` | *multiple rows* in the data table with values *(string)test* and *(string)test2* | `test` | template is replaced with the first result only as the indexed value of *0* is requested
| `This is {{column_name[0]}} for the test.` | *multiple rows* in the data table with values *(string)test* and *(string)test2* | `This is test for the test.` | the template can be a part of a broader text.
| `{{column_name(S=',')}}` | *multiple rows* in the data table with values *(string)test* and *(string)test2* | `test,test2` | template is replaced by joining all values found with the separator requested by the *S* parameter
| `{{column_name(S=", ")}}` | *multiple rows* in the data table with values *(string)test* and *(string)test2* | `test, test2` | same as above, but shows that intervals are honored.

## In tables

| template | DataTable | result | explanation |
|---|---|---|---|
| `just text` | filled or empty data table | `just text` | template text and styling is migrated 'as is'
| `{{column_name}}` | *one row* in the data table with the value *(string)test* | `test` | the row will not be expanded as there is a single value only. Can be used in worksheet names too.
| `{{column_name}}` | *multiple rows* in the data table with values *(string)test* and *(string)test2* | row 1: `test`; row 2: `test` | the data will expand *vertically* on new rows.
| `{{column_name(F=H)}}` | *multiple rows* in the data table with values *(string)test* and *(string)test2* | row 1: `test`; row 2: `test` | the data will expand *horizontally* on new rows.
| `{{column_name(S=", ")}}` | *multiple rows* in the data table with values *(string)test* and *(string)test2* | `test, test2` | rows will not be expanded as the separator will join the data in a single value.
| `text {{column_name[0]}} text` | *multiple rows* in the data table with values *(string)test* and *(string)test2* | `text test text` | tags can be used mixed with text.
| `{{=sum(A1)}}` | *multiple rows* in the data table with *value* column - *(int)1* and *(int)2*. The *A1* cell is the template *{{value}}* | `3` as number | tag with leading *=* declares a function. More about functions read in the next page.
| `{{==sum(A1)}}` | *multiple rows* in the data table with *value* column - *(int)1* and *(int)2*. The *A1* cell is the template *{{value}}* | `=sum(A1,A2)` as spreadsheet function | tag with leading *==* declares an excel function. More about excel functions read in the next page.
| `{{=sum(A1:B1)}}` | *multiple rows* in the data table | sum of the numeric values in the expanded region | the region reference can be provided as an address like *A1:B1* too
| `{{=sum(A1:B1,C2:C3)}}` | *multiple rows* in the data table | sum of the numeric values in the expanded regions | the referenced regions can be more then one, provided their addresses are separated with *,*
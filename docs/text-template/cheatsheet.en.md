---
order: 2
title: Text Template - Cheat sheet
menu: Cheat sheet
toc: false
--- 

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
| `{{column_name(S="\r\n")}}` | *multiple rows* in the data table with values *(string)test* and *(string)test2* | error | new lines in the parameter value lead to template malformation and not being recognized as template tag
| `{{column_name(S="$rn")}}` | *multiple rows* in the data table with values *(string)test* and *(string)test2* | `test\r\ntest2` | new line can be set as separator by using *$rn* as separator parameter value
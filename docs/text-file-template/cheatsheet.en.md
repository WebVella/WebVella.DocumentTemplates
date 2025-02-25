---
order: 2
title: Cheat sheet - HTML Template
menu: Cheat sheet
toc: false
--- 

| template | DataTable | result | explanation |
|---|---|---|---|
| `just text` | filled or empty data table | byte[] of `just text` | the text content is returned as is
| `just\r\ntext` | filled or empty data table | byte[] of `just\r\ntext` | new lines are migrated
| `{{column_name}}` | *one row* in the data table with the value *(string)test* | byte[] of `test` | new lines are migrated
| `{{column_name}}` | *multiple rows* in the data table with values *(string)test* and *(string)test2* | byte[] of `test\r\ntest2` | each result is generated on a new line
| `{{column_name}} other text` | *multiple rows* in the data table with values *(string)test* and *(string)test2* | byte[] of `test other text\r\ntest2 other text` | each result is generated on a new line with the additional text expended

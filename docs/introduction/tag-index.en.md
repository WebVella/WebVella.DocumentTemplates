---
order: 4
title: Tag Index
menu: Tag Index
toc: false
--- 

When you the dataset will return multiple records but you need to select just a single index of present you can use indexing.

| example | explanation |
|---|---|---|---|
| `{{column_name}}` | no indexing. All dataset rows will be used.
| `{{column_name[0]}}` | will use the value of the first row.
| `{{column_name[1]}}` | will use the value of the second row.
| `{{column_name[0]}} {{column_name[1]}}` | will use the value of the first row for the first tag and the second row for the second tag.
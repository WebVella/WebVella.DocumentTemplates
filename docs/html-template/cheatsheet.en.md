---
order: 2
title: Cheat sheet - HTML Template
menu: Cheat sheet
toc: false
--- 

| template | DataTable | result | explanation |
|---|---|---|---|
| `just text` | filled or empty data table | `<span>just text</span>` | the HTML is returned as is, but wrapped in a span to be proper HTML
| `just<br>text` | filled or empty data table | `<span>just</span><br><span>text</span>` | the HTML is returned as is, but the text parts wrapped in a span to be proper HTML
| `just<br>{{name[0]}}` | *one row* in the data table with the value for the column *name* as *(string)test*  | `<span>just</span><br><span>test</span>` | the HTML is returned as is, but the text parts wrapped in a span to be proper HTML
| `<ul><li>{{name}}</li></ul>` | *two rows* in the data table with the values for the column *name* as *(string)test* and *(string)test2*  | `<ul><li>test</li><li>test2</li></ul>` | when value needs to be expanded the wrapping tag follows its expansion
| `<ul><li>{{name(S=",")}}</li></ul>` | *two rows* in the data table with the values for the column *name* as *(string)test* and *(string)test2*  | `<ul><li>test,test2</li></ul>` | when a separator is defined, the value does not expand but is concatenated
| `<div>{{name}} <span>{{name[0]}}</span></div>` | *two rows* in the data table with the values for the column *name* as *(string)test* and *(string)test2*  | `<div>test <span>test</span></div>test2 <div><span>test</span></div>` | when the parent tag is expanded its node branch is expanded with it
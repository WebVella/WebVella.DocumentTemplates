---
order: 3
title: HTML Template - Supported Tag Parameters
menu: Tag Parameters
toc: false
--- 
## Supported in this template
The list of supported tag parameters with WvTextTemplate are as follows:

| template | explanation |
|---|---|---|---|
| `<span>{{column_name}}</span>` | no parameters, will return all values of this column. Each value will be wrapped in *<span>*
| `<ul><li>{{column_name}}</li></ul>` | no parameters, will return all values of this column. Each value will be wrapped in *<li>*. The outside *<ul>* wrapper will be preserved.
| `<span>{{column_name(S=',')}}</span>` | will return all values of this column in a concatenated list separated by "," and the whole list will be wrapped in one *<span>*
| `<span>{{column_name}}</span>\r\n` | will return multiple values wrapped in a *<span>*, than will finish with a new line

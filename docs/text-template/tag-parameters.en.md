---
order: 3
title: Text Template - Supported Tag Parameters
menu: Tag Parameters
toc: false
--- 
## Supported in this template
The list of supported tag parameters with WvTextTemplate are as follows:

| template | explanation |
|---|---|---|---|
| `{{column_name}}` | no parameters, will return all values of this column in a concatenated list
| `{{column_name(S=',')}}` | will return all values of this column in a concatenated list separated by ","
| `{{column_name(S="$rn")}}` | will return multiple values separated by a new line. The WvTextTemplate is meant to return one line text string, but in certain cases new line is needed. Consider using the WvTextFileTemplate in this case.

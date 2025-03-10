---
order: 3
title: Tag Parameters
menu: Tag Parameters
toc: false
--- 

The tag parameters are provide option to fine tune the result of the template tag. The parameter support may differ from template to template, but there are also a lot of similarities. 
Parameters are grouped in a *parameter group* which is defined by the *()* after the template tag name.

| example | explanation |
|---|---|---|---|
| `{{column_name}}` | this is a tag without parameters
| `{{column_name()}}` | this is a tag has an empty parameter group
| `{{column_name(F=H)}}` | one parameter *F* with value *H* in a single group. It will be sent to the processor
| `{{column_name(F=H,S=',')}}` | two parameters *F* and *S* are defined in a single group
| `{{column_name(F=H)(S=',')}}` | two parameter groups with single parameter each are defined
| `{{column_name(A1:B1)}}` | one parameter *A1:B1* in a single group. Used in spreadsheet processor.
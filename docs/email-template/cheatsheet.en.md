---
order: 2
title: Email Template - Cheat sheet
menu: Cheat sheet
toc: true
--- 
The create a email template you need to provide the *WvEmail* object to the template. Here are one of the most common values for its properties:

## Sender
Required. In an email sender is an email of the person that sends the email. It could be hard-coded or provided from template

| template | DataTable | result | explanation |
|---|---|---|---|
| `test@domain.com` | filled or empty data table | `test@domain.com` | the email will be used as provided
| `{{email[0]}}` | *two rows* in the data table with the values for the column *email* as *(string)test@domain.com* and *(string)test2@domain.com*  | `test@domain.com` | Sender should be only one email, so if the column of the template is not a group by column, better use an *[0]*

## Recipients
Required. A list of emails separated by ```;``` that should receive the email.

| template | DataTable | result | explanation |
|---|---|---|---|
| `{{email[0]}}` | *two rows* in the data table with the values for the column *email* as *(string)test@domain.com* and *(string)test2@domain.com*  | `test@domain.com` | When you need only one email use indexing with *[0]*
| `{{email(S=';')}}` | *two rows* in the data table with the values for the column *email* as *(string)test@domain.com* and *(string)test2@domain.com*  | `test@domain.com;test2@domain.com` | When you need more than one email in a list use the separator parameter *(S=';')*

## CcRecipients
Optional. This property is an email list and should be handles in the same way as the *Recipients*

## BccRecipients
Optional. This property is an email list and should be handles in the same way as the *Recipients*

## Subject
Required. This property will be processed by the [Text template](/en/docs/text-template/cheatsheet) and works the same way. Here is how it applies here:

| template | DataTable | result | explanation |
|---|---|---|---|
| `Email about {{item[0]}}` | *multiple rows* in the data table with values *(string)item1* and *(string)item1* | `Email about item1` | the template can be a part of a broader text. Index *[0]* will use only the first row's value
| `Email about {{item(S=', ')}}` | *multiple rows* in the data table with values *(string)item1* and *(string)item1* | `Email about item1, item2` | use separator to list the values from all rows

## HtmlContent
Optional when TextContent is provided. This property will be processed by the [HTML template](/en/docs/html-template/cheatsheet).

| template | DataTable | result | explanation |
|---|---|---|---|
| `<ul><li>{{name}}</li></ul>` | *two rows* in the data table with the values for the column *name* as *(string)test* and *(string)test2*  | `<ul><li>test</li><li>test2</li></ul>` | when value needs to be expanded the wrapping tag follows its expansion
| `<ul><li>{{name(S=",")}}</li></ul>` | *two rows* in the data table with the values for the column *name* as *(string)test* and *(string)test2*  | `<ul><li>test,test2</li></ul>` | when a separator is defined, the value does not expand but is concatenated
| `<div>{{name}} <span>{{name[0]}}</span></div>` | *two rows* in the data table with the values for the column *name* as *(string)test* and *(string)test2*  | `<div>test <span>test</span></div>test2 <div><span>test</span></div>` | when the parent tag is expanded its node branch is expanded with it

## TextContent
Optional when HtmlContent is provided. This property will be processed by the [Text file template](/en/docs/text-file-template/cheatsheet).

| template | DataTable | result | explanation |
|---|---|---|---|
| `just text` | filled or empty data table | byte[] of `just text` | the text content is returned as is
| `just\r\ntext` | filled or empty data table | byte[] of `just\r\ntext` | new lines are migrated
| `{{column_name}} other text` | *multiple rows* in the data table with values *(string)test* and *(string)test2* | byte[] of `test other text\r\ntest2 other text` | each result is generated on a new line with the additional text expended

## Attachment list
Optional. The attachments are defined by filling the corresponding property of the Email template. Currently only two types are supported: TextFile (processed by the TextFile template) and SpreadsheetFile (processed by the Excel file template). 
When the attachment *GroupDataByColumns* is defined it will work with the subset of each email result (if it has its own data grouping) and will create subsets from it. In this way one attachment can result in multiple files.
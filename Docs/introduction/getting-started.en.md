---
order: 2
title: Getting started
---

## Installation
Add the latest version of WebVella.DocumentTemplates NUGET package to your project
## Template tag
This is the piece of text that the library will look for in order it to be replaced. If result can be generated for the tag it will be replaced with it, if not (eg. wrong column name) the tag will be replaced with an empty string.
The usual format of the tag is as follows:

```csharp
// should match of the column name in the DataTable.
// Case sensitivity depends on the DataTable settings (insensitive by default).
{{column_name}} 
```

```csharp
//square brackets are used for providing row index
{{column_name[0]}}
```
```csharp
//brackets are used to provide parameter and parameter values
{{column_name(S=",")}}
```
```csharp
//multiple parameters can be provided
{{column_name(S=",",F=H)}}
```
```csharp
//functional tags start with '='
{{=sum(A2)}}
```
```csharp
//when excel function is expected to follow tags start with '=='
{{==sum(A2)}}
```
## Template
The template (a.k.a. the template processor) inherits the WvTemplateBase class. It defines the characteristics of the provided template and provides ready result when its 'Process' method is used. 
Currently available templates are: WvExcelFileTemplate, WvEmailTemplate, WvTextFileTemplate, WvTextTemplate, WvHtmlTemplate and the WvWordFileTemplate as a next objective. 
More details can be reviewed in the relevant pages of this site, but in general this is what a template class looks like:

```csharp
public class WvExcelFileTemplate : WvTemplateBase
{
	// template objects reflect the media that is being processed 
	// and they can be different for each template
	public XLWorkbook? Template { get; set; } 
	public WvExcelFileTemplateProcessResult Process(DataTable? dataSource, 
		CultureInfo? culture = null)
	{
		//calling the processor for result generation
	}
}
public abstract class WvTemplateBase
{
	public List<string> GroupDataByColumns { get; set; } = new List<string>();
}
```
## How to use it
The use process is pretty straight forward. You need to follow these steps:

1. Get your data
1. Setup your template object
1. Execute and get the results

## Example without data grouping
Each processing operation have result that can include one or more result items. This usually happens when the data is group in more then one groups by the GroupDataByColumns property. 
In this way as an example, you can generate multiple emails by grouping the data by recipient.

```csharp
using System.Data;
using WebVella.DocumentTemplates.Engines.Text;

namespace WebVella.DocumentTemplates.Examples;
internal class Program
{
	static void Main(string[] args)
	{
		//1. Get the DataTable
		DataTable dt = new DataTable();
		dt.Columns.Add("email", typeof(string));
		var row = dt.NewRow();
		row["email"] = $"john@domain.com";
		dt.Rows.Add(row);
		var row2 = dt.NewRow();
		row2["email"] = $"peter@domain.com";
		dt.Rows.Add(row2);

		//2. Setup the template
		WvTextTemplate template = new()
		{
			GroupDataByColumns = new List<string>(),
			Template = "{{email(S=\",\")}}"
		};

		//3. Execute and get the results
		WvTextTemplateProcessResult result = template.Process(dt);

		Console.WriteLine(result.ResultItems.Count);
		//Output: 1

		Console.WriteLine(result.ResultItems[0]);
		//Output: john@domain.com,peter@domain.com
	}
}
```

## Example with data grouping
Now lets see the same example if we group the data by email

```csharp
using System.Data;
using WebVella.DocumentTemplates.Engines.Text;

namespace WebVella.DocumentTemplates.Examples;
internal class Program
{
	static void Main(string[] args)
	{
		//1. Get the DataTable
		DataTable dt = new DataTable();
		dt.Columns.Add("email", typeof(string));
		var row = dt.NewRow();
		row["email"] = $"john@domain.com";
		dt.Rows.Add(row);
		var row2 = dt.NewRow();
		row2["email"] = $"john@domain.com";
		dt.Rows.Add(row2);
		var row3 = dt.NewRow();
		row3["email"] = $"peter@domain.com";
		dt.Rows.Add(row3);

		//2. Setup the template
		WvTextTemplate template = new()
		{
			//NOTE: we provide grouping column this time
			GroupDataByColumns = new List<string>() { "email" },
			Template = "{{email(S=\",\")}}"
		};

		//3. Execute and get the results
		WvTextTemplateProcessResult result = template.Process(dt);

		Console.WriteLine(result.ResultItems.Count);
		//Output: 2

		Console.WriteLine(result.ResultItems[0]);
		//Output: john@domain.com,john@domain.com

		Console.WriteLine(result.ResultItems[1]);
		//Output: peter@domain.com
	}
}
```
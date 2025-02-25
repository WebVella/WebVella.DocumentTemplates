---
order: 1
title: Introduction to the Text Template
menu: Introduction
toc: false
---
The text template is the simplest and fastest way to use this library. It converts a template text to result texts depending on the data and column grouping provided.

## Example
```csharp
using System.Data;
using WebVella.DocumentTemplates.Engines.Text;

namespace WebVella.DocumentTemplates.Examples;
internal class Program
{
	static void Main(string[] args)
	{
		//Creating the DataTable
		DataTable dt = new DataTable();
		dt.Columns.Add("email", typeof(string));
		//Row 1
		var row = dt.NewRow();
		row["email"] = $"john@domain.com";
		dt.Rows.Add(row);
		//Row 2
		var row2 = dt.NewRow();
		row2["email"] = $"peter@domain.com";
		dt.Rows.Add(row2);

		//Creating the template
		WvTextTemplate template = new()
		{
			GroupDataByColumns = new List<string>(),
			Template = "{{email(S=\",\")}}"
		};

		//Execution
		WvTextTemplateProcessResult result = template.Process(dt);

		Console.WriteLine(result.ResultItems.Count);
		//Output: 1

		Console.WriteLine(result.ResultItems[0]);
		//Output: john@domain.com,peter@domain.com
	}
}
```
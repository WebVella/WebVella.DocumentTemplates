---
order: 1
title: Introduction to the Email Template
menu: Introduction
toc: false
---
The email template is very useful when you need to create a number of emails with or without attachments based on a data set.
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
		row2["email"] = $"john@domain.com";
		dt.Rows.Add(row2);
		//Row 3
		var row3 = dt.NewRow();
		row3["email"] = $"peter@domain.com";
		dt.Rows.Add(row3);

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
		//Output: john@domain.com,john@domain.com,peter@domain.com
	}
}
```
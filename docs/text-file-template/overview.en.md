---
order: 1
title: Introduction to the Text File Template
menu: Introduction
toc: false
---
The main difference between the text template and the text file template is that with the file the template could expand as new rows as part of the same result. 
In this way the same expandable template, in *text file template* will produce *one result* with many rows, as the *text template* it will produce *many results*.

## Example
The Template1.txt file contents are:
```csharp
{{email}}
```
The test code is:
```csharp
using System.Data;
using System.Text;
using WebVella.DocumentTemplates.Engines.TextFile;

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
		WvTextFileTemplate template = new()
		{
			GroupDataByColumns = new List<string>(),
			Template = LoadFile("Template1.txt")
		};

		//Execution
		WvTextFileTemplateProcessResult result = template.Process(dt);

		Console.WriteLine(result.ResultItems.Count);
		//Output: 1
		var resultString = Encoding.UTF8
			.GetString(result.ResultItems[0].Result?.ToArray() ?? new byte[0]);

		Console.WriteLine(resultString);
		//Output: john@domain.com\r\npeter@domain.com
	}

	public static MemoryStream LoadFile(string fileName)
	{
		var path = Path.Combine(Environment.CurrentDirectory, $"Files\\{fileName}");
		var fi = new FileInfo(path);
		if (!fi.Exists) throw new FileNotFoundException();
		var bytes =  File.ReadAllBytes(path);
		return new MemoryStream(bytes);
	}
}
```
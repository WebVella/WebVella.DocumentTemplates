---
order: 1
title: Spreadsheet File Template - Introduction
menu: Introduction
toc: false
---
When you need to convert your data to one or more Spreadsheet files (eg. MS Excel .xlxs), the ```WvExcelFileTemplate``` will be useful for you.

This example is also available in a [GitHub repository](https://github.com/WebVella/WebVella.DocumentTemplates.ExcelExample) so you can review it in details. 

## Template - Spreadsheet file
The goal is to process a template Spreadsheet file "TemplateDoc1.xlsx" with a DataTable. The template file is always processed from left to right and top to bottom. Cell dependencies are taken into consideration. The template content is:

![TemplateDoc1.xlsx contents](/docs/media/spreadsheet-template-1.png)

## Processing the template
There are two helper methods *LoadFileAsStream* and *SaveFileFromStream* which are not displayed here, but if you are interested in their contents, check out the [Github Repository](https://github.com/WebVella/WebVella.DocumentTemplates.ExcelExample/blob/main/Utils.cs).

```csharp
using System.Data;
using WebVella.DocumentTemplates.Engines.SpreadsheetFile;

namespace WebVella.DocumentTemplates.ExcelExample;
internal class Program
{
	static void Main(string[] args)
	{
		var templateFile = "TemplateDoc1.xlsx";
		var ds = new DataTable();
		ds.Columns.Add("position", typeof(int));
		ds.Columns.Add("sku", typeof(string));
		ds.Columns.Add("item", typeof(string));
		ds.Columns.Add("price", typeof(decimal));

		for (int i = 1; i < 6; i++)
		{
			var dsrow = ds.NewRow();
			dsrow["position"] = i;
			dsrow["sku"] = $"SKU{i}";
			dsrow["item"] = $"item {i} description text";
			dsrow["price"] = i * (decimal)0.98;
			ds.Rows.Add(dsrow);
		}

		var template = new WvSpreadsheetFileTemplate
		{
			Template = new Utils().LoadFileAsStream(templateFile)
		};
		WvSpreadsheetFileTemplateProcessResult? result = template.Process(ds);

		new Utils().SaveFileFromStream(result.ResultItems[0].Result!, $"result-{templateFile}");
	}
}

```

## The result

![TemplateDoc1.xlsx result](/docs/media/spreadsheet-template-1-result.png)

Take notice to:

- image, image positioning, cell and text styling are migrated from the template
- cell merge is migrated
- the data is expanded vertically (by default) and new additional rows were created
- table borders and styling is migrated from the template
- the function SUM is calculated and set with its value. (You have the option of using spreadsheet functions as well which will be kept as excel functions during migration, but their regions will be expanded with the data)
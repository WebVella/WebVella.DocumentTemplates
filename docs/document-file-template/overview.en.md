---
order: 1
title: Document File Template - Introduction
menu: Introduction
toc: false
---
For converting data to document files such as MS Word .docx, you can use the ```WvDocumentFileTemplate``` processor. 

This example is also available in a [GitHub repository](https://github.com/WebVella/WebVella.DocumentTemplates.WordExample) so you can review it in details. 

## Template - Document file
The goal is to process a template MS Word file "TemplateDoc1.docx" with a DataTable. The template file is always processed from top to bottom.

![TemplateDoc1.docx original](/docs/static/document-template-original.png)

## Processing the template
There are two helper methods *LoadFileAsStream* and *SaveFileFromStream* which are not displayed here, but if you are interested in their contents, check out the [Github Repository](https://github.com/WebVella/WebVella.DocumentTemplates.WordExample/blob/main/Utils.cs).

The test code is:
```csharp
using System.Data;
using WebVella.DocumentTemplates.Engines.DocumentFile;

namespace WebVella.DocumentTemplates.WordExample;

internal class Program
{
	static void Main(string[] args)
	{
		var templateFile = "TemplateDoc1.docx";
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

		var template = new WvDocumentFileTemplate
		{
			Template = new Utils().LoadFileAsStream(templateFile)
		};
		WvDocumentFileTemplateProcessResult? result = template.Process(ds);

		new Utils().SaveFileFromStream(result.ResultItems[0].Result!, $"result-{templateFile}");
	}
}
```

## The result

![TemplateDoc1.docx result](/docs/static/document-template-result.png)

Take notice to:

- image, image positioning and text styling are migrated from the template
- headers and footers are processed also. Images are migrated properly.
- the table is migrated and expanded vertically (by default) and new additional rows were created
- table borders and styling is migrated from the template
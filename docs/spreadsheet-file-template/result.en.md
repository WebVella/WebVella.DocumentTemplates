---
order: 6
title: Spreadsheet File Template - Result and Handling errors
menu: Result & Errors
toc: false
---

## Process Result
As a result of the generation process, you will receive a specif to the template type response, but with properties that are inherited by all template results. For this template they are:

```csharp
//Result object
/////////////////////////
public abstract class WvTemplateProcessResultBase
{
	public List<string> GroupByDataColumns { get; set; } = new List<string>();
	public virtual object? Template { get; set; }
	public virtual List<WvTemplateProcessResultItemBase> ResultItems { get; set; } = new();
}

public class WvSpreadsheetFileTemplateProcessResult : WvTemplateProcessResultBase
{
	public XLWorkbook? Workbook { get; set; } = null;
	public new MemoryStream? Template { get; set; } = null;
	public List<WvSpreadsheetFileTemplateRow> TemplateRows { get; set; } = new();
	public List<WvSpreadsheetFileTemplateContext> TemplateContexts { get; set; } = new();
	public new List<WvSpreadsheetFileTemplateProcessResultItem> ResultItems { get; set; } = new();
}

//Result item
/////////////////////////
public abstract class WvTemplateProcessResultItemBase
{
	public virtual object? Result { get; set; }
	public virtual List<WvTemplateProcessContextBase> Contexts { get; set; } = new();
	public bool HasError => Contexts.Any(x => x.HasError);
	public long NumberOfDataTableRows { get; set; } = 0;
}

public class WvSpreadsheetFileTemplateProcessResultItem : WvTemplateProcessResultItemBase
{
	public new XLWorkbook? Result { get; set; } = new();
	public new List<WvSpreadsheetFileTemplateProcessContext> Contexts { get; set; } = new();
	public List<WvSpreadsheetFileTemplateProcessResultItemRow> ResultRows { get; set; } = new();
	public Dictionary<Guid, int> ContextProcessLog { get; set; } = new();
}

//Process context
/////////////////////////
public abstract class WvTemplateProcessContextBase
{
	public bool HasError => Errors is not null && Errors.Count > 0;
	public List<string> Errors { get; set;} = new();
}

public class WvSpreadsheetFileTemplateProcessResultItemContext
{
	public Guid TemplateContextId { get; set; }
	public IXLRange? Range { get; set; }
	public int ExpandCount { get; set; } = 1;
	public WvSpreadsheetFileTemplateProcessResultItemContextError? Error { get; set; } = null;
	public string? ErrorMessage { get; set; } = null;

}

```

## Handling Errors
During the processing, errors can occur. They will not be thrown but reported in the result process contexts by their *Errors* property.There you can find if any error occured during the template generation process.
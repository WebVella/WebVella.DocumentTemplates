---
order: 3
title: Result and Handling errors
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

public class WvTextFileTemplateProcessResult : WvTemplateProcessResultBase
{
	public new MemoryStream? Template { get; set; } = null;
	public new List<WvTextFileTemplateProcessResultItem> ResultItems { get; set; } = new();
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

public class WvTextFileTemplateProcessResultItem : WvTemplateProcessResultItemBase
{
	public new MemoryStream? Result { get; set; } = null;
	public new List<WvTextFileTemplateProcessContext> Contexts { get; set; } = new();
}

//Process context
/////////////////////////
public abstract class WvTemplateProcessContextBase
{
	public bool HasError => Errors is not null && Errors.Count > 0;
	public List<string> Errors { get; set;} = new();
}

public class WvTextFileTemplateProcessContext : WvTemplateProcessContextBase
{
}

```

## Handling Errors
During the processing, errors can occur. They will not be thrown but reported in the result process contexts by their *Errors* property.There you can find if any error occured during the template generation process.
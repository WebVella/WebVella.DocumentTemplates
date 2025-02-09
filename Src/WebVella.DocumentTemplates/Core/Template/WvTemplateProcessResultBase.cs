namespace WebVella.DocumentTemplates.Core;
public abstract class WvTemplateProcessResultBase
{
	public List<string> GroupByDataColumns { get; set; } = new List<string>();
	public virtual object? Template { get; set; }
	public virtual List<WvTemplateProcessResultItemBase> ResultItems { get; set; } = new();
}

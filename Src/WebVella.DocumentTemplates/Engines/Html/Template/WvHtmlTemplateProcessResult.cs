using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Html;
public class WvHtmlTemplateProcessResult : WvTemplateProcessResultBase
{
	public new string? Template { get; set; } = null;
	public new List<WvHtmlTemplateProcessResultItem> ResultItems { get; set; } = new();
}
using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Html;
public class WvHtmlTemplateProcessResultItem : WvTemplateProcessResultItemBase
{
	public new string? Result { get; set; } = null;
	public new List<WvTemplateProcessContextBase> Contexts { get; set; } = new();
}
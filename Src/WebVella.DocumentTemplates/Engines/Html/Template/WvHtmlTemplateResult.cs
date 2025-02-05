using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Html;
public class WvHtmlTemplateResult : WvTemplateResult
{
	public new string? Template { get; set; } = null;
	public new string? Result { get; set; } = null;
	public new IEnumerable<WvTemplateContext> Contexts { get; set; } = Enumerable.Empty<WvTemplateContext>();
}
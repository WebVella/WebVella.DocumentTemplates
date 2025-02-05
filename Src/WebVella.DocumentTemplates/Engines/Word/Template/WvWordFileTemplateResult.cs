using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Text;
public class WvWordFileTemplateResult : WvTemplateResult
{
	public new string? Template { get; set; } = null;
	public new string? Result { get; set; } = null;
	public new IEnumerable<WvTemplateContext> Contexts { get; set; } = Enumerable.Empty<WvTemplateContext>();
}
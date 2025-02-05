using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Text;
public class WvTextFileTemplateResult : WvTemplateResult
{
	public new byte[]? Template { get; set; } = null;
	public new byte[]? Result { get; set; } = null;
	public new IEnumerable<WvTemplateContext> Contexts { get; set; } = Enumerable.Empty<WvTemplateContext>();
}
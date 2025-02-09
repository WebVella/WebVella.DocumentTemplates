using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Text;
public class WvTextTemplateProcessResultItem : WvTemplateProcessResultItemBase
{
	public new string? Result { get; set; } = null;
	public new List<WvTextTemplateProcessContext> Contexts { get; set; } = new();
}
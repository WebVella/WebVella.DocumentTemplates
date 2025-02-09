using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Text;
public class WvEmailTemplateProcessResultItem : WvTemplateProcessResultItemBase
{
	public new string? Result { get; set; } = null;
	public new List<WvEmailTemplateProcessContext> Contexts { get; set; } = new();
}
using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.TextFile;
public class WvTextFileTemplateProcessResultItem : WvTemplateProcessResultItemBase
{
	public new MemoryStream? Result { get; set; } = null;
	public List<WvTextFileTemplateProcessContext> Contexts { get; set; } = new();
}
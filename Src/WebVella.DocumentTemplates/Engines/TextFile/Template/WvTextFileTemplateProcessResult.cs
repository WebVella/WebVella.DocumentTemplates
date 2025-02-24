using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.TextFile;
public class WvTextFileTemplateProcessResult : WvTemplateProcessResultBase
{
	public new MemoryStream? Template { get; set; } = null;
	public new List<WvTextFileTemplateProcessResultItem> ResultItems { get; set; } = new();
}
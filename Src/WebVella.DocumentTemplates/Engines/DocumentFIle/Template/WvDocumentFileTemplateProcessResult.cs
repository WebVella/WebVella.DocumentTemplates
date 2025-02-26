using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Text;
public class WvDocumentFileTemplateProcessResult : WvTemplateProcessResultBase
{
	public new string? Template { get; set; } = null;
	public new List<WvDocumentFileTemplateProcessResultItem> ResultItems { get; set; } = new();
}
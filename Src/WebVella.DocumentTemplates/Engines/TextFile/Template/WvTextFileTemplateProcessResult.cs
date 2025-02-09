using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Text;
public class WvTextFileTemplateProcessResult : WvTemplateProcessResultBase
{
	public new byte[]? Template { get; set; } = null;
	public new List<WvTextFileTemplateProcessResultItem> ResultItems { get; set; } = new();
}
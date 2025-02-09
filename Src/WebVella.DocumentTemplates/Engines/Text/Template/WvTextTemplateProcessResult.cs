using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Text;
public class WvTextTemplateProcessResult : WvTemplateProcessResultBase
{
	public new string? Template { get; set; } = null;
	public new List<WvTextTemplateProcessResultItem> ResultItems { get; set; } = new();
}
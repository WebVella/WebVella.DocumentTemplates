using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Text;
public class WvWordFileTemplateProcessResult : WvTemplateProcessResultBase
{
	public new string? Template { get; set; } = null;
	public new List<WvWordFileTemplateProcessResultItem> ResultItems { get; set; } = new();
}
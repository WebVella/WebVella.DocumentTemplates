using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Text;
public class WvEmailTemplateProcessResult : WvTemplateProcessResultBase
{
	public new string? Template { get; set; } = null;
	public new List<WvEmailTemplateProcessResultItem> ResultItems { get; set; } = new();
}
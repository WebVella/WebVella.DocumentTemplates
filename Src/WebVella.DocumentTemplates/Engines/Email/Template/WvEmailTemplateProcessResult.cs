using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Engines.Email.Models;
namespace WebVella.DocumentTemplates.Engines.Email;
public class WvEmailTemplateProcessResult : WvTemplateProcessResultBase
{
	public new WvEmail? Template { get; set; } = null;
	public new List<WvEmailTemplateProcessResultItem> ResultItems { get; set; } = new();
}
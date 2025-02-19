using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Engines.Email.Models;
namespace WebVella.DocumentTemplates.Engines.Email;
public class WvEmailTemplateProcessResultItem : WvTemplateProcessResultItemBase
{
	public new WvEmail? Result { get; set; } = null;
}
using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Text;
public class WvTextFileTemplateProcessResultItem : WvTemplateProcessResultItemBase
{
	public new byte[]? Result { get; set; } = null;
}
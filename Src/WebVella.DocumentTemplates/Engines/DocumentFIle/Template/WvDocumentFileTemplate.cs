using System.Data;
using System.Globalization;
using System.Text;
using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Text;
public class WvDocumentFileTemplate : WvTemplateBase
{
	public string? Template { get; set; }
	public WvDocumentFileTemplateProcessResult Process(DataTable dataSource, CultureInfo? culture = null)
	{
		return new WvDocumentFileTemplateProcessResult() { Template = Template };
	}
}
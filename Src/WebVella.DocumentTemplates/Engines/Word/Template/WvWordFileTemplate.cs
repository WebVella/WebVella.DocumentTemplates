using System.Data;
using System.Globalization;
using System.Text;
using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Text;
public class WvWordFileTemplate : WvTemplateBase
{
	public string? Template { get; set; }
	public WvWordFileTemplateProcessResult Process(DataTable dataSource, CultureInfo? culture = null)
	{
		return new WvWordFileTemplateProcessResult() { Template = Template };
	}
}
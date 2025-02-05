using ClosedXML.Excel;
using System.Data;
using System.Globalization;
using System.Text;
using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Excel;
public class WvExcelFileTemplate : WvTemplateBase
{
	public XLWorkbook? Template { get; set; }
	public WvExcelFileTemplateResult? Process(DataTable dt, CultureInfo? culture = null)
	{
		return null;
	}
}
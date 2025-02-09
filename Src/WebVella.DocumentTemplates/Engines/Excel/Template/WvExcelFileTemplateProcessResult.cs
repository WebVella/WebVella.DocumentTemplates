using ClosedXML.Excel;
using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Excel;
public class WvExcelFileTemplateProcessResult : WvTemplateProcessResultBase
{
	public new XLWorkbook? Template { get; set; } = null;
	public new List<WvExcelFileTemplateProcessResultItem> ResultItems { get; set; } = new();
}
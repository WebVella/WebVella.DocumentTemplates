using ClosedXML.Excel;
using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.ExcelFile;
public class WvExcelFileTemplateProcessResult : WvTemplateProcessResultBase
{
	public XLWorkbook? Workbook { get; set; } = null;
	public new MemoryStream? Template { get; set; } = null;
	public List<WvExcelFileTemplateRow> TemplateRows { get; set; } = new();
	public List<WvExcelFileTemplateContext> TemplateContexts { get; set; } = new();
	public new List<WvExcelFileTemplateProcessResultItem> ResultItems { get; set; } = new();
}
using ClosedXML.Excel;
using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.SpreadsheetFile;
public class WvSpreadsheetFileTemplateProcessResult : WvTemplateProcessResultBase
{
	public XLWorkbook? Workbook { get; set; } = null;
	public new MemoryStream? Template { get; set; } = null;
	public List<WvSpreadsheetFileTemplateRow> TemplateRows { get; set; } = new();
	public List<WvSpreadsheetFileTemplateContext> TemplateContexts { get; set; } = new();
	public new List<WvSpreadsheetFileTemplateProcessResultItem> ResultItems { get; set; } = new();
}
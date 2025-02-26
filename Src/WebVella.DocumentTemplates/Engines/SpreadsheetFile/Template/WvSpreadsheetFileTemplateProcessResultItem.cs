using ClosedXML.Excel;
using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.SpreadsheetFile;
public class WvSpreadsheetFileTemplateProcessResultItem : WvTemplateProcessResultItemBase
{
	public new MemoryStream? Result { get; set; } = new();
	public XLWorkbook? Workbook { get; set; } = new();
	public new List<WvSpreadsheetFileTemplateProcessContext> Contexts { get; set; } = new();
	public List<WvSpreadsheetFileTemplateProcessResultItemRow> ResultRows { get; set; } = new();
	public Dictionary<Guid, int> ContextProcessLog { get; set; } = new();
}


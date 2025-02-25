using ClosedXML.Excel;
using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.ExcelFile;
public class WvExcelFileTemplateProcessResultItem : WvTemplateProcessResultItemBase
{
	public new MemoryStream? Result { get; set; } = new();
	public XLWorkbook? Workbook { get; set; } = new();
	public new List<WvExcelFileTemplateProcessContext> Contexts { get; set; } = new();
	public List<WvExcelFileTemplateProcessResultItemRow> ResultRows { get; set; } = new();
	public Dictionary<Guid, int> ContextProcessLog { get; set; } = new();
}


using ClosedXML.Excel;
using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.ExcelFile;
public class WvExcelFileTemplateProcessResultItem : WvTemplateProcessResultItemBase
{
	public new XLWorkbook? Result { get; set; } = new();
	public List<WvExcelFileTemplateProcessContext> Contexts { get; set; } = new();
	public List<WvExcelFileTemplateProcessResultItemRow> ResultRows { get; set; } = new();
	public Dictionary<Guid, int> ContextProcessLog { get; set; } = new();
}


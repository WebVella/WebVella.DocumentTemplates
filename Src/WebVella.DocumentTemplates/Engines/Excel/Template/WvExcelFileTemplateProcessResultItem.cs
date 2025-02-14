using ClosedXML.Excel;
using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Excel;
public class WvExcelFileTemplateProcessResultItem : WvTemplateProcessResultItemBase
{
	public new XLWorkbook? Result { get; set; } = new();
	public List<WvExcelFileTemplateProcessResultItemContext> ResultContexts { get; set; } = new();
	public List<WvExcelFileTemplateProcessResultItemRow> ResultRows { get; set; } = new();
	public Dictionary<Guid, int> ContextProcessLog { get; set; } = new();
}


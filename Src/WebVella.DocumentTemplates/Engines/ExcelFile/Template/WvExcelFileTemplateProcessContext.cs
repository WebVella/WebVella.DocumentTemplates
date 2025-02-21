using ClosedXML.Excel;
using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.ExcelFile;
public class WvExcelFileTemplateProcessContext : WvTemplateProcessContextBase
{
	public Guid TemplateContextId { get; set; }
	public IXLRange? Range { get; set; }
	public int ExpandCount { get; set; } = 1;
	public WvExcelFileTemplateProcessResultItemContextError? ExcelCellError { get; set; } = null;
	public string? ExcelCellErrorMessage { get; set; } = null;

}


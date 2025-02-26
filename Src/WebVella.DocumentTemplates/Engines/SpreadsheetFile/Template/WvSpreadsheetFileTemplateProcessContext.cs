using ClosedXML.Excel;
using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.SpreadsheetFile;
public class WvSpreadsheetFileTemplateProcessContext : WvTemplateProcessContextBase
{
	public Guid TemplateContextId { get; set; }
	public IXLRange? Range { get; set; }
	public int ExpandCount { get; set; } = 1;
	public WvSpreadsheetFileTemplateProcessResultItemContextError? SpreadsheetCellError { get; set; } = null;
	public string? SpreadsheetCellErrorMessage { get; set; } = null;

}


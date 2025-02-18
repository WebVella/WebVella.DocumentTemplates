using ClosedXML.Excel;
namespace WebVella.DocumentTemplates.Engines.Excel;
public class WvExcelFileTemplateProcessResultItemContext
{
	public Guid TemplateContextId { get; set; }
	public IXLRange? Range { get; set; }
	public int ExpandCount { get; set; } = 1;
	public WvExcelFileTemplateProcessResultItemContextError? Error { get; set; } = null;
	public string? ErrorMessage { get; set; } = null;

}


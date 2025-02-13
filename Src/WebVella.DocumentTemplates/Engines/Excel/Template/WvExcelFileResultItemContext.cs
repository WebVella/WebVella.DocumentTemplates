using ClosedXML.Excel;
using DocumentFormat.OpenXml.Validation;
namespace WebVella.DocumentTemplates.Engines.Excel;
public class WvExcelFileResultItemContext
{
	public Guid TemplateContextId { get; set; }
	public IXLRange? Range { get; set; }
	public WvExcelFileResultItemContextError? Error { get; set; } = null;
	public string? ErrorMessage { get; set; } = null;

}

public enum WvExcelFileResultItemContextError
{
	DependencyOverflow = 0,
	ProcessError = 1,
}
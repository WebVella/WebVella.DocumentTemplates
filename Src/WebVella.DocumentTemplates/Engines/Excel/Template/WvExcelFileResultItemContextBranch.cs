using ClosedXML.Excel;
using DocumentFormat.OpenXml.Validation;
namespace WebVella.DocumentTemplates.Engines.Excel;
public class WvExcelFileResultItemContextBranch
{
	public Guid TemplateContextBranchId { get; set; }
	public IXLRange? Range { get; set; }
	public List<WvExcelFileResultItemContext> ResultItemContexts { get; set; } = new();

}

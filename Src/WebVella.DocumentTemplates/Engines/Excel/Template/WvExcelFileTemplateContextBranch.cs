using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Excel;
public class WvExcelFileTemplateContextBranch
{
	public Guid Id { get; set; }
	public IXLWorksheet? Worksheet { get; set; }
	public IXLRange? Range { get; set; }
	public List<WvExcelFileTemplateContext> TemplateContexts { get; set; } = new();
}

using ClosedXML.Excel;
using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Excel;
public class WvExcelFileTemplateProcessContext : WvTemplateProcessContextBase
{
	public IXLWorksheet? TemplateWorksheet { get; set; }
	public IXLWorksheet? ResultWorksheet { get; set; }
	public IXLRange? TemplateRange { get; set; }
	public IXLRange? ResultRange { get; set; }
	public List<WvExcelFileTemplateContextRangeAddress> ResultRangeSlots { get; set; } = new();
}
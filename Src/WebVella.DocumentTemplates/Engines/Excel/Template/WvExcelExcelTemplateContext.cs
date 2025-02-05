using ClosedXML.Excel;
using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Excel;
public class WvExcelExcelTemplateContext : WvTemplateContext
{
	public IXLWorksheet? TemplateWorksheet { get; set; }
	public IXLWorksheet? ResultWorksheet { get; set; }
	public IXLRange? TemplateRange { get; set; }
	public IXLRange? ResultRange { get; set; }
	public IEnumerable<WvExcelExcelTemplateContextRangeAddress> ResultRangeSlots { get; set; } = Enumerable.Empty<WvExcelExcelTemplateContextRangeAddress>();
}
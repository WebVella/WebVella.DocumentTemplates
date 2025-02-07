using ClosedXML.Excel;
using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Excel;
public class WvExcelFileTemplateResult : WvTemplateResult
{
	public new XLWorkbook? Template { get; set; } = null;
	public new XLWorkbook? Result { get; set; } = new();
	public new List<WvExcelFileTemplateContext> Contexts { get; set; } = new List<WvExcelFileTemplateContext>();
}
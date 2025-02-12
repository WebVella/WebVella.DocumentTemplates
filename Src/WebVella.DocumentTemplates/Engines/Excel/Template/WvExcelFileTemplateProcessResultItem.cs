using ClosedXML.Excel;
using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Excel;
public class WvExcelFileTemplateProcessResultItem : WvTemplateProcessResultItemBase
{
	public new XLWorkbook? Result { get; set; } = new();
	public List<WvExcelFileResultItemContext> ResultContexts { get; set; } = new();

	public new List<WvExcelFileTemplateProcessContext> Contexts { get; set; } = new List<WvExcelFileTemplateProcessContext>();
}
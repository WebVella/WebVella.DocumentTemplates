using ClosedXML.Excel;
using System.Data;
using WebVella.DocumentTemplates.Core;

namespace WebVella.DocumentTemplates.Engines.Excel.Models;
public interface IWvExcelFileTemplateExcelFunctionProcessor
{
	public string Name { get; }
	public int Priority { get; }
	public string? FormulaA1 { get; set; }
	public bool HasError { get; set; }
	public string? ErrorMessage { get; set; }
	public void Process(
		WvTemplateTag tag,
		DataTable dataSource,
		WvExcelFileTemplateProcessResult result,
		WvExcelFileTemplateProcessResultItem resultItem,
		IXLWorksheet worksheet		
	);

}

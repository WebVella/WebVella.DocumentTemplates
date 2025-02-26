using ClosedXML.Excel;
using WebVella.DocumentTemplates.Core;

namespace WebVella.DocumentTemplates.Engines.SpreadsheetFile.Models;
public interface IWvSpreadsheetFileTemplateSpreadsheetFunctionProcessor
{
	public string Name { get; }
	public int Priority { get; }
	public string? FormulaA1 { get; set; }
	public bool HasError { get; set; }
	public string? ErrorMessage { get; set; }
	public object? Process(
		object? value,
		WvTemplateTag tag,
		WvSpreadsheetFileTemplateContext templateContext,
		int expandPosition,
		int expandPositionMax,
		WvSpreadsheetFileTemplateProcessResult result,
		WvSpreadsheetFileTemplateProcessResultItem resultItem,
		IXLWorksheet processedWorksheet
	);

}

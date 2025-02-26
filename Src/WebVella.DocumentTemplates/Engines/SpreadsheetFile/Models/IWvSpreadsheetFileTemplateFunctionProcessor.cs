using ClosedXML.Excel;
using System.Data;
using WebVella.DocumentTemplates.Core;

namespace WebVella.DocumentTemplates.Engines.SpreadsheetFile.Models;
public interface IWvSpreadsheetFileTemplateFunctionProcessor
{
	public string Name { get; }
	public int Priority { get; }
	public bool HasError { get; set; }
	public string? ErrorMessage { get; set; }
	public object? Process(
		string? value,
		WvTemplateTag tag,
		WvSpreadsheetFileTemplateContext templateContext,
		int expandPosition,
		int expandPositionMax,
		DataTable dataSource,
		WvSpreadsheetFileTemplateProcessResult result,
		WvSpreadsheetFileTemplateProcessResultItem resultItem,
		IXLRange processedCellRange,
		IXLWorksheet processedWorksheet		
	);

}

using ClosedXML.Excel;
using System.Data;
using WebVella.DocumentTemplates.Core;

namespace WebVella.DocumentTemplates.Engines.Excel.Models;
public interface IWvExcelFileTemplateFunctionProcessor
{
	public string Name { get; }
	public int Priority { get; }
	public bool HasError { get; set; }
	public string? ErrorMessage { get; set; }
	public object? Process(
		string? tagValue,
		WvTemplateTag tag,
		WvExcelFileTemplateContext templateContext,
		int expandPosition,
		int expandPositionMax,
		DataTable dataSource,
		WvExcelFileTemplateProcessResult result,
		WvExcelFileTemplateProcessResultItem resultItem,
		IXLRange processedCellRange,
		IXLWorksheet processedWorksheet		
	);

}

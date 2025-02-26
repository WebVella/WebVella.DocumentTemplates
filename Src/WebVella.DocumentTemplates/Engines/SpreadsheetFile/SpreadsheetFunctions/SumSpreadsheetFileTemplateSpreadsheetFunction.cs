using ClosedXML.Excel;
using System.Data;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Engines.SpreadsheetFile.Models;
using WebVella.DocumentTemplates.Engines.SpreadsheetFile.Utility;

namespace WebVella.DocumentTemplates.Engines.SpreadsheetFile.SpreadsheetFunctions;
public class SumSpreadsheetFileTemplateSpreadsheetFunction : IWvSpreadsheetFileTemplateSpreadsheetFunctionProcessor
{
	public string Name { get; } = "sum";
	public int Priority { get; } = 10000;
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
			IXLWorksheet worksheet
		)
	{
		if (tag is null) throw new ArgumentException(nameof(tag));
		if (result is null) throw new ArgumentException(nameof(result));
		if (resultItem is null) throw new ArgumentException(nameof(resultItem));
		if (worksheet is null) throw new ArgumentException(nameof(worksheet));

		if (tag.Type != WvTemplateTagType.SpreadsheetFunction) throw new ArgumentException("Template tag is not SpreadsheetFunction type", nameof(tag));

		var rangeList = new WvSpreadsheetRangeHelpers().GetRangeAddressesForTag(
			tag: tag,
			templateContext: templateContext,
			expandPosition: expandPosition,
			expandPositionMax: expandPositionMax,
			result: result,
			resultItem: resultItem,
			worksheet: worksheet
		);
		if (rangeList.Count > 0)
			FormulaA1 = $"=SUM({String.Join(",", rangeList)})";
		else
		{
			HasError = true;
			ErrorMessage = "No ranges can be determined from template";
		}

		return value;
	}
}

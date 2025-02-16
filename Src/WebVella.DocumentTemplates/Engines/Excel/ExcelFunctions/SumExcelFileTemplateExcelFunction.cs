using ClosedXML.Excel;
using System.Data;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Engines.Excel.Models;
using WebVella.DocumentTemplates.Engines.Excel.Utility;

namespace WebVella.DocumentTemplates.Engines.Excel.ExcelFunctions;
public class SumExcelFileTemplateExcelFunction : IWvExcelFileTemplateExcelFunctionProcessor
{
	public string Name { get; } = "sum";
	public int Priority { get; } = 10000;
	public string? FormulaA1 { get; set; }
	public bool HasError { get; set; }
	public string? ErrorMessage { get; set; }

	public void Process(
			WvTemplateTag tag,
			DataTable dataSource,
			WvExcelFileTemplateProcessResult result,
			WvExcelFileTemplateProcessResultItem resultItem,
			IXLWorksheet worksheet
		)
	{
		if (tag is null) throw new ArgumentException(nameof(tag));
		if (dataSource is null) throw new ArgumentException(nameof(dataSource));
		if (result is null) throw new ArgumentException(nameof(result));
		if (resultItem is null) throw new ArgumentException(nameof(resultItem));
		if (worksheet is null) throw new ArgumentException(nameof(worksheet));

		if (tag.Type != WvTemplateTagType.ExcelFunction) throw new ArgumentException("Template tag is not ExcelFunction type", nameof(tag));


		if (tag.ParamGroups.Count > 0
			&& tag.ParamGroups[0].Parameters.Count > 0
			&& !String.IsNullOrWhiteSpace(tag.FullString))
		{
			var rangeList = new List<string>();
			foreach (var param in tag.ParamGroups[0].Parameters)
			{
				if (String.IsNullOrWhiteSpace(param.ValueString)) continue;
				var range = WvExcelRangeHelpers.GetRangeFromString(param.ValueString ?? String.Empty);
				if (range is not null)
				{
					var rangeTemplateContexts = result.TemplateContexts.GetIntersections(
						worksheetPosition: worksheet.Position,
						range: range,
						type: WvExcelFileTemplateContextType.CellRange
					);
					var resultContexts = resultItem.ResultContexts.Where(x => rangeTemplateContexts.Any(y => y.Id == x.TemplateContextId));
					foreach (var resultCtx in resultContexts)
					{
						if (resultCtx.Range?.RangeAddress is null) continue;
						rangeList.Add(resultCtx.Range.RangeAddress.ToString() ?? String.Empty);
					}
				}
			}
			if (rangeList.Count > 0)
				FormulaA1 = $"=SUM({String.Join(",", rangeList)})";
		}
	}
}

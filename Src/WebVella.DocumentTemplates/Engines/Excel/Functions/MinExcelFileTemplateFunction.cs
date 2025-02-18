using ClosedXML.Excel;
using System.Data;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Engines.Excel.Models;
using WebVella.DocumentTemplates.Engines.Excel.Utility;

namespace WebVella.DocumentTemplates.Engines.Excel.Functions;
public class MinExcelFileTemplateFunction : IWvExcelFileTemplateFunctionProcessor
{
	public string Name { get; } = "min";
	public int Priority { get; } = 10000;
	public bool HasError { get; set; }
	public string? ErrorMessage { get; set; }
	public object? Process(
			string? tagValue,
			WvTemplateTag tag,
			DataTable dataSource,
			WvExcelFileTemplateProcessResult result,
			WvExcelFileTemplateProcessResultItem resultItem,
			IXLRange resultRange,
			IXLWorksheet worksheet
		)
	{
		if (tag is null) throw new ArgumentException(nameof(tag));
		if (dataSource is null) throw new ArgumentException(nameof(dataSource));
		if (result is null) throw new ArgumentException(nameof(result));
		if (resultItem is null) throw new ArgumentException(nameof(resultItem));
		if (worksheet is null) throw new ArgumentException(nameof(worksheet));
		if (tag.Type != WvTemplateTagType.Function)
			throw new ArgumentException("Template tag is not Function type", nameof(tag));

		if (String.IsNullOrWhiteSpace(tag.FunctionName))
			throw new Exception($"Unsupported function name: {tag.Name} in tag");

		object? resultValue = null;
		if (tag.ParamGroups.Count > 0
			&& tag.ParamGroups[0].Parameters.Count > 0
			&& !String.IsNullOrWhiteSpace(tag.FullString))
		{
			long? min = null;
			foreach (var parameter in tag.ParamGroups[0].Parameters)
			{
				if (String.IsNullOrWhiteSpace(parameter.ValueString)) continue;
				var range = new WvExcelRangeHelpers().GetRangeFromString(parameter.ValueString ?? String.Empty);
				if (range is not null)
				{
					var rangeTemplateContexts = result.TemplateContexts.GetIntersections(
						worksheetPosition: worksheet.Position,
						range: range,
						type: WvExcelFileTemplateContextType.CellRange
					);
					var resultContexts = resultItem.ResultContexts.Where(x => rangeTemplateContexts.Any(y => y.Id == x.TemplateContextId));
					var processedCellsHS = new HashSet<string>();
					foreach (var resContext in resultContexts)
					{
						var firstAddress = resContext.Range!.RangeAddress.FirstAddress;
						var lastAddress = resContext.Range!.RangeAddress.LastAddress;
						for (var rowNum = firstAddress.RowNumber; rowNum <= lastAddress.RowNumber; rowNum++)
						{
							for (var colNum = firstAddress.ColumnNumber; colNum <= lastAddress.ColumnNumber; colNum++)
							{
								var resultCell = worksheet.Cell(rowNum, colNum);
								if (processedCellsHS.Contains(resultCell.Address.ToString() ?? "")) continue;
								var mergedRange = resultCell.MergedRange();
								var mergedRows = 1;
								var mergedCols = 1;
								if (mergedRange != null)
								{
									foreach (var cell in mergedRange.Cells())
									{
										processedCellsHS.Add(cell.Address.ToString() ?? "");
									}
									mergedRows = mergedRange.RowCount();
									mergedCols = mergedRange.ColumnCount();
								}
								else
								{
									processedCellsHS.Add(resultCell.Address.ToString() ?? "");
								}

								var valueString = resultCell.Value.ToString();
								if (long.TryParse(valueString, out long longValue))
								{
									if (min is null || min > longValue)
										min = longValue;
								}
							}
						}
					}
				}
			}
			if (min is null)
			{
				HasError = true;
				ErrorMessage = "no suitable values found in range";
				return null;
			};

			if (!String.IsNullOrWhiteSpace(tagValue))
			{
				if (tagValue == tag.FullString)
					resultValue = min;
				else
					resultValue = ((string)tagValue).Replace(tag.FullString ?? String.Empty, min.ToString());
			}
		}

		return resultValue;
	}
}

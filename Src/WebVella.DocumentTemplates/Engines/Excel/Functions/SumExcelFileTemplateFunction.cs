using ClosedXML.Excel;
using System.Data;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Engines.Excel.Models;
using WebVella.DocumentTemplates.Engines.Excel.Utility;

namespace WebVella.DocumentTemplates.Engines.Excel.Functions;
public class SumExcelFileTemplateFunction : IWvExcelFileTemplateFunctionProcessor
{
	public string Name { get; } = "sum";
	public int Priority { get; } = 10000;
	public bool HasError { get; set; }
	public string? ErrorMessage { get; set; }

	public object? Process(
			object? value,
			WvTemplateTag tag,
			WvTemplateTagResultList input,
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

		var resultTagList = new WvTemplateTagResultList()
		{
			Tags = input.Tags.ToList(),
			Values = new(),
		};


		if (String.IsNullOrWhiteSpace(tag.FunctionName))
			throw new Exception($"Unsupported function name: {tag.Name} in tag");

		if (tag.ParamGroups.Count > 0
			&& tag.ParamGroups[0].Parameters.Count > 0
			&& !String.IsNullOrWhiteSpace(tag.FullString))
		{
			long sum = 0;
			foreach (var parameter in tag.ParamGroups[0].Parameters)
			{
				if (String.IsNullOrWhiteSpace(parameter.ValueString)) continue;
				var range = WvExcelRangeHelpers.GetRangeFromString(parameter.ValueString ?? String.Empty);
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
									sum += longValue;
								}
							}
						}
					}
				}
			}
			foreach (var tagValue in input.Values)
			{
				if (tagValue is null) continue;
				if (tagValue is not null && tagValue is string)
				{
					if (input.Tags.Count == 1)
						resultTagList.Values.Add(sum);
					else
						resultTagList.Values.Add(((string)tagValue).Replace(tag.FullString ?? String.Empty, sum.ToString()));
				}
				else
				{
					resultTagList.Values.Add(tagValue!);
				}
			}
		}

		return resultTagList;
	}
}

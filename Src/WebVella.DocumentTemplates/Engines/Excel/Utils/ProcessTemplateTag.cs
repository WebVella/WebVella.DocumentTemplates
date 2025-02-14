using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using System.Data;
using System.Globalization;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Extensions;

namespace WebVella.DocumentTemplates.Engines.Excel.Utility;
public static partial class WvExcelFileEngineUtility
{
	private static string getCellAddress(int worksheetPosition, int row, int col) => $"{worksheetPosition}:{row}:{col}";

	private static void CopyCellProperties(IXLCell template, IXLRange result)
	{
		//Commenting for optimization purposes
		//foreach (IXLCell destination in result.Cells())
		//{
		//	destination.ShowPhonetic = template.ShowPhonetic;
		//	destination.FormulaA1 = template.FormulaA1;
		//	destination.FormulaR1C1 = template.FormulaR1C1;
		//	if (template.FormulaReference is not null)
		//		destination.FormulaReference = template.FormulaReference;
		//	destination.ShareString = template.ShareString;
		//	destination.Style = template.Style;
		//	destination.Active = template.Active;
		//}
		if (!String.IsNullOrWhiteSpace(template.FormulaA1))
			result.FormulaA1 = template.FormulaA1;
		if (!String.IsNullOrWhiteSpace(template.FormulaR1C1))
			result.FormulaR1C1 = template.FormulaR1C1;

		result.ShareString = template.ShareString;
		result.Style = template.Style;
	}

	private static void CopyRowProperties(IXLRow origin, IXLRow result)
	{
		result.OutlineLevel = origin.OutlineLevel;
		result.Height = origin.Height;
	}

	private static void CopyRowProperties(IXLRange templateRange, IXLRange resultRange, IXLWorksheet templateWs, IXLWorksheet resultWs)
	{
		var resultRangeRowsCount = resultRange.RowCount();
		for (var i = 0; i < templateRange.RowCount(); i++)
		{
			if (resultRangeRowsCount - 1 < i) continue;
			var tempRow = templateWs.Row(templateRange.RangeAddress.FirstAddress.RowNumber + i);
			var resultRow = resultWs.Row(resultRange.RangeAddress.FirstAddress.RowNumber + i);
			CopyRowProperties(tempRow, resultRow);
		}
	}
	private static void CopyColumnsProperties(IXLRange templateRange, IXLRange resultRange, IXLWorksheet templateWs, IXLWorksheet resultWs)
	{
		var dsRowResultColumnsCount = resultRange.ColumnCount();
		for (var i = 0; i < templateRange.ColumnCount(); i++)
		{
			if (dsRowResultColumnsCount - 1 < i) continue;
			var tempColumn = templateWs.Column(templateRange.RangeAddress.FirstAddress.ColumnNumber + i);
			var resultColumn = resultWs.Column(resultRange.RangeAddress.FirstAddress.ColumnNumber + i);
			resultColumn.OutlineLevel = tempColumn.OutlineLevel;
			resultColumn.Width = tempColumn.Width;
		}
	}

	private static bool IsFlowHorizontal(WvTemplateTagResultList tagResult)
	{
		var allTagsHorizontal = true;
		foreach (var tag in tagResult.Tags)
		{
			if (tag.ParamGroups.Count == 0)
			{
				allTagsHorizontal = false;
				break;
			}
			var paramGroup = tag.ParamGroups[0];
			if (paramGroup.Parameters.Count == 0)
			{
				allTagsHorizontal = false;
				break;
			}
			foreach (var parameter in paramGroup.Parameters)
			{
				if (parameter.Type.FullName == typeof(WvTemplateTagDataFlowParameter).FullName)
				{
					var paramObj = (WvTemplateTagDataFlowParameter)parameter;
					if (paramObj.Value == WvTemplateTagDataFlow.Vertical)
					{
						allTagsHorizontal = false;
						break;
					}
				}
			}
		}
		return allTagsHorizontal;
	}
	private static void addRangeToGrid(int firstRow, int firstColumn, int lastRow, int lastColumn, Guid contextId,
		WvExcelFileTemplateProcessResultItemRow resultRow, int resultFirstRow)
	{
		if (firstRow > 1)
		{
			firstRow -= resultFirstRow - 1;
			lastRow -= resultFirstRow - 1;
		}
		for (int rowNum = firstRow; rowNum <= lastRow; rowNum++)
		{
			for (int colNum = firstColumn; colNum <= lastColumn; colNum++)
			{
				if (resultRow.RelativeGridAddressHs.Contains($"{rowNum}:{colNum}"))
					throw new Exception($"Context overlap at result row: {rowNum} column: {colNum}");

				resultRow.RelativeGrid.Add(new WvExcelFileTemplateProcessResultItemRowGridItem
				{
					Row = rowNum,
					Column = colNum,
					ContextId = contextId,
				});
				resultRow.RelativeGridAddressHs.Add($"{rowNum}:{colNum}");
			}
		}

		if (resultRow.RelativeGridMaxRow < lastRow) resultRow.RelativeGridMaxRow = lastRow;
		if (resultRow.RelativeGridMaxColumn < lastColumn) resultRow.RelativeGridMaxColumn = lastColumn;
	}

	private static WvTemplateTagResultList postProcessTemplateTagListForFunction(
		WvTemplateTagResultList input,
		DataTable dataSource,
		WvExcelFileTemplateProcessResult result,
		WvExcelFileTemplateProcessResultItem resultItem,
		IXLWorksheet worksheet
		)
	{
		if (!input.Tags.Any(x => x.Type == WvTemplateTagType.Function)) return input;
		if (input.Values.Count == 0) return input;

		var resultTagList = new WvTemplateTagResultList()
		{
			Tags = input.Tags.ToList(),
			Values = new(),
		};

		foreach (var tag in input.Tags.Where(x => x.Type == WvTemplateTagType.Function))
		{
			if (String.IsNullOrWhiteSpace(tag.FunctionName))
				throw new Exception($"Unsupported function name: {tag.Name} in tag");
			switch (tag.FunctionName)
			{
				case "sum":
					{
						if (tag.ParamGroups.Count > 0
							&& tag.ParamGroups[0].Parameters.Count > 0
							&& !String.IsNullOrWhiteSpace(tag.ParamGroups[0].Parameters[0].ValueString)
							&& !String.IsNullOrWhiteSpace(tag.FullString))
						{
							var range = WvExcelRangeHelpers.GetRangeFromString(tag.ParamGroups[0].Parameters[0].ValueString ?? String.Empty);
							if (range is not null)
							{
								var rangeTemplateContexts = result.TemplateContexts.GetIntersections(
									worksheetPosition: worksheet.Position,
									range: range,
									type: WvExcelFileTemplateContextType.CellRange
								);
								long sum = 0;
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

								foreach (var tagValue in input.Values)
								{
									if (tagValue is null) continue;
									if (tagValue is not null && tagValue is string)
									{
										if(input.Tags.Count == 1)
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
						}
						break;
					}
				default:
					throw new Exception($"Unsupported function name: {tag.Name} in tag");
			}
		}

		return resultTagList;
	}
}

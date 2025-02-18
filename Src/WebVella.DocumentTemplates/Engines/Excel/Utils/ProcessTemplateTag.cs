using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using System.Data;
using System.Globalization;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Extensions;

namespace WebVella.DocumentTemplates.Engines.Excel.Utility;
public partial class WvExcelFileEngineUtility
{
	public WvTemplateTagResultList ProcessTemplateTag(
		string? template,
		DataTable dataSource,
		WvExcelFileTemplateContext templateContext,
		WvExcelFileTemplateProcessResultItem resultItem)
	{
		var result = new WvTemplateTagResultList();
		result.Tags = new WvTemplateUtility().GetTagsFromTemplate(template);
		result.ExpandCount = 1;
		foreach (var tag in result.Tags)
		{
			var hasSeparator = tag.ParamGroups.Any(g => g.Parameters.Any(x => x.Type.FullName == typeof(WvTemplateTagSeparatorParameterProcessor).FullName));
			if (tag.Type == WvTemplateTagType.Data)
			{
				//If not indexed and no separator is defined return the row count
				if (tag.IndexList.Count == 0 && !hasSeparator)
				{
					if (result.ExpandCount < dataSource.Rows.Count)
						result.ExpandCount = dataSource.Rows.Count;
				}
			}
			else if (tag.Type == WvTemplateTagType.Function
			|| tag.Type == WvTemplateTagType.ExcelFunction)
			{
				if (result.Tags.Count == 0)
				{
					//Expand with the parent if:
					//Parent is LeftContext
					//Parent is Forced
					if (
						(templateContext.LeftContext is not null
							&& templateContext.ParentContext is not null
							&& templateContext.LeftContext.Id == templateContext.ParentContext.Id)
						|| (templateContext.ForcedContext is not null
							&& templateContext.ParentContext is not null
							&& templateContext.ForcedContext.Id == templateContext.ParentContext.Id)
							)
					{
						var parentResult = resultItem.ResultContexts.Single(x => x.TemplateContextId == templateContext.ParentContext.Id);
						if (result.ExpandCount < parentResult.ExpandCount)
							result.ExpandCount = parentResult.ExpandCount;
					}
					//Do not expand in other cases
				}
			}
		}

		return result;
	}
	private string getCellAddress(int worksheetPosition, int row, int col) => $"{worksheetPosition}:{row}:{col}";

	private void CopyCellProperties(IXLCell template, IXLRange result)
	{
		if (!String.IsNullOrWhiteSpace(template.FormulaA1))
			result.FormulaA1 = template.FormulaA1;
		if (!String.IsNullOrWhiteSpace(template.FormulaR1C1))
			result.FormulaR1C1 = template.FormulaR1C1;

		result.ShareString = template.ShareString;
		result.Style = template.Style;
		//We need to copy theme color differently as themed colors as the palletes can differ
		//for different types of cells - text or number
		//We need to copy theme color differently as themed colors as the palletes can differ
		//for different types of cells - text or number
		var templateColor = template.GetColor();
		var templateBackgroundColor = template.GetBackgroundColor();
		if (templateColor is not null)
		{
			result.Style.Font.SetFontColor(XLColor.FromColor(templateColor.Value));
		}
		if (templateBackgroundColor is not null)
		{
			result.Style.Fill.SetBackgroundColor(XLColor.FromColor(templateBackgroundColor.Value));
		}

	}

	private XLColor? GetThemeColor(XLThemeColor themeColor, XLWorkbook workbook)
	{
		XLColor color;
		switch (themeColor)
		{
			case XLThemeColor.Background1:
				color = workbook.Theme.Background1;
				break;
			case XLThemeColor.Background2:
				color = workbook.Theme.Background2;
				break;
			case XLThemeColor.Text1:
				color = workbook.Theme.Text1;
				break;
			case XLThemeColor.Text2:
				color = workbook.Theme.Text2;
				break;
			case XLThemeColor.Accent1:
				color = workbook.Theme.Accent1;
				break;
			case XLThemeColor.Accent2:
				color = workbook.Theme.Accent2;
				break;
			case XLThemeColor.Accent3:
				color = workbook.Theme.Accent3;
				break;
			case XLThemeColor.Accent4:
				color = workbook.Theme.Accent4;
				break;
			case XLThemeColor.Accent5:
				color = workbook.Theme.Accent5;
				break;
			case XLThemeColor.Accent6:
				color = workbook.Theme.Accent6;
				break;
			case XLThemeColor.Hyperlink:
				color = workbook.Theme.Hyperlink;
				break;
			case XLThemeColor.FollowedHyperlink:
				color = workbook.Theme.FollowedHyperlink;
				break;
			default:
				return null;
		}
		return XLColor.FromArgb(color.Color.A, color.Color.R, color.Color.G, color.Color.B);
	}

	private void CopyRowProperties(IXLRow origin, IXLRow result)
	{
		result.OutlineLevel = origin.OutlineLevel;
		result.Height = origin.Height;
	}

	private void CopyRowProperties(IXLRange templateRange, IXLRange resultRange,
		IXLWorksheet templateWs, IXLWorksheet resultWs,
		HashSet<long> processedRows)
	{
		var resultRangeRowsCount = resultRange.RowCount();
		for (var i = 0; i < templateRange.RowCount(); i++)
		{
			if (processedRows.Contains(i + 1)) continue;
			processedRows.Add(i + 1);
			if (resultRangeRowsCount - 1 < i) continue;
			var tempRow = templateWs.Row(templateRange.RangeAddress.FirstAddress.RowNumber + i);
			var resultRow = resultWs.Row(resultRange.RangeAddress.FirstAddress.RowNumber + i);
			CopyRowProperties(tempRow, resultRow);
		}
	}
	private void CopyColumnsProperties(IXLRange templateRange, IXLRange resultRange,
		IXLWorksheet templateWs, IXLWorksheet resultWs,
		HashSet<long> processedColumns)
	{
		var dsRowResultColumnsCount = resultRange.ColumnCount();
		for (var i = 0; i < templateRange.ColumnCount(); i++)
		{
			if (processedColumns.Contains(i + 1)) continue;
			processedColumns.Add(i + 1);
			if (dsRowResultColumnsCount - 1 < i) continue;
			var tempColumn = templateWs.Column(templateRange.RangeAddress.FirstAddress.ColumnNumber + i);
			var resultColumn = resultWs.Column(resultRange.RangeAddress.FirstAddress.ColumnNumber + i);
			resultColumn.OutlineLevel = tempColumn.OutlineLevel;
			resultColumn.Width = tempColumn.Width;
		}
	}

	private bool IsFlowHorizontal(WvTemplateTagResultList tagResult)
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
				if (parameter.Type.FullName == typeof(WvTemplateTagDataFlowParameterProcessor).FullName)
				{
					var paramObj = (WvTemplateTagDataFlowParameterProcessor)parameter;
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
	private void addRangeToGrid(int firstRow, int firstColumn, int lastRow, int lastColumn, Guid contextId,
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


	private WvTemplateTagResultList postProcessTemplateTagListForFunction(
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
							&& !String.IsNullOrWhiteSpace(tag.FullString))
						{
							long sum = 0;
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
						break;
					}
				default:
					throw new Exception($"Unsupported function name: {tag.Name} in tag");
			}
		}

		return resultTagList;
	}
}

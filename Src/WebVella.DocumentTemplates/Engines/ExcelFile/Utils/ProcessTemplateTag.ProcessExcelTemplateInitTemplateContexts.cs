﻿using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;

namespace WebVella.DocumentTemplates.Engines.ExcelFile.Utility;
public partial class WvExcelFileEngineUtility
{
	/// <summary>
	/// Parses the Template excel and generates contexts
	/// </summary>
	/// <param name="result"></param>
	/// <exception cref="ArgumentException"></exception>
	public void ProcessExcelTemplateInitTemplateContexts(WvExcelFileTemplateProcessResult? result)
	{
		if (result is null) throw new ArgumentException("No result provided!", nameof(result));
		if (result.Template is null) throw new ArgumentException("No Template provided!", nameof(result));
		if (result.Template.Worksheets is null || result.Template.Worksheets.Count == 0) throw new ArgumentException("No worksheets in template provided!", nameof(result));

		//Cell contexts
		var processedAddresses = new HashSet<string>();
		foreach (IXLWorksheet ws in result.Template.Worksheets)
		{
			var (usedRowsCount, usedColumnsCount) = ws.GetGrossUsedRange();

			for (var rowPosition = 1; rowPosition <= usedRowsCount; rowPosition++)
			{
				var templateRow = new WvExcelFileTemplateRow
				{
					Worksheet = ws,
					Position = rowPosition,
					Contexts = new(),
				};
				for (var colPosition = 1; colPosition <= usedColumnsCount; colPosition++)
				{
					var cellAddress = getCellAddress(ws.Position, rowPosition, colPosition);
					if (processedAddresses.Contains(cellAddress)) continue;

					IXLCell cell = ws.Cell(rowPosition, colPosition);
					var mergedRange = cell.MergedRange();
					//Mark all cell addresses in the range
					if (mergedRange is not null)
					{
						int mergeRangeFirstRow = mergedRange.FirstRow().RowNumber();
						int mergeRangeLastRow = mergedRange.LastRow().RowNumber();
						int mergeRangeFirstColumn = mergedRange.FirstColumn().ColumnNumber();
						int mergeRangeLastColumn = mergedRange.LastColumn().ColumnNumber();

						for (var mergeRangeRow = mergeRangeFirstRow; mergeRangeRow <= mergeRangeLastRow; mergeRangeRow++)
						{
							for (var mergeRangeCol = mergeRangeFirstColumn; mergeRangeCol <= mergeRangeLastColumn; mergeRangeCol++)
							{
								processedAddresses.Add(getCellAddress(ws.Position, mergeRangeRow, mergeRangeCol));
							}
						}
					}
					else
					{
						processedAddresses.Add(cellAddress);
					}

					WvTemplateTagDataFlow? forcedCellDataFlow = null;
					var cellTags = new WvTemplateUtility().GetTagsFromTemplate(cell.Value.ToString());
					if (rowPosition == 1 && colPosition == 1)
					{
						forcedCellDataFlow = WvTemplateTagDataFlow.Vertical;
					}
					if (cellTags.Count > 0)
					{
						var hasUndecided = false;
						var hasVertical = false;
						var hasHorizontal = false;
						foreach (var tag in cellTags)
						{
							if (tag.Flow is null) hasUndecided = true;
							else if (tag.Flow.Value is WvTemplateTagDataFlow.Vertical) hasVertical = true;
							else if (tag.Flow.Value is WvTemplateTagDataFlow.Horizontal) hasHorizontal = true;
						}
						if (hasUndecided && !hasVertical && !hasHorizontal)
							forcedCellDataFlow = null;
						else if (!hasUndecided && hasVertical && !hasHorizontal)
							forcedCellDataFlow = WvTemplateTagDataFlow.Vertical;
						else if (!hasUndecided && !hasVertical && hasHorizontal)
							forcedCellDataFlow = WvTemplateTagDataFlow.Horizontal;
						else if (hasUndecided && hasVertical)
							forcedCellDataFlow = WvTemplateTagDataFlow.Vertical;

					}
					WvExcelFileTemplateContext? leftContext = null;
					WvExcelFileTemplateContext? topContext = null;
					if (rowPosition > 1)
					{
						topContext = result.TemplateContexts
							.GetByAddress(ws.Position, rowPosition - 1, colPosition, WvExcelFileTemplateContextType.CellRange)
							.FirstOrDefault();
					}
					if (colPosition > 1)
					{
						leftContext = result.TemplateContexts
							.GetByAddress(ws.Position, rowPosition, colPosition - 1, WvExcelFileTemplateContextType.CellRange)
							.FirstOrDefault();
					}

					WvExcelFileTemplateContext? forcedParentContext = null;
					bool isNullForcedParentContext = false;
					foreach (WvTemplateTag tag in cellTags)
					{
						if (isNullForcedParentContext || forcedParentContext is not null) break;
						foreach (var paramGroup in tag.ParamGroups)
						{
							if (isNullForcedParentContext || forcedParentContext is not null) break;

							foreach (var param in paramGroup.Parameters)
							{
								if (isNullForcedParentContext || forcedParentContext is not null) break;
								if (param.Type != typeof(WvTemplateTagParentContextParameterProcessor)) continue;
								if (String.IsNullOrWhiteSpace(param.ValueString)) continue;

								if (param.ValueString.ToLowerInvariant() == "none"
									|| param.ValueString.ToLowerInvariant() == "null")
								{
									isNullForcedParentContext = true;
								}
								else
								{
									var range = new WvExcelRangeHelpers().GetRangeFromString(param.ValueString ?? String.Empty);
									if (range is not null)
									{
										forcedParentContext = result.TemplateContexts.GetIntersections(
																worksheetPosition: ws.Position,
																range: range,
																type: WvExcelFileTemplateContextType.CellRange
															).FirstOrDefault();
									}
								}
							}
						}
					}


					//Create context
					var context = new WvExcelFileTemplateContext()
					{
						Id = Guid.NewGuid(),
						Worksheet = ws,
						Row = templateRow,
						Range = mergedRange is not null ? mergedRange : cell.AsRange(),
						Type = WvExcelFileTemplateContextType.CellRange,
						ForcedFlow = forcedCellDataFlow,
						ForcedContext = forcedParentContext,
						IsNullContextForced = isNullForcedParentContext,
						Picture = null,
						LeftContext = leftContext,
						TopContext = topContext,
						ContextDependencies = new()
					};
					templateRow.Contexts.Add(context.Id);
					result.TemplateContexts.Add(context);
				}
				result.TemplateRows.Add(templateRow);
			}
		}

		//Picture contexts
		foreach (IXLWorksheet ws in result.Template.Worksheets)
		{
			foreach (IXLPicture picture in ws.Pictures)
			{
				int rowPosition = picture.TopLeftCell.Address.RowNumber;
				int colPosition = picture.TopLeftCell.Address.ColumnNumber;
				var cell = ws.Cell(rowPosition, colPosition);

				//Create context
				var context = new WvExcelFileTemplateContext()
				{
					Id = Guid.NewGuid(),
					Worksheet = ws,
					Row = null,
					Range = null,
					Picture = picture,
					ForcedFlow = WvTemplateTagDataFlow.Vertical,
					Type = WvExcelFileTemplateContextType.Picture,
					LeftContext = null,
					TopContext = null,
					ContextDependencies = new()
				};
				result.TemplateContexts.Add(context);
			}
		}
	}
}

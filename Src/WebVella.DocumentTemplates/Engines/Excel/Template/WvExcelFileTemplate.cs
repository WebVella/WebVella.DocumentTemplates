﻿using ClosedXML.Excel;
using System.Data;
using System.Globalization;
using System.Text;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Extensions;
namespace WebVella.DocumentTemplates.Engines.Excel;
public class WvExcelFileTemplate : WvTemplateBase
{
	public XLWorkbook? Template { get; set; }
	public WvExcelFileTemplateProcessResult Process(DataTable? dataSource, CultureInfo? culture = null)
	{
		if (culture == null) culture = new CultureInfo("en-US");
		if (dataSource is null) throw new ArgumentException("No datasource provided!", nameof(dataSource));
		var result = new WvExcelFileTemplateProcessResult()
		{
			Template = Template,
			GroupByDataColumns = GroupDataByColumns,
			ResultItems = new()
		};
		if (Template is null)
		{
			result.ResultItems.Add(new WvExcelFileTemplateProcessResultItem
			{
				Result = null
			});
			return result;
		};

		var datasourceGroups = dataSource.GroupBy(GroupDataByColumns);
		foreach (var grouptedDs in datasourceGroups)
		{
			var resultItem = new WvExcelFileTemplateProcessResultItem{ 
				Result = new XLWorkbook()
			};
			ProcessExcelTemplatePlacement(Template,resultItem, grouptedDs, culture);
			ProcessExcelTemplateDependencies(resultItem, grouptedDs);
			ProcessExcelTemplateData(resultItem, grouptedDs, culture);

			result.ResultItems.Add(resultItem);
		}
		return result;
	}

	public void ProcessExcelTemplatePlacement(XLWorkbook template, WvExcelFileTemplateProcessResultItem resultItem, DataTable dataSource, CultureInfo culture)
	{
		if (resultItem is null) throw new Exception("No result provided!");
		if (dataSource is null) throw new Exception("No datasource provided!");
		if (template is null) throw new Exception("No Template provided!");
		if (resultItem.Result is null) resultItem.Result = new XLWorkbook();
		foreach (IXLWorksheet tempWs in template.Worksheets)
		{
			var tempWsUsedRowsCount = tempWs.LastRowUsed()?.RowNumber() ?? 1;
			var tempWsUsedColumnsCount = tempWs.LastColumnUsed()?.ColumnNumber() ?? 1;
			var resultWs = resultItem.Result.AddWorksheet();
			resultWs.Name = tempWs.Name;
			var resultCurrentRow = 1;
			for (var rowIndex = 0; rowIndex < tempWsUsedRowsCount; rowIndex++)
			{
				IXLRow tempRow = tempWs.Row(rowIndex + 1);
				var resultCurrentColumn = 1;
				var tempRowResultRowsUsed = 0;

				for (var colIndex = 0; colIndex < tempWsUsedColumnsCount; colIndex++)
				{
					IXLCell tempCell = tempWs.Cell(tempRow.RowNumber(), (colIndex + 1));
					var tempRangeStartRowNumber = resultCurrentRow;
					var tempRangeStartColumnNumber = resultCurrentColumn;
					var tempRangeEndRowNumber = resultCurrentRow;
					var tempRangeEndColumnNumber = resultCurrentColumn;
					var templateMergedRowCount = 1;
					var templateMergedColumnCount = 1;
					var mergedRange = tempCell.MergedRange();
					List<WvExcelFileTemplateContextRangeAddress> resultRangeSlots = new();
					//Process only the first cell in the merge
					if (mergedRange is not null)
					{
						var margeRangeCells = mergedRange.Cells();
						if (margeRangeCells.First().Address.ToString() != tempCell.Address.ToString())
						{
							continue;
						}
						templateMergedRowCount = mergedRange.RowCount();
						templateMergedColumnCount = mergedRange.ColumnCount();
					}

					tempRangeEndRowNumber = tempRangeEndRowNumber + templateMergedRowCount - 1;
					tempRangeEndColumnNumber = tempRangeEndColumnNumber + templateMergedColumnCount - 1;

					var tagProcessResult = WvTemplateUtility.ProcessTemplateTag(tempCell.Value.ToString(), dataSource, culture);
					var contextDataProcessed = false;
					if (tagProcessResult.Tags.Count > 0)
					{
						var isFlowHorizontal = IsFlowHorizontal(tagProcessResult);
						for (var i = 0; i < tagProcessResult.Values.Count; i++)
						{
							var dsRowStarRowNumber = tempRangeStartRowNumber;
							var dsRowStartColumnNumber = tempRangeStartColumnNumber;

							var dsRowEndRowNumber = tempRangeEndRowNumber;
							var dsRowEndColumnNumber = tempRangeEndColumnNumber;

							//if vertical - increase rows with the merge
							if (!isFlowHorizontal)
							{
								dsRowStarRowNumber = tempRangeStartRowNumber + (templateMergedRowCount * i);
								dsRowEndRowNumber = tempRangeEndRowNumber + (templateMergedRowCount * i);
							}
							//if hirozontal - increase columns with the merge
							else
							{
								dsRowStartColumnNumber = tempRangeStartColumnNumber + (templateMergedColumnCount * i);
								dsRowEndColumnNumber = tempRangeEndColumnNumber + (templateMergedColumnCount * i);
							}

							IXLRange dsRowResultRange = resultWs.Range(
								firstCellRow: dsRowStarRowNumber,
								firstCellColumn: dsRowStartColumnNumber,
								lastCellRow: dsRowEndRowNumber,
								lastCellColumn: dsRowEndColumnNumber);

							dsRowResultRange.Merge();

							if (tagProcessResult.AllDataTags)
							{
								dsRowResultRange.Value = XLCellValue.FromObject(tagProcessResult.Values[i]);
								contextDataProcessed = true;
							}


							//Boz: for optimization purposes is move outside the loop
							//CopyCellProperties(tempCell, dsRowResultRange);
							//var templateRange = mergedRange is not null ? mergedRange : tempCell.AsRange();
							//CopyColumnsProperties(templateRange, dsRowResultRange, tempWs, resultWs);
							resultRangeSlots.Add(new WvExcelFileTemplateContextRangeAddress(
								firstRow: dsRowResultRange.RangeAddress.FirstAddress.RowNumber,
								firstColumn: dsRowResultRange.RangeAddress.FirstAddress.ColumnNumber,
								lastRow: dsRowResultRange.RangeAddress.LastAddress.RowNumber,
								lastColumn: dsRowResultRange.RangeAddress.LastAddress.ColumnNumber
							));
						}
						if (!isFlowHorizontal)
						{
							tempRangeEndRowNumber = tempRangeStartRowNumber + (tagProcessResult.Values.Count * templateMergedRowCount) - 1;
						}
						else
						{
							tempRangeEndColumnNumber = tempRangeStartColumnNumber + (tagProcessResult.Values.Count * templateMergedColumnCount) - 1;
						}
					}
					else
					{
						tempRangeEndRowNumber = resultCurrentRow + templateMergedRowCount - 1;
						tempRangeEndColumnNumber = resultCurrentColumn + templateMergedColumnCount - 1;
						IXLRange dsRowResultRange = resultWs.Range(
								firstCellRow: tempRangeStartRowNumber,
								firstCellColumn: tempRangeStartColumnNumber,
								lastCellRow: tempRangeEndRowNumber,
								lastCellColumn: tempRangeEndColumnNumber);
						dsRowResultRange.Merge();
						dsRowResultRange.Value = tempCell.Value;
						contextDataProcessed = true;
						//Boz: for optimization purposes is move outside the loop
						//CopyCellProperties(tempCell, dsRowResultRange);
						//var templateRange = mergedRange is not null ? mergedRange : tempCell.AsRange();
						//CopyColumnsProperties(templateRange, dsRowResultRange, tempWs, resultWs);

						resultRangeSlots.Add(new WvExcelFileTemplateContextRangeAddress(
								firstRow: dsRowResultRange.RangeAddress.FirstAddress.RowNumber,
								firstColumn: dsRowResultRange.RangeAddress.FirstAddress.ColumnNumber,
								lastRow: dsRowResultRange.RangeAddress.LastAddress.RowNumber,
								lastColumn: dsRowResultRange.RangeAddress.LastAddress.ColumnNumber
							));
					}

					IXLRange resultRange = resultWs.Range(
							firstCellRow: tempRangeStartRowNumber,
							firstCellColumn: tempRangeStartColumnNumber,
							lastCellRow: tempRangeEndRowNumber,
							lastCellColumn: tempRangeEndColumnNumber);

					//Copy column styles


					//Create context
					var context = new WvExcelFileTemplateProcessContext()
					{
						Id = Guid.NewGuid(),
						TagProcessResult = tagProcessResult,
						Tags = tagProcessResult.Tags,
						TemplateWorksheet = tempWs.Worksheet,
						ResultWorksheet = resultWs.Worksheet,
						TemplateRange = mergedRange is not null ? mergedRange : tempCell.AsRange(),
						ResultRange = resultRange,
						ResultRangeSlots = resultRangeSlots,
						IsDataSet = contextDataProcessed
					};
					resultItem.Contexts.Add(context);

					//Boz: for optimization purposes is move from inside the loop
					CopyCellProperties(tempCell, context.ResultRange);
					var templateRange = mergedRange is not null ? mergedRange : tempCell.AsRange();
					CopyColumnsProperties(templateRange, context.ResultRange, tempWs, resultWs);

					if (tempRowResultRowsUsed < resultRange.RowCount())
						tempRowResultRowsUsed = resultRange.RowCount();

					resultCurrentColumn = resultCurrentColumn + resultRange.ColumnCount();
				}
				//Copy Row Styles
				for (var i = 0; i < tempRowResultRowsUsed; i++)
				{
					var resultRow = resultWs.Row(resultCurrentRow + i);
					CopyRowProperties(tempRow, resultRow);
				}

				resultCurrentRow += tempRowResultRowsUsed;
			}
		}
	}

	public void ProcessExcelTemplateDependencies(WvExcelFileTemplateProcessResultItem resultItem, DataTable dataSource)
	{
		foreach (var context in resultItem.Contexts)
		{
			HashSet<Guid> contextDependancies = new();
			foreach (var tag in context.TagProcessResult.Tags)
			{
				foreach (var tagParamGroup in tag.ParamGroups)
				{
					foreach (var tagParam in tagParamGroup.Parameters)
					{
						var baseParam = (IWvTemplateTagParameterBase)tagParam;
						if (baseParam.Type.IsAssignableTo(typeof(IWvExcelFileTemplateTagParameter)))
						{
							var excelParam = (IWvExcelFileTemplateTagParameter)baseParam;
							contextDependancies.Union(excelParam.GetDependencies(
								resultItem: resultItem,
								context: context,
								tag: tag,
								parameterGroup: tagParamGroup,
								parameter: excelParam
							));
						}
					}
				}
			}
			context.Dependencies.Union(contextDependancies);
			foreach (var contextId in contextDependancies)
			{
				resultItem.Contexts.Single(x => x.Id == contextId).Dependants.Add(context.Id);
			}
		}
	}

	public void ProcessExcelTemplateData(WvExcelFileTemplateProcessResultItem resultItem, DataTable dataSource, CultureInfo culture)
	{
		int processAttemptsLimit = 200;
		if (resultItem is null) throw new Exception("No result provided!");
		if (dataSource is null) throw new Exception("No datasource provided!");
		if (resultItem.Result is null) resultItem.Result = new XLWorkbook();
		#region << Process Worksheet Names>>
		if (dataSource.Rows.Count > 0)
		{
			var firstRowDt = dataSource.CreateAsNew(new List<int> { 0 });
			var resultWorksheets = resultItem.Result.Worksheets.ToList();
			for (int i = 0; i < resultWorksheets.Count; i++)
			{
				var resultWs = resultWorksheets[i];
				var templateResult = WvTemplateUtility.ProcessTemplateTag(resultWs.Name, firstRowDt, culture);
				if (templateResult != null && templateResult.Values.Count > 0
					&& templateResult.Values[0] is not null
					&& templateResult.Values[0] is string
					&& !String.IsNullOrWhiteSpace((string)templateResult.Values[0]))
				{
					resultWs.Name = templateResult!.Values[0]!.ToString() ?? "";
				}
				else
				{
					resultWs.Name = $"Worksheet {i + 1}";
				}
			}
		}
		#endregion

		#region << Process Worksheet Data>>
		var contextDict = resultItem.Contexts.ToDictionary(x => x.Id);
		Queue<Guid> queue = new Queue<Guid>();
		foreach (var contextId in resultItem.Contexts.Where(x => x.Dependencies.Count == 0).Select(x => x.Id))
		{
			queue.Enqueue(contextId);
		}
		while (queue.Count > 0)
		{
			var contextId = queue.Dequeue();
			if (!resultItem.ContextProcessLog.ContainsKey(contextId))
				resultItem.ContextProcessLog[contextId] = 1;
			if (resultItem.ContextProcessLog[contextId] > processAttemptsLimit) continue;

			var context = contextDict[contextId];
			if (context.ResultWorksheet is null) continue;

			if (!context.IsDataSet)
			{
				var resultCount = context.TagProcessResult.Values.Count;
				for (var i = 0; i < context.ResultRangeSlots.Count; i++)
				{
					var address = context.ResultRangeSlots[i];
					var slot = context.ResultWorksheet.Range(
						address.FirstRow, address.FirstColumn,
						address.LastRow, address.LastColumn);
					if (resultCount < i + 1) break;
					slot.Value = XLCellValue.FromObject(context.TagProcessResult.Values[i]);
				}
			}
			resultItem.ProcessedContexts.Add(contextId);
			foreach (var dependantId in context.Dependants)
			{
				if (resultItem.ProcessedContexts.Contains(dependantId))
					throw new Exception("Dependant was calculated before the context it depends on");
				queue.Enqueue(dependantId);
			}
		};

		if (resultItem.ProcessedContexts.Count < resultItem.Contexts.Count)
		{
			throw new Exception("Not all excel cells were processed. Check for possible circular logic in formulas.");
		}
		#endregion
	}

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

	public static WvExcelFileTemplateProcessContext? FindContextByCell(WvExcelFileTemplateProcessResultItem resultItem, int worksheetPosition, int row, int column)
	{
		foreach (var item in resultItem.Contexts)
		{
			if (item.TemplateWorksheet is null) continue;
			if (item.TemplateRange is null) continue;
			if (item.TemplateWorksheet.Position != worksheetPosition) continue;
			if (
				item.TemplateRange.RangeAddress.FirstAddress.RowNumber >= row
				&& item.TemplateRange.RangeAddress.LastAddress.RowNumber <= row
				&& item.TemplateRange.RangeAddress.FirstAddress.ColumnNumber >= column
				&& item.TemplateRange.RangeAddress.LastAddress.ColumnNumber <= column
			)
			{
				return item;
			}
		}
		return null;
	}
}
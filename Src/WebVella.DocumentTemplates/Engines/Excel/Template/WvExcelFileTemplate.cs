using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using System.Data;
using System.Globalization;
using System.Text;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Engines.Excel.Utility;
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
			var resultItem = new WvExcelFileTemplateProcessResultItem
			{
				Result = new XLWorkbook()
			};
			ProcessExcelTemplateInitTemplateContexts(result);
			ProcessExcelTemplateCalculateDependencies(result);
			ProcessExcelTemplateGenerateResultContexts(result, resultItem, grouptedDs, culture);
			//ProcessExcelTemplatePlacementAndContexts(Template, resultItem, grouptedDs, culture);
			//ProcessExcelTemplateDependencies(resultItem, grouptedDs);
			//ProcessExcelTemplateData(resultItem, grouptedDs, culture);
			//ProcessExcelTemplateEmbeddings(Template, resultItem);

			result.ResultItems.Add(resultItem);
		}
		return result;
	}

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
				var contextBranch = new WvExcelFileTemplateContextBranch
				{
					Id = Guid.NewGuid(),
					TemplateContexts = new(),
					Worksheet = ws,
					Range = ws.Range(rowPosition, 1, rowPosition, usedColumnsCount)
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

					var cellFlow = WvTemplateTagDataFlow.Vertical;
					var cellTags = WvTemplateUtility.GetTagsFromTemplate(cell.Value.ToString());
					if (cellTags.Count > 0 && !cellTags.Any(x => x.Flow == WvTemplateTagDataFlow.Vertical))
					{
						cellFlow = WvTemplateTagDataFlow.Horizontal;
					}

					WvExcelFileTemplateContext? topContext = null;
					WvExcelFileTemplateContext? leftContext = null;
					if (rowPosition > 1) topContext = result.TemplateContexts.GetByAddress(
						worksheetPosition: ws.Position,
						row: rowPosition - 1,
						column: colPosition,
						type: WvExcelFileTemplateContextType.CellRange).FirstOrDefault();

					if (colPosition > 1) leftContext = result.TemplateContexts.GetByAddress(
						worksheetPosition: ws.Position,
						row: rowPosition,
						column: colPosition - 1,
						type: WvExcelFileTemplateContextType.CellRange).FirstOrDefault();

					//Create context
					var context = new WvExcelFileTemplateContext()
					{
						Id = Guid.NewGuid(),
						Worksheet = ws,
						Range = mergedRange is not null ? mergedRange : cell.AsRange(),
						Type = WvExcelFileTemplateContextType.CellRange,
						Flow = cellFlow,
						TopContext = topContext,
						LeftContext = leftContext,
						Picture = null,
						ContextDependencies = new()
					};
					contextBranch.TemplateContexts.Add(context);
					result.TemplateContexts.Add(context);
				}

				result.TemplateContextBranches.Add(contextBranch);
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
					Picture = picture,
					Flow = WvTemplateTagDataFlow.Vertical,
					Type = WvExcelFileTemplateContextType.Picture,
					TopContext = null,
					LeftContext = null,
					ContextDependencies = new()
				};
				//Copy top left context from cellRange if present
				var cellContext = result.TemplateContexts
					.GetByAddress(ws.Position, rowPosition, colPosition)
					.Where(x => x.Type == WvExcelFileTemplateContextType.CellRange)
					.FirstOrDefault();
				if (cellContext != null)
				{
					context.TopContext = cellContext.TopContext;
					context.LeftContext = cellContext.LeftContext;

				}
				else
				{
					if (rowPosition > 1) context.TopContext = result.TemplateContexts.GetByAddress(
						worksheetPosition: ws.Position,
						row: rowPosition - 1,
						column: colPosition,
						type: WvExcelFileTemplateContextType.CellRange).FirstOrDefault();

					if (colPosition > 1) context.LeftContext = result.TemplateContexts.GetByAddress(
						worksheetPosition: ws.Position,
						row: rowPosition,
						column: colPosition - 1,
						type: WvExcelFileTemplateContextType.CellRange).FirstOrDefault();
				}
				result.TemplateContexts.Add(context);
				var pictureBranch = result.TemplateContextBranches.GetIntersections(ws.Position, new Models.WvExcelRange
				{
					FirstRow = rowPosition,
					FirstColumn = colPosition,
					LastRow = rowPosition,
					LastColumn = colPosition,
				});
				if (pictureBranch.Count == 1)
				{
					pictureBranch[0].TemplateContexts.Add(context);
				}
				else
				{
					throw new Exception("Unsupported case. There should be only one branch corresponding to one templat cell");
				}
			}
		}
	}

	/// <summary>
	/// Generates dependencies between the contexts based on formula fields ranges
	/// </summary>
	/// <param name="result"></param>
	public void ProcessExcelTemplateCalculateDependencies(WvExcelFileTemplateProcessResult result)
	{
		foreach (var branch in result.TemplateContextBranches)
		{
			foreach (var context in branch.TemplateContexts)
			{
				context.CalculateDependencies(branch.TemplateContexts);
			}
		}
	}

	/// <summary>
	/// Creates result contexts and result Excel by processing the template contexts, fill them with data
	/// and situate them on the result Excel based on their parent Context
	/// </summary>
	/// <param name="result"></param>
	/// <param name="resultItem"></param>
	/// <param name="dataSource"></param>
	/// <param name="culture"></param>
	public void ProcessExcelTemplateGenerateResultContexts(WvExcelFileTemplateProcessResult result,
	 WvExcelFileTemplateProcessResultItem resultItem, DataTable dataSource, CultureInfo culture)
	{
		if (resultItem is null) throw new Exception("No result provided!");
		if (dataSource is null) throw new Exception("No datasource provided!");
		if (result.Template is null) throw new Exception("No Template provided!");
		if (result.TemplateContextBranches is null) throw new Exception("No Template provided!");
		if (resultItem.Result is null) resultItem.Result = new XLWorkbook();
		if (result.TemplateContexts.Count == 0) return;
		#region << Create Result Worksheets>>
		if (result.Template.Worksheets.Count == 0) return;
		var firstRowDt = dataSource.CreateAsNew(new List<int> { 0 });
		foreach (var templateWs in result.Template.Worksheets.OrderBy(x => x.Position))
		{
			var resultWs = resultItem.Result.AddWorksheet();
			resultWs.Name = templateWs.Name;
			var templateResult = WvTemplateUtility.ProcessTemplateTag(templateWs.Name, firstRowDt, culture);
			if (templateResult != null && templateResult.Values.Count > 0
				&& templateResult.Values[0] is not null
				&& templateResult.Values[0] is string
				&& !String.IsNullOrWhiteSpace((string)templateResult.Values[0]))
			{
				resultWs.Name = templateResult!.Values[0]!.ToString() ?? "";
			}
			else
			{
				resultWs.Name = $"Worksheet {templateWs.Position}";
			}

		}
		#endregion
		foreach (var branch in result.TemplateContextBranches)
		{
			var templateContextsDict = branch.TemplateContexts.ToDictionary(x => x.Id);
			var resultContextsDict = new Dictionary<Guid, WvExcelFileResultItemContext>();
			int processAttemptsLimit = 200;
			Queue<Guid> queue = new Queue<Guid>();
			foreach (var contextId in branch.TemplateContexts.Select(x => x.Id))
			{
				queue.Enqueue(contextId);
			}

			while (queue.Count > 0)
			{
				var contextId = queue.Dequeue();
				var templateContext = templateContextsDict[contextId];

				//Check Limit
				if (processLimitIsReached(
					templateContext: templateContext,
					resultItem: resultItem,
					limit: processAttemptsLimit
				)) continue;

				//Check if unfulfilled Dependency
				if (hasUnfulfilledDependencies(
					templateContext: templateContext,
					resultContextsDict: resultContextsDict,
					queue: queue
				)) continue;

				//Check if there is fulfulled dependency with error
				if (hasfulfilledDependencyWithError(
					templateContext: templateContext,
					resultItem: resultItem,
					resultContextsDict: resultContextsDict
				)) continue;

				//Generate result region
				try
				{
					var processedResultItemContext = generateResultRegion(
						templateContext: templateContext,
						resultItem: resultItem,
						resultContextsDict: resultContextsDict,
						dataSource: dataSource,
						culture: culture
					);
					if (processedResultItemContext is not null)
					{
						resultItem.ResultContexts.Add(processedResultItemContext);
						resultContextsDict[processedResultItemContext.TemplateContextId] = processedResultItemContext;
					}
				}
				catch (Exception ex)
				{
					resultItem.ResultContexts.Add(new WvExcelFileResultItemContext
					{
						TemplateContextId = templateContext.Id,
						Error = WvExcelFileResultItemContextError.ProcessError,
						ErrorMessage = ex.Message,
						Range = templateContext.Range
					});
				}
			}
		}
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

	#region << Private >>
	private bool processLimitIsReached(WvExcelFileTemplateContext templateContext,
		WvExcelFileTemplateProcessResultItem resultItem, int limit)
	{

		if (!resultItem.ContextProcessLog.ContainsKey(templateContext.Id))
			resultItem.ContextProcessLog[templateContext.Id] = 1;

		if (resultItem.ContextProcessLog[templateContext.Id] > limit)
		{
			resultItem.ResultContexts.Add(new WvExcelFileResultItemContext
			{
				TemplateContextId = templateContext.Id,
				Error = WvExcelFileResultItemContextError.DependencyOverflow,
				Range = templateContext.Range
			});
			return true;
		};
		resultItem.ContextProcessLog[templateContext.Id]++;
		return false;
	}
	private bool hasUnfulfilledDependencies(WvExcelFileTemplateContext templateContext,
		Dictionary<Guid, WvExcelFileResultItemContext> resultContextsDict,
		Queue<Guid> queue)
	{
		if (templateContext.ContextDependencies.Any(x => !resultContextsDict.Keys.Contains(x)))
		{
			queue.Enqueue(templateContext.Id);
			return true;
		}

		return false;
	}

	private bool hasfulfilledDependencyWithError(WvExcelFileTemplateContext templateContext,
		WvExcelFileTemplateProcessResultItem resultItem,
		Dictionary<Guid, WvExcelFileResultItemContext> resultContextsDict)
	{
		if (templateContext.ContextDependencies.Any(x => resultContextsDict[x].Error is not null))
		{
			var firstDependecyContextIdWithError = templateContext.ContextDependencies
				.Where(x => resultContextsDict[x].Error is not null).First();

			var dependencyError = resultContextsDict[firstDependecyContextIdWithError].Error;
			resultItem.ResultContexts.Add(new WvExcelFileResultItemContext
			{
				TemplateContextId = templateContext.Id,
				Error = dependencyError,
				Range = templateContext.Range
			});
			return true;
		}
		return false;
	}

	private WvExcelFileResultItemContext? generateResultRegion(WvExcelFileTemplateContext templateContext,
		WvExcelFileTemplateProcessResultItem resultItem,
		Dictionary<Guid, WvExcelFileResultItemContext> resultContextsDict,
		DataTable dataSource, CultureInfo culture)
	{
		if (templateContext.Type == WvExcelFileTemplateContextType.CellRange)
		{
			return generateResultRegionCellRange(
				templateContext: templateContext,
				resultItem: resultItem,
				resultContextsDict: resultContextsDict,
				dataSource: dataSource,
				culture: culture);
		}
		else if (templateContext.Type == WvExcelFileTemplateContextType.Picture)
		{
			return generateResultRegionCellPicture(
				templateContext: templateContext,
				resultItem: resultItem,
				resultContextsDict: resultContextsDict,
				dataSource: dataSource,
				culture: culture);
		}
		return null;
	}

	private WvExcelFileResultItemContext generateResultRegionCellRange(WvExcelFileTemplateContext templateContext,
		WvExcelFileTemplateProcessResultItem resultItem,
		Dictionary<Guid, WvExcelFileResultItemContext> resultContextsDict,

		DataTable dataSource, CultureInfo culture)
	{
		if (templateContext.Range is null)
			throw new ArgumentException("context has no range", nameof(templateContext));
		if (templateContext.Worksheet is null)
			throw new ArgumentException("context has no Worksheet", nameof(templateContext));
		if (resultItem.Result is null)
			throw new ArgumentException("result item result has no value", nameof(resultItem));

		WvExcelFileResultItemContext result = null;
		var resultWs = resultItem.Result!.Worksheets.Single(x => x.Position == templateContext.Worksheet.Position);

		//Context region should start its expansion based on its context and rules
		var (startRow, startCol) = calculateStartingAddress(
			templateContext: templateContext,
			resultContextsDict: resultContextsDict
		);
		var currentRow = startRow;
		var currentCol = startCol;
		var processedTemplateCells = new HashSet<string>();
		var firstAddress = templateContext.Range.RangeAddress.FirstAddress;
		var lastAddress = templateContext.Range.RangeAddress.LastAddress;

		for (var rowNum = firstAddress.RowNumber;
			rowNum <= lastAddress.RowNumber; rowNum++)
		{
			for (var colNum = firstAddress.ColumnNumber;
				colNum <= lastAddress.ColumnNumber; colNum++)
			{
				IXLCell tempCell = templateContext.Worksheet.Cell(rowNum, colNum);
				//Check and maintain processed List
				if (processedTemplateCells.Contains(tempCell.Address.ToString() ?? ""))
					continue;
				var mergedRange = tempCell.MergedRange();
				var mergedRows = 1;
				var mergedCols = 1;
				if (mergedRange != null)
				{
					foreach (var cell in mergedRange.Cells())
					{
						processedTemplateCells.Add(cell.Address.ToString() ?? "");
					}
					mergedRows = mergedRange.RowCount();
					mergedCols = mergedRange.ColumnCount();
				}
				else
				{
					processedTemplateCells.Add(tempCell.Address.ToString() ?? "");
				}

				var tagProcessResult = WvTemplateUtility.ProcessTemplateTag(tempCell.Value.ToString(), dataSource, culture);
				var isFlowHorizontal = IsFlowHorizontal(tagProcessResult);
				if (tagProcessResult.Tags.Count > 0)
				{
					for (var i = 0; i < tagProcessResult.Values.Count; i++)
					{
						//Vertical
						var endRow = currentRow + (mergedRows - 1);
						var endCol = currentCol + (mergedCols - 1);
						IXLRange resultRange = resultWs.Range(currentRow, currentCol, endRow, endCol);
						resultRange.Value = XLCellValue.FromObject(tagProcessResult.Values[i]);
						if (mergedRows > 1 || mergedCols > 1)
							resultRange.Merge();
						CopyCellProperties(tempCell, resultRange);

						if (isFlowHorizontal)
						{
							currentRow = (i != tagProcessResult.Values.Count - 1 ? currentRow : endRow);
							currentCol = endCol + (i != tagProcessResult.Values.Count - 1 ? 1 : 0);
						}
						else
						{
							currentRow = endRow + (i != tagProcessResult.Values.Count - 1 ? 1 : 0);
							currentCol = (i != tagProcessResult.Values.Count - 1 ? currentCol : endCol);
						}
					}

				}
				else
				{
					IXLRange resultRange = resultWs.Range(currentRow, currentCol, currentRow, currentCol);
					resultRange.Value = tempCell.Value;
					CopyCellProperties(tempCell, resultRange);
					if (isFlowHorizontal)
					{
						currentCol = currentCol + (mergedCols - 1);
					}
					else
					{
						currentRow = currentRow + (mergedRows - 1);
					}
				}
			}
		}

		result = new WvExcelFileResultItemContext
		{
			TemplateContextId = templateContext.Id,
			Range = resultWs.Range(startRow, startCol, currentRow, currentCol),
		};
		CopyColumnsProperties(templateContext.Range, result.Range, templateContext.Worksheet, resultWs);
		CopyRowProperties(templateContext.Range, result.Range, templateContext.Worksheet, resultWs);
		return result;
	}

	private WvExcelFileResultItemContext generateResultRegionCellPicture(WvExcelFileTemplateContext templateContext,
		WvExcelFileTemplateProcessResultItem resultItem,
		Dictionary<Guid, WvExcelFileResultItemContext> resultContextsDict,
		DataTable dataSource, CultureInfo culture)
	{
		//Context region should start its expansion based on its context and rules
		return null;
	}

	private (int, int) calculateStartingAddress(WvExcelFileTemplateContext templateContext,
		Dictionary<Guid, WvExcelFileResultItemContext> resultContextsDict
	)
	{
		int startRow = 1;
		int startCol = 1;
		if (templateContext.ParentContext is not null && templateContext.ParentContext.Range is not null)
		{
			if (!resultContextsDict.ContainsKey(templateContext.ParentContext.Id))
				throw new Exception($"Parent context not found! Range {templateContext.ParentContext.Range.RangeAddress.ToString()}");

			var parentResultContext = resultContextsDict[templateContext.ParentContext.Id];
			if (parentResultContext.Range is null)
				throw new Exception($"Parent context has no result Range! Range {templateContext.ParentContext.Range.RangeAddress.ToString()}");


			//Attach on the topright side of the parent
			if (templateContext.ParentContext.Id == templateContext.LeftContext?.Id)
			{
				startRow = parentResultContext.Range.FirstRow().RowNumber();
				startCol = parentResultContext.Range.LastColumn().ColumnNumber() + 1;
			}
			//Attach on the bottomleft of the parent
			else if (templateContext.ParentContext.Id == templateContext.TopContext?.Id)
			{
				startRow = parentResultContext.Range.LastRow().RowNumber() + 1;
				startCol = parentResultContext.Range.FirstColumn().ColumnNumber();
			}
		}
		return (startRow, startCol);
	}

	private string getCellAddress(int worksheetPosition, int row, int col) => $"{worksheetPosition}:{row}:{col}";

	#endregion
}
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Engines.Excel.Services;
using WebVella.DocumentTemplates.Extensions;

namespace WebVella.DocumentTemplates.Engines.Excel.Utility;
public static partial class WvExcelFileEngineUtility
{
	/// <summary>
	/// Creates result contexts and result Excel by processing the template contexts, fill them with data
	/// and situate them on the result Excel based on their parent Context
	/// </summary>
	/// <param name="result"></param>
	/// <param name="resultItem"></param>
	/// <param name="dataSource"></param>
	/// <param name="culture"></param>
	public static void ProcessExcelTemplateGenerateResultContexts(WvExcelFileTemplateProcessResult result,
	 WvExcelFileTemplateProcessResultItem resultItem, DataTable dataSource, CultureInfo culture,
	 Dictionary<Guid, WvExcelFileTemplateContext> templateContextDict)
	{
		//Validate
		if (resultItem is null) throw new Exception("No result provided!");
		if (dataSource is null) throw new Exception("No datasource provided!");
		if (result.Template is null) throw new Exception("No Template provided!");
		if (result.TemplateContexts is null) throw new Exception("No Template provided!");
		if (resultItem.Result is null) resultItem.Result = new XLWorkbook();

		//Init
		if (result.TemplateContexts.Count == 0) return;
		int processAttemptsLimit = 200;

		createResultWorksheets(
			result: result,
			resultItem: resultItem,
			dataSource: dataSource,
			culture: culture
		);
		var currentResultRowStartRow = 1;
		//Process
		foreach (var resultWorksheet in resultItem.Result.Worksheets)
		{
			#region << Process Cell Ranges >>
			foreach (var templateRow in result.TemplateRows.Where(x =>
				x.Worksheet?.Position == resultWorksheet.Position))
			{
				Queue<Guid> queue = new Queue<Guid>();
				templateRow.Contexts.ForEach(x => queue.Enqueue(x));

				var resultRow = new WvExcelFileTemplateProcessResultItemRow()
				{
					TemplateFirstRow = templateRow.Position,
					Worksheet = resultWorksheet,
					Contexts = new(),
					ResultFirstRow = currentResultRowStartRow,
					RelativeGrid = new()
				};
				while (queue.Count > 0)
				{
					var contextId = queue.Dequeue();
					var templateContext = templateContextDict[contextId];
					if (templateContext.Type != WvExcelFileTemplateContextType.CellRange) continue;
					//Check Limit
					if (processLimitIsReached(
						templateContext: templateContext,
						resultItem: resultItem,
						limit: processAttemptsLimit
					)) continue;

					//Check if unfulfilled Dependency
					if (hasUnfulfilledDependencies(
						templateContext: templateContext,
						resultItem: resultItem,
						queue: queue
					)) continue;

					//Check if there is fulfulled dependency with error
					if (hasfulfilledDependencyWithError(
						templateContext: templateContext,
						resultItem: resultItem
					)) continue;
					try
					{
						AddTemplateCellRangeContextInResult(
							templateContext: templateContext,
							result: result,
							resultItem: resultItem,
							resultRow: resultRow,
							dataSource: dataSource,
							culture: culture
						);
					}
					catch (Exception ex)
					{
						resultItem.ResultContexts.Add(new WvExcelFileTemplateProcessResultItemContext
						{
							TemplateContextId = templateContext.Id,
							Error = WvExcelFileTemplateProcessResultItemContextError.ProcessError,
							ErrorMessage = ex.Message
						});
						AddTemplateCellRangeContextErrorInResult(
							templateContext: templateContext,
							resultItem: resultItem,
							resultRow: resultRow,
							errorType: WvExcelFileTemplateProcessResultItemContextError.ProcessError,
							errorMessage: ex.Message
						);
					}
				}

				if (resultRow.RelativeGrid.Any())
					currentResultRowStartRow += resultRow.RelativeGridMaxRow;

				resultItem.ResultRows.Add(resultRow);
				FillInEmptyResultRowGridItems(
					result: result,
					resultItem: resultItem,
					resultRow: resultRow,
					templateContextDict: templateContextDict
				);
			}
			#endregion

			#region << Process Pictures >>
			foreach (var templateContext in result.TemplateContexts.Where(x =>
				x.Type == WvExcelFileTemplateContextType.Picture
				&& x.Worksheet!.Position == resultWorksheet.Position))
			{
				AddTemplatePictureContextInResult(
					templateContext: templateContext,
					worksheet: resultWorksheet,
					resultItem: resultItem,
					result: result);
			}
		}
		#endregion
	}

	#region << Private >>

	private static void createResultWorksheets(
			WvExcelFileTemplateProcessResult result,
			WvExcelFileTemplateProcessResultItem resultItem,
			DataTable dataSource, CultureInfo culture
		)
	{
		if (result.Template is null) throw new ArgumentException("result.Template not initialized", nameof(result));
		if (resultItem.Result is null) throw new ArgumentException("resultItem.Result not initialized", nameof(resultItem));
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
	}

	private static bool processLimitIsReached(WvExcelFileTemplateContext templateContext,
		WvExcelFileTemplateProcessResultItem resultItem, int limit)
	{

		if (!resultItem.ContextProcessLog.ContainsKey(templateContext.Id))
			resultItem.ContextProcessLog[templateContext.Id] = 0;

		if (resultItem.ContextProcessLog[templateContext.Id] > limit)
		{
			resultItem.ResultContexts.Add(new WvExcelFileTemplateProcessResultItemContext
			{
				TemplateContextId = templateContext.Id,
				Error = WvExcelFileTemplateProcessResultItemContextError.DependencyOverflow,
				Range = templateContext.Range
			});
			return true;
		};
		resultItem.ContextProcessLog[templateContext.Id]++;
		return false;
	}
	private static bool hasUnfulfilledDependencies(WvExcelFileTemplateContext templateContext,
		WvExcelFileTemplateProcessResultItem resultItem,
		Queue<Guid> queue)
	{
		foreach (var contextId in templateContext.ContextDependencies)
		{
			var dependency = resultItem.ResultContexts.FirstOrDefault(x => x.TemplateContextId == contextId);
			if (dependency is null)
			{
				queue.Enqueue(templateContext.Id);
				return true;
			}
		}
		return false;
	}

	private static bool hasfulfilledDependencyWithError(WvExcelFileTemplateContext templateContext,
		WvExcelFileTemplateProcessResultItem resultItem)
	{
		foreach (var contextId in templateContext.ContextDependencies)
		{
			var dependency = resultItem.ResultContexts.FirstOrDefault(x => x.TemplateContextId == contextId);
			if (dependency is not null && dependency.Error is not null)
			{
				resultItem.ResultContexts.Add(new WvExcelFileTemplateProcessResultItemContext
				{
					TemplateContextId = templateContext.Id,
					Error = dependency.Error
				});
				return true;
			}
		}
		return false;
	}

	private static void AddTemplateCellRangeContextInResult(
		WvExcelFileTemplateContext templateContext,
		WvExcelFileTemplateProcessResult result,
		WvExcelFileTemplateProcessResultItem resultItem,
		WvExcelFileTemplateProcessResultItemRow resultRow,
		DataTable dataSource, CultureInfo culture)
	{
		#region << Validation >>
		if (templateContext.Range is null)
			throw new ArgumentException("context has no range", nameof(templateContext));
		if (templateContext.Worksheet is null)
			throw new ArgumentException("context has no Worksheet", nameof(templateContext));
		if (resultRow is null)
			throw new ArgumentException("resultRow has no value", nameof(resultRow));
		#endregion

		#region << Init >>
		//Context should start working from top to bottom, left to right
		//of the available grid space
		var resultStartRow = resultRow.ResultFirstRow;
		var resultStartCol = resultRow.RelativeGridMaxColumn + 1;
		var processedTemplateCells = new HashSet<string>();
		var templateFirstAddress = templateContext.Range.RangeAddress.FirstAddress;
		var templateLastAddress = templateContext.Range.RangeAddress.LastAddress;
		var currentRow = resultStartRow;
		var currentCol = resultStartCol;
		#endregion


		for (var rowNum = templateFirstAddress.RowNumber;
			rowNum <= templateLastAddress.RowNumber; rowNum++)
		{
			for (var colNum = templateFirstAddress.ColumnNumber;
				colNum <= templateLastAddress.ColumnNumber; colNum++)
			{
				IXLCell tempCell = templateContext.Worksheet.Cell(rowNum, colNum);
				//Check and maintain processed List
				if (processedTemplateCells.Contains(tempCell.Address.ToString() ?? "")) continue;
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
				if (tagProcessResult.Tags.Any(x => x.Type == Core.WvTemplateTagType.ExcelFunction))
				{
					IXLRange resultRange = resultRow.Worksheet!.Range(currentRow, currentCol, currentRow, currentCol);
					//Excel Functions Can have only one result				
					if (tagProcessResult.Tags.Count == 1
						&& tagProcessResult.Tags[0].Type == WvTemplateTagType.ExcelFunction)
					{
						var tag = tagProcessResult.Tags[0];
						var functionProcessor = WvExcelFileTemplateService.GetExcelFunctionProcessorByName(tag.FunctionName ?? String.Empty);
						if (functionProcessor is null)
						{
							throw new Exception($"Unsupported function name: {tag.Name} in tag");
						}
						functionProcessor.Process(
							tag: tagProcessResult.Tags[0],
							dataSource: dataSource,
							result: result,
							resultItem: resultItem,
							worksheet: resultRow.Worksheet!
						);
						if (functionProcessor.HasError)
						{
							resultRange.Value = XLError.IncompatibleValue;
							IXLCell resultCell = resultRow.Worksheet!.Cell(currentRow, currentCol);
							var comment = resultCell.CreateComment();
							comment.AddText(functionProcessor.ErrorMessage);
						}
						else
						{
							resultRange.FormulaA1 = functionProcessor.FormulaA1;
						}

					}
					addRangeToGrid(
									firstRow: currentRow,
									firstColumn: currentCol,
									lastRow: currentRow,
									lastColumn: currentCol,
									contextId: templateContext.Id,
									resultRow: resultRow,
									resultFirstRow: resultRow.ResultFirstRow);
					CopyCellProperties(tempCell, resultRange);
					currentCol = currentCol + (mergedCols - 1);
				}
				else
				{
					var isFlowHorizontal = IsFlowHorizontal(tagProcessResult);
					if (tagProcessResult.Tags.Count > 0)
					{
						//Postprocess results if function
						foreach (var tag in tagProcessResult.Tags)
						{
							if (tag.Type != WvTemplateTagType.Function) continue;
							var functionProcessor = WvExcelFileTemplateService.GetFunctionProcessorByName(tag.FunctionName ?? String.Empty);
							if (functionProcessor is null)
							{
								throw new Exception($"Unsupported function name: {tag.Name} in tag");
							}
							tagProcessResult = functionProcessor.Process(
														tag: tag,
														input: tagProcessResult,
														dataSource: dataSource,
														result: result,
														resultItem: resultItem,
														worksheet: resultRow.Worksheet!
													);
						}
						for (var i = 0; i < tagProcessResult.Values.Count; i++)
						{
							//Vertical
							var endRow = currentRow + (mergedRows - 1);
							var endCol = currentCol + (mergedCols - 1);
							IXLRange resultRange = resultRow.Worksheet!.Range(currentRow, currentCol, endRow, endCol);

							resultRange.Value = XLCellValue.FromObject(tagProcessResult.Values[i]);

							if (mergedRows > 1 || mergedCols > 1)
								resultRange.Merge();
							CopyCellProperties(tempCell, resultRange);
							addRangeToGrid(
								firstRow: currentRow,
								firstColumn: currentCol,
								lastRow: endRow,
								lastColumn: endCol,
								contextId: templateContext.Id,
								resultRow: resultRow,
								resultFirstRow: resultRow.ResultFirstRow);

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
						IXLRange resultRange = resultRow.Worksheet!.Range(currentRow, currentCol, currentRow, currentCol);
						resultRange.Value = tempCell.Value;
						addRangeToGrid(
							firstRow: currentRow,
							firstColumn: currentCol,
							lastRow: currentRow,
							lastColumn: currentCol,
							contextId: templateContext.Id,
							resultRow: resultRow,
							resultFirstRow: resultRow.ResultFirstRow);
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
		}

		var resultItemContext = new WvExcelFileTemplateProcessResultItemContext
		{
			TemplateContextId = templateContext.Id,
			Range = resultRow.Worksheet!.Range(resultStartRow, resultStartCol, currentRow, currentCol),

		};
		resultItem.ResultContexts.Add(resultItemContext);

		CopyColumnsProperties(templateContext.Range, resultItemContext.Range, templateContext.Worksheet, resultRow.Worksheet);
		CopyRowProperties(templateContext.Range, resultItemContext.Range, templateContext.Worksheet, resultRow.Worksheet);
	}

	private static void AddTemplatePictureContextInResult(
		WvExcelFileTemplateContext templateContext,
		IXLWorksheet worksheet,
		WvExcelFileTemplateProcessResultItem resultItem,
		WvExcelFileTemplateProcessResult result)
	{
		//Picture should be placed relative to the top left corner of the context
		if (templateContext is null) return;
		if (templateContext.Picture is null) return;
		var pictureTopLeftRow = templateContext.Picture.TopLeftCell.Address.RowNumber;
		var pictureTopLeftColumn = templateContext.Picture.TopLeftCell.Address.ColumnNumber;
		var anchorTemplateContext = result.TemplateContexts.GetByAddress(
			worksheetPosition: worksheet.Position,
			row: pictureTopLeftRow,
			column: pictureTopLeftColumn).FirstOrDefault();

		if (anchorTemplateContext is null) return;

		var templateResultContexts = resultItem.ResultContexts.Where(x =>
			x.TemplateContextId == anchorTemplateContext.Id).ToList();

		var resultRow = templateResultContexts.Min(x => x.Range!.RangeAddress.FirstAddress.RowNumber);
		var resultColumn = templateResultContexts.Min(x => x.Range!.RangeAddress.FirstAddress.ColumnNumber);
		AddPictureToWorksheet(
		worksheet: worksheet,
		picture: templateContext.Picture,
		row: resultRow,
		column: resultColumn
		);

	}

	private static void AddTemplateCellRangeContextErrorInResult(
		WvExcelFileTemplateContext templateContext,
		WvExcelFileTemplateProcessResultItem resultItem,
		WvExcelFileTemplateProcessResultItemRow resultRow,
		WvExcelFileTemplateProcessResultItemContextError errorType,
		string errorMessage)
	{
		#region << Validation >>
		if (templateContext.Range is null)
			throw new ArgumentException("context has no range", nameof(templateContext));
		if (templateContext.Worksheet is null)
			throw new ArgumentException("context has no Worksheet", nameof(templateContext));
		if (resultRow is null)
			throw new ArgumentException("resultRow has no value", nameof(resultRow));
		#endregion

		#region << Init >>
		//Context should start working from top to bottom, left to right
		//of the available grid space
		var resultStartRow = resultRow.ResultFirstRow;
		var resultStartCol = resultRow.RelativeGridMaxColumn + 1;
		var processedTemplateCells = new HashSet<string>();
		var templateFirstAddress = templateContext.Range.RangeAddress.FirstAddress;
		var templateLastAddress = templateContext.Range.RangeAddress.LastAddress;
		var currentRow = resultStartRow;
		var currentCol = resultStartCol;
		#endregion


		for (var rowNum = templateFirstAddress.RowNumber;
			rowNum <= templateLastAddress.RowNumber; rowNum++)
		{
			for (var colNum = templateFirstAddress.ColumnNumber;
				colNum <= templateLastAddress.ColumnNumber; colNum++)
			{
				IXLCell tempCell = templateContext.Worksheet.Cell(rowNum, colNum);
				//Check and maintain processed List
				if (processedTemplateCells.Contains(tempCell.Address.ToString() ?? "")) continue;
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

				IXLRange resultRange = resultRow.Worksheet!.Range(currentRow, currentCol, currentRow, currentCol);
				IXLCell resultCell = resultRow.Worksheet!.Cell(currentRow, currentCol);
				resultRange.Value = XLError.IncompatibleValue;
				var comment = resultCell.CreateComment();
				comment.AddText(errorMessage);
				addRangeToGrid(
					firstRow: currentRow,
					firstColumn: currentCol,
					lastRow: currentRow,
					lastColumn: currentCol,
					contextId: templateContext.Id,
					resultRow: resultRow,
					resultFirstRow: resultRow.ResultFirstRow);
				CopyCellProperties(tempCell, resultRange);
				currentRow = currentRow + (mergedRows - 1);

			}
		}

		var resultItemContext = new WvExcelFileTemplateProcessResultItemContext
		{
			TemplateContextId = templateContext.Id,
			Range = resultRow.Worksheet!.Range(resultStartRow, resultStartCol, currentRow, currentCol),

		};
		resultItem.ResultContexts.Add(resultItemContext);

		CopyColumnsProperties(templateContext.Range, resultItemContext.Range, templateContext.Worksheet, resultRow.Worksheet);
		CopyRowProperties(templateContext.Range, resultItemContext.Range, templateContext.Worksheet, resultRow.Worksheet);
	}

	private static void FillInEmptyResultRowGridItems(
	WvExcelFileTemplateProcessResult result,
	WvExcelFileTemplateProcessResultItem resultItem,
	WvExcelFileTemplateProcessResultItemRow resultRow,
	Dictionary<Guid, WvExcelFileTemplateContext> templateContextDict
	)
	{
		if (!resultRow.RelativeGrid.Any()) return;

		var firstRow = 1;
		var firstColumn = 1;
		var lastRow = resultRow.RelativeGridMaxRow;
		var lastColumn = resultRow.RelativeGridMaxColumn;
		var processedTemplateCells = new HashSet<string>();
		for (var rowNum = firstRow; rowNum <= lastRow; rowNum++)
		{
			for (var columnNum = firstColumn; columnNum <= lastColumn; columnNum++)
			{
				if (resultRow.RelativeGridAddressHs.Contains($"{rowNum}:{columnNum}")) continue;
				if (processedTemplateCells.Contains($"{rowNum}-{columnNum}")) continue;

				var resultRowNumber = rowNum + resultRow.ResultFirstRow - 1;
				//the responsible context can be found by the context of any cell in the same column
				var sameColumnCell = resultRow.RelativeGrid.FirstOrDefault(x => x.Column == columnNum);
				if (sameColumnCell is null) continue;
				var templateContext = templateContextDict[sameColumnCell.ContextId];
				var tempCellRow = templateContext.Range!.RangeAddress.FirstAddress.RowNumber;
				var tempCellColumn = templateContext.Range.RangeAddress.FirstAddress.ColumnNumber;
				var tempCell = templateContext.Worksheet!.Cell(tempCellRow, tempCellColumn);
				var mergedRange = tempCell.MergedRange();

				var mergedRows = 1;
				var mergedCols = 1;
				if (mergedRange != null)
				{
					mergedRows = mergedRange.RowCount();
					mergedCols = mergedRange.ColumnCount();

					for (var i = 0; i < mergedRows; i++)
					{
						for (var j = 0; j < mergedCols; j++)
						{
							processedTemplateCells.Add($"{rowNum + i}-{columnNum + j}");
						}
					}
				}
				else
				{
					processedTemplateCells.Add($"{rowNum}-{columnNum}");
				}

				var endRow = resultRowNumber + (mergedRows - 1);
				var endCol = columnNum + (mergedCols - 1);
				IXLRange resultRange = resultRow.Worksheet!.Range(resultRowNumber, columnNum, endRow, endCol);
				if (mergedRows > 1 || mergedCols > 1)
					resultRange.Merge();
				addRangeToGrid(
					firstRow: resultRowNumber,
					firstColumn: columnNum,
					lastRow: endRow,
					lastColumn: endCol,
					contextId: templateContext.Id,
					resultRow: resultRow,
					resultFirstRow: resultRow.ResultFirstRow);
				CopyCellProperties(tempCell, resultRange);
			}
		}
	}

	private static void AddPictureToWorksheet(IXLWorksheet worksheet, IXLPicture picture, int row, int column)
	{
		var newPicture = picture.CopyTo(worksheet);
		var anchorCell = worksheet.Cell(row, column);
		newPicture.MoveTo(anchorCell, picture.Left, picture.Top);
	}

	#endregion
}

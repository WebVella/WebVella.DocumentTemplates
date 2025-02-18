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
public partial class WvExcelFileEngineUtility
{
	/// <summary>
	/// Creates result contexts and result Excel by processing the template contexts, fill them with data
	/// and situate them on the result Excel based on their parent Context
	/// </summary>
	/// <param name="result"></param>
	/// <param name="resultItem"></param>
	/// <param name="dataSource"></param>
	/// <param name="culture"></param>
	public void ProcessExcelTemplateGenerateResultContexts(WvExcelFileTemplateProcessResult result,
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

	private void createResultWorksheets(
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
			var templateResult = new WvTemplateUtility().ProcessTemplateTag(templateWs.Name, firstRowDt, culture);
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

	private bool processLimitIsReached(WvExcelFileTemplateContext templateContext,
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
	private bool hasUnfulfilledDependencies(WvExcelFileTemplateContext templateContext,
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

	private bool hasfulfilledDependencyWithError(WvExcelFileTemplateContext templateContext,
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

	private void AddTemplateCellRangeContextInResult(
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

		var processedResultRows = new HashSet<long>();
		var processedResultColumns = new HashSet<long>();
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
				var endRow = 1;
				var endCol = 1;
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


				var tagProcessResult1 = new WvExcelFileEngineUtility().ProcessTemplateTag(
					template: tempCell.Value.ToString(),
					dataSource: dataSource,
					templateContext: templateContext,
					resultItem: resultItem);

				for (int expandIndex = 0; expandIndex < tagProcessResult1.ExpandCount; expandIndex++)
				{
					endRow = currentRow + (mergedRows - 1);
					endCol = currentCol + (mergedCols - 1);
					IXLRange resultRange = resultRow.Worksheet!.Range(currentRow, currentCol, endRow, endCol);
					if (mergedRange != null) resultRange.Merge();
					if (tempCell.Value.IsText)
					{
						var tags = new WvTemplateUtility().GetTagsFromTemplate(tempCell.Value.ToString());
						if (tags.Count == 0 || String.IsNullOrWhiteSpace(tempCell.Value.ToString()))
						{
							resultRange.Value = tempCell.Value;
						}
						else if (tags.Count == 1 && tempCell.Value.ToString() == tags[0].FullString)
						{
							var tag = tags[0];

							//Only the tag in the template - common for functions
							if (tag.Type == WvTemplateTagType.Data)
							{
								var value = new WvTemplateUtility().GetTemplateValue(
									template: tempCell.Value.ToString(),
									dataRowIndex: expandIndex,
									dataSource: dataSource,
									culture: culture
								);
								resultRange.Value = XLCellValue.FromObject(value);
							}
							else if (tag.Type == WvTemplateTagType.Function)
							{
								var functionProcessor = new WvExcelFileMetaService().GetFunctionProcessorByName(tag.FunctionName ?? String.Empty);
								if (functionProcessor is null)
								{
									throw new Exception($"Unsupported function name: {tag.Name} in tag");
								}

								//Need to work here to make it relative
								var value = functionProcessor.Process(
													tagValue: tempCell.Value.ToString(),
													tag: tag,
													dataSource: dataSource,
													result: result,
													resultItem: resultItem,
													processedCellRange: resultRange,
													processedWorksheet: resultRow.Worksheet
												);
								if (functionProcessor.HasError) throw new Exception(functionProcessor.ErrorMessage);
								resultRange.Value = XLCellValue.FromObject(value);
							}
							else if (tag.Type == WvTemplateTagType.ExcelFunction)
							{
								var functionProcessor = new WvExcelFileMetaService().GetExcelFunctionProcessorByName(tag.FunctionName ?? String.Empty);
								if (functionProcessor is null)
								{
									throw new Exception($"Unsupported function name: {tag.Name} in tag");
								}

								//TODO: BOZ Need to work here to make it relative
								var value = functionProcessor.Process(
													value: tempCell.Value.ToString(),
													tag: tag,
													dataSource: dataSource,
													result: result,
													resultItem: resultItem,
													processedCellRange: resultRange,
													processedWorksheet: resultRow.Worksheet
												);
								if (functionProcessor.HasError) throw new Exception(functionProcessor.ErrorMessage);
								resultRange.FormulaA1 = functionProcessor.FormulaA1;
							}
						}
						else
						{
							object? value = tempCell.Value.ToString();
							foreach (var tag in tags)
							{
								if (tag.Type == WvTemplateTagType.Data)
								{
									value = new WvTemplateUtility().GetTemplateValue(
										template: value!.ToString(),
										dataRowIndex: expandIndex,
										dataSource: dataSource,
										culture: culture);
								}
								else if (tag.Type == WvTemplateTagType.Function)
								{
									var functionProcessor = new WvExcelFileMetaService().GetFunctionProcessorByName(tag.FunctionName ?? String.Empty);
									if (functionProcessor is null)
									{
										throw new Exception($"Unsupported function name: {tag.Name} in tag");
									}
									//TODO: BOZ Need to work here to make it relative
									value = functionProcessor.Process(
														tagValue: value.ToString(),
														tag: tag,
														dataSource: dataSource,
														result: result,
														resultItem: resultItem,
														processedCellRange: resultRange,
														processedWorksheet: resultRow.Worksheet
													);
									if (functionProcessor.HasError) throw new Exception(functionProcessor.ErrorMessage);
								}
								else if (tag.Type == WvTemplateTagType.ExcelFunction)
								{
									//Need to work here to make it relative
								}
								if (value is not string) break;
							}


							resultRange.Value = XLCellValue.FromObject(value);
						}
					}
					else
					{
						resultRange.Value = tempCell.Value;
					}
					if (templateContext.Flow == WvTemplateTagDataFlow.Vertical)
					{
						currentRow = endRow + 1;
					}
					else
					{
						currentCol = endCol + 1;
					}
				}

				var resultItemContext = new WvExcelFileTemplateProcessResultItemContext
				{
					TemplateContextId = templateContext.Id,
					Range = resultRow.Worksheet!.Range(resultStartRow, resultStartCol, endRow, endCol),
				};
				resultItem.ResultContexts.Add(resultItemContext);
				addRangeToGrid(
					firstRow: resultStartRow,
					firstColumn: resultStartCol,
					lastRow: endRow,
					lastColumn: endCol,
					contextId: templateContext.Id,
					resultRow: resultRow,
					resultFirstRow: resultRow.ResultFirstRow);
				CopyCellProperties(tempCell, resultItemContext.Range);
				CopyColumnsProperties(templateContext.Range, resultItemContext.Range, templateContext.Worksheet, resultRow.Worksheet, processedResultColumns);
				CopyRowProperties(templateContext.Range, resultItemContext.Range, templateContext.Worksheet, resultRow.Worksheet, processedResultRows);
			}
		}
	}

	private void AddTemplatePictureContextInResult(
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

	private void AddTemplateCellRangeContextErrorInResult(
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
				currentRow = currentRow + (mergedRows - 1);

			}
		}

		var resultItemContext = new WvExcelFileTemplateProcessResultItemContext
		{
			TemplateContextId = templateContext.Id,
			Range = resultRow.Worksheet!.Range(resultStartRow, resultStartCol, currentRow, currentCol),

		};
		resultItem.ResultContexts.Add(resultItemContext);
		addRangeToGrid(
			firstRow: currentRow,
			firstColumn: currentCol,
			lastRow: currentRow,
			lastColumn: currentCol,
			contextId: templateContext.Id,
			resultRow: resultRow,
			resultFirstRow: resultRow.ResultFirstRow);
		IXLCell firstTempCell = templateContext.Worksheet.Cell(templateFirstAddress.RowNumber, templateFirstAddress.ColumnNumber);
		CopyCellProperties(firstTempCell, resultItemContext.Range);
		CopyColumnsProperties(templateContext.Range, resultItemContext.Range, templateContext.Worksheet, resultRow.Worksheet, new HashSet<long>());
		CopyRowProperties(templateContext.Range, resultItemContext.Range, templateContext.Worksheet, resultRow.Worksheet, new HashSet<long>());
	}

	private void FillInEmptyResultRowGridItems(
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
	private void AddPictureToWorksheet(IXLWorksheet worksheet, IXLPicture picture, int row, int column)
	{
		var newPicture = picture.CopyTo(worksheet);
		var anchorCell = worksheet.Cell(row, column);
		newPicture.MoveTo(anchorCell, picture.Left, picture.Top);
	}

	#endregion
}

using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Engines.SpreadsheetFile.Models;
using WebVella.DocumentTemplates.Engines.SpreadsheetFile.Services;
using WebVella.DocumentTemplates.Extensions;

namespace WebVella.DocumentTemplates.Engines.SpreadsheetFile.Utility;
public partial class WvSpreadsheetFileEngineUtility
{
	/// <summary>
	/// Creates result contexts and result Spreadsheet by processing the template contexts, fill them with data
	/// and situate them on the result Spreadsheet based on their parent Context
	/// </summary>
	/// <param name="result"></param>
	/// <param name="resultItem"></param>
	/// <param name="dataSource"></param>
	/// <param name="culture"></param>
	public void ProcessSpreadsheetTemplateGenerateResultContexts(WvSpreadsheetFileTemplateProcessResult result,
	 WvSpreadsheetFileTemplateProcessResultItem resultItem, DataTable dataSource, CultureInfo culture,
	 Dictionary<Guid, WvSpreadsheetFileTemplateContext> templateContextDict)
	{
		//Validate
		if (resultItem is null) throw new Exception("No result provided!");
		if (dataSource is null) throw new Exception("No datasource provided!");
		if (result.Template is null) throw new Exception("No Template provided!");
		if (result.TemplateContexts is null) throw new Exception("No Template provided!");
		if (resultItem.Workbook is null) resultItem.Workbook = new XLWorkbook();

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
		foreach (var resultWorksheet in resultItem.Workbook.Worksheets)
		{
			#region << Process Cell Ranges >>
			foreach (var templateRow in result.TemplateRows.Where(x =>
				x.Worksheet?.Position == resultWorksheet.Position))
			{
				Queue<Guid> queue = new Queue<Guid>();
				templateRow.Contexts.ForEach(x => queue.Enqueue(x));

				var resultRow = new WvSpreadsheetFileTemplateProcessResultItemRow()
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
					if (templateContext.Type != WvSpreadsheetFileTemplateContextType.CellRange) continue;
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
						resultItem.Contexts.Add(new WvSpreadsheetFileTemplateProcessContext
						{
							TemplateContextId = templateContext.Id,
							SpreadsheetCellError = WvSpreadsheetFileTemplateProcessResultItemContextError.ProcessError,
							SpreadsheetCellErrorMessage = ex.Message,
							Errors = new List<string> { ex.Message }
						});
						AddTemplateCellRangeContextErrorInResult(
							templateContext: templateContext,
							resultItem: resultItem,
							resultRow: resultRow,
							errorType: WvSpreadsheetFileTemplateProcessResultItemContextError.ProcessError,
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
				x.Type == WvSpreadsheetFileTemplateContextType.Picture
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
			WvSpreadsheetFileTemplateProcessResult result,
			WvSpreadsheetFileTemplateProcessResultItem resultItem,
			DataTable dataSource, CultureInfo culture
		)
	{
		if (result.Workbook is null) throw new ArgumentException("result.Template not initialized", nameof(result));
		if (resultItem.Workbook is null) throw new ArgumentException("resultItem.Result not initialized", nameof(resultItem));
		if (result.Workbook.Worksheets.Count == 0) return;

		var firstRowDt = dataSource.CreateAsNew(new List<int> { 0 });
		foreach (var templateWs in result.Workbook.Worksheets.OrderBy(x => x.Position))
		{
			var resultWs = resultItem.Workbook.AddWorksheet();
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

	private bool processLimitIsReached(WvSpreadsheetFileTemplateContext templateContext,
		WvSpreadsheetFileTemplateProcessResultItem resultItem, int limit)
	{

		if (!resultItem.ContextProcessLog.ContainsKey(templateContext.Id))
			resultItem.ContextProcessLog[templateContext.Id] = 0;

		if (resultItem.ContextProcessLog[templateContext.Id] > limit)
		{
			resultItem.Contexts.Add(new WvSpreadsheetFileTemplateProcessContext
			{
				TemplateContextId = templateContext.Id,
				SpreadsheetCellError = WvSpreadsheetFileTemplateProcessResultItemContextError.DependencyOverflow,
				Range = templateContext.Range
			});
			return true;
		};
		resultItem.ContextProcessLog[templateContext.Id]++;
		return false;
	}
	private bool hasUnfulfilledDependencies(WvSpreadsheetFileTemplateContext templateContext,
		WvSpreadsheetFileTemplateProcessResultItem resultItem,
		Queue<Guid> queue)
	{
		foreach (var contextId in templateContext.ContextDependencies)
		{
			var dependency = resultItem.Contexts.FirstOrDefault(x => x.TemplateContextId == contextId);
			if (dependency is null)
			{
				queue.Enqueue(templateContext.Id);
				return true;
			}
		}
		return false;
	}

	private bool hasfulfilledDependencyWithError(WvSpreadsheetFileTemplateContext templateContext,
		WvSpreadsheetFileTemplateProcessResultItem resultItem)
	{
		foreach (var contextId in templateContext.ContextDependencies)
		{
			var dependency = resultItem.Contexts.FirstOrDefault(x => x.TemplateContextId == contextId);
			if (dependency is not null && dependency.SpreadsheetCellError is not null)
			{
				resultItem.Contexts.Add(new WvSpreadsheetFileTemplateProcessContext
				{
					TemplateContextId = templateContext.Id,
					SpreadsheetCellError = dependency.SpreadsheetCellError
				});
				return true;
			}
		}
		return false;
	}

	private void AddTemplateCellRangeContextInResult(
		WvSpreadsheetFileTemplateContext templateContext,
		WvSpreadsheetFileTemplateProcessResult result,
		WvSpreadsheetFileTemplateProcessResultItem resultItem,
		WvSpreadsheetFileTemplateProcessResultItemRow resultRow,
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


				var tagProcessResult = new WvSpreadsheetFileEngineUtility().ProcessTemplateTag(
					template: tempCell.Value.ToString(),
					dataSource: dataSource,
					templateContext: templateContext,
					resultItem: resultItem);

				for (int expandPosition = 1; expandPosition <= tagProcessResult.ExpandCount; expandPosition++)
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
									dataRowPosition: expandPosition,
									dataSource: dataSource,
									culture: culture
								);
								resultRange.Value = XLCellValue.FromObject(value);
							}
							else if (tag.Type == WvTemplateTagType.Function)
							{
								var processorType = new WvSpreadsheetFileMetaService().GetFunctionProcessorByName(tag.ItemName ?? String.Empty);
								if (processorType is null)
								{
									throw new Exception($"Unsupported function name: {tag.Operator} in tag");
								}
								IWvSpreadsheetFileTemplateFunctionProcessor? processor = Activator.CreateInstance(processorType) as IWvSpreadsheetFileTemplateFunctionProcessor;
								if (processor is null)
								{
									throw new Exception($"Failed to init processor of type: {processorType.FullName} for function name: {tag.Operator} in tag");
								}
								//Need to work here to make it relative
								var value = processor.Process(
													value: tempCell.Value.ToString(),
													tag: tag,
													templateContext: templateContext,
													expandPosition: expandPosition,
													expandPositionMax: tagProcessResult.ExpandCount,
													dataSource: dataSource,
													result: result,
													resultItem: resultItem,
													processedCellRange: resultRange,
													processedWorksheet: resultRow.Worksheet
												);
								if (processor.HasError) throw new Exception(processor.ErrorMessage);
								resultRange.Value = XLCellValue.FromObject(value);
							}
							else if (tag.Type == WvTemplateTagType.SpreadsheetFunction)
							{
								var processorType = new WvSpreadsheetFileMetaService().GetSpreadsheetFunctionProcessorByName(tag.ItemName ?? String.Empty);
								if (processorType is null)
								{
									throw new Exception($"Unsupported function name: {tag.Operator} in tag");
								}
								IWvSpreadsheetFileTemplateSpreadsheetFunctionProcessor? processor = Activator.CreateInstance(processorType) as IWvSpreadsheetFileTemplateSpreadsheetFunctionProcessor;
								if (processor is null)
								{
									throw new Exception($"Failed to init processor of type: {processorType.FullName} for function name: {tag.Operator} in tag");
								}
								var value = processor.Process(
													value: tempCell.Value.ToString(),
													tag: tag,
													expandPosition: expandPosition,
													expandPositionMax: tagProcessResult.ExpandCount,
													templateContext: templateContext,
													result: result,
													resultItem: resultItem,
													processedWorksheet: resultRow.Worksheet
												);
								if (processor.HasError) throw new Exception(processor.ErrorMessage);
								resultRange.FormulaA1 = processor.FormulaA1;
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
										dataRowPosition: expandPosition,
										dataSource: dataSource,
										culture: culture);
								}
								else if (tag.Type == WvTemplateTagType.Function)
								{
									var processorType = new WvSpreadsheetFileMetaService().GetFunctionProcessorByName(tag.ItemName ?? String.Empty);
									if (processorType is null)
									{
										throw new Exception($"Unsupported function name: {tag.Operator} in tag");
									}
									IWvSpreadsheetFileTemplateFunctionProcessor? processor = Activator.CreateInstance(processorType) as IWvSpreadsheetFileTemplateFunctionProcessor;
									if (processor is null)
									{
										throw new Exception($"Failed to init processor of type: {processorType.FullName} for function name: {tag.Operator} in tag");
									}
									value = processor.Process(
														value: value.ToString(),
														tag: tag,
														expandPosition: expandPosition,
														expandPositionMax: tagProcessResult.ExpandCount,
														templateContext: templateContext,
														dataSource: dataSource,
														result: result,
														resultItem: resultItem,
														processedCellRange: resultRange,
														processedWorksheet: resultRow.Worksheet
													);
									if (processor.HasError) throw new Exception(processor.ErrorMessage);
								}
								else if (tag.Type == WvTemplateTagType.SpreadsheetFunction)
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

				var resultItemContext = new WvSpreadsheetFileTemplateProcessContext
				{
					TemplateContextId = templateContext.Id,
					Range = resultRow.Worksheet!.Range(resultStartRow, resultStartCol, endRow, endCol),
					ExpandCount = tagProcessResult.ExpandCount
				};
				resultItem.Contexts.Add(resultItemContext);
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
		WvSpreadsheetFileTemplateContext templateContext,
		IXLWorksheet worksheet,
		WvSpreadsheetFileTemplateProcessResultItem resultItem,
		WvSpreadsheetFileTemplateProcessResult result)
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

		var templateResultContexts = resultItem.Contexts.Where(x =>
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
		WvSpreadsheetFileTemplateContext templateContext,
		WvSpreadsheetFileTemplateProcessResultItem resultItem,
		WvSpreadsheetFileTemplateProcessResultItemRow resultRow,
		WvSpreadsheetFileTemplateProcessResultItemContextError errorType,
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

		var resultItemContext = new WvSpreadsheetFileTemplateProcessContext
		{
			TemplateContextId = templateContext.Id,
			Range = resultRow.Worksheet!.Range(resultStartRow, resultStartCol, currentRow, currentCol),

		};
		resultItem.Contexts.Add(resultItemContext);
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
	WvSpreadsheetFileTemplateProcessResult result,
	WvSpreadsheetFileTemplateProcessResultItem resultItem,
	WvSpreadsheetFileTemplateProcessResultItemRow resultRow,
	Dictionary<Guid, WvSpreadsheetFileTemplateContext> templateContextDict
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

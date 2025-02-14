using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;

namespace WebVella.DocumentTemplates.Engines.Excel.Utility;
public static partial class WvExcelFileEngineUtility
{
	/// <summary>
	/// Parses the Template excel and generates contexts
	/// </summary>
	/// <param name="result"></param>
	/// <exception cref="ArgumentException"></exception>
	public static void ProcessExcelTemplateInitTemplateContexts(WvExcelFileTemplateProcessResult? result)
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

					var cellFlow = WvTemplateTagDataFlow.Vertical;
					var cellTags = WvTemplateUtility.GetTagsFromTemplate(cell.Value.ToString());
					if (cellTags.Count > 0 && !cellTags.Any(x => x.Flow == WvTemplateTagDataFlow.Vertical))
					{
						cellFlow = WvTemplateTagDataFlow.Horizontal;
					}

					//Create context
					var context = new WvExcelFileTemplateContext()
					{
						Id = Guid.NewGuid(),
						Worksheet = ws,
						Row = templateRow,
						Range = mergedRange is not null ? mergedRange : cell.AsRange(),
						Type = WvExcelFileTemplateContextType.CellRange,
						Flow = cellFlow,
						Picture = null,
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
					Flow = WvTemplateTagDataFlow.Vertical,
					Type = WvExcelFileTemplateContextType.Picture,
					ContextDependencies = new()
				};
				result.TemplateContexts.Add(context);
			}
		}
	}
}

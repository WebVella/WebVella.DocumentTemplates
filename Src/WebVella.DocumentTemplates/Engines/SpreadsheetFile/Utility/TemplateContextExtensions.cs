using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Engines.SpreadsheetFile.Models;

namespace WebVella.DocumentTemplates.Engines.SpreadsheetFile.Utility;
public static class TemplateContextExtensions
{
	public static List<WvSpreadsheetFileTemplateContext> GetByAddress(
		this List<WvSpreadsheetFileTemplateContext> contextList,
		int worksheetPosition,
		int row,
		int column,
		WvSpreadsheetFileTemplateContextType? type = null)
	{
		var result = new List<WvSpreadsheetFileTemplateContext>();
		if (row < 1 || column < 1) return result;
		foreach (var context in contextList)
		{
			if (type is not null && context.Type != type.Value) continue;
			if (context.Worksheet is null || context.Worksheet.Position != worksheetPosition) continue;

			if (context.Range is null) continue;

			var firstRow = context.Range!.FirstRow().RowNumber();
			var lastRow = context.Range!.LastRow().RowNumber();

			var firstColumn = context.Range!.FirstColumn().ColumnNumber();
			var lastColumn = context.Range!.LastColumn().ColumnNumber();
			if (firstRow > row || lastRow < row) continue;
			if (firstColumn > column || lastColumn < column) continue;
			result.Add(context);
		}
		return result;
	}
	public static List<WvSpreadsheetFileTemplateContext> GetIntersections(
		this List<WvSpreadsheetFileTemplateContext> contextList,
		int worksheetPosition,
		WvSpreadsheetRange range,
		WvSpreadsheetFileTemplateContextType? type = null)
	{
		if (range.FirstRow <= 0 || range.LastRow <= 0
			|| range.FirstColumn <= 0 || range.LastColumn <= 0)
			throw new ArgumentException("range positions should be > 0", nameof(range));

		if (range.FirstRow > range.LastRow)
			throw new ArgumentException("first row should be <= than last row", nameof(range));
		if (range.FirstColumn > range.LastColumn)
			throw new ArgumentException("first column should be <= than last row", nameof(range));

		var result = new List<WvSpreadsheetFileTemplateContext>();

		foreach (var context in contextList)
		{
			if (type is not null && context.Type != type.Value) continue;
			if (context.Worksheet is null || context.Worksheet.Position != worksheetPosition) continue;

			if (context.Range is null) continue;

			var contextRange = new WvSpreadsheetRange()
			{
				FirstRow = context.Range!.FirstRow().RowNumber(),
				LastRow = context.Range!.LastRow().RowNumber(),
				FirstColumn = context.Range!.FirstColumn().ColumnNumber(),
				LastColumn = context.Range!.LastColumn().ColumnNumber()
			};
			if (range.CheckIntersection(contextRange))
				result.Add(context);
		}
		return result.Where(x => type == null || x.Type == type.Value).ToList();
	}

	public static void CalculateDependencies(this WvSpreadsheetFileTemplateContext context, List<WvSpreadsheetFileTemplateContext> contextList)
	{
		if (context is null) throw new ArgumentNullException(nameof(context));
		if (contextList is null || contextList.Count == 0) return;
		if (context.Worksheet is null) return;
		if (context.Range is null) return;
		var rangeValues = context.Range.CellsUsed().Select(x => x.Value).ToList();
		if (rangeValues.Count == 0) return;
		foreach (var value in rangeValues)
		{
			var tags = new WvTemplateUtility().GetTagsFromTemplate(value.ToString());
			foreach (var tag in tags)
			{
				if (tag.Type != Core.WvTemplateTagType.SpreadsheetFunction
					&& tag.Type != Core.WvTemplateTagType.Function) continue;
				if (String.IsNullOrWhiteSpace(tag.ItemName)) continue;
				var valueAddressList = tag.GetApplicableRangeForFunctionTag();
				if (valueAddressList is null || valueAddressList.Count == 0) continue;
				foreach (var address in valueAddressList)
				{
					var intersects = contextList
						.GetIntersections(context.Worksheet!.Position, address)
						.Where(x => x.Id != context.Id).ToList();
					foreach (var interContext in intersects)
					{
						if (context.ContextDependencies.Contains(interContext.Id)) continue;
						context.ContextDependencies.Add(interContext.Id);
					}

				}
			}
		}
	}
	private static bool CheckIntersection(this WvSpreadsheetRange range1, WvSpreadsheetRange range2)
	{
		// Check if the ranges intersect
		bool intersects = !(range1.LastRow < range2.FirstRow ||
						   range1.LastColumn < range2.FirstColumn ||
						   range1.FirstRow > range2.LastRow ||
						   range1.FirstColumn > range2.LastColumn);

		return intersects;
	}
}

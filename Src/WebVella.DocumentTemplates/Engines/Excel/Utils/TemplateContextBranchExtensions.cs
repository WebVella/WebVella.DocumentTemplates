using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Engines.Excel.Models;

namespace WebVella.DocumentTemplates.Engines.Excel.Utility;
public static class TemplateContextBranchExtensions
{
	public static List<WvExcelFileTemplateContextBranch> GetByAddress(
		this List<WvExcelFileTemplateContextBranch> contextList,
		int worksheetPosition,
		int row,
		int column)
	{
		var result = new List<WvExcelFileTemplateContextBranch>();
		if (row < 1 || column < 1) return result;
		foreach (var context in contextList)
		{
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
	public static List<WvExcelFileTemplateContextBranch> GetIntersections(
		this List<WvExcelFileTemplateContextBranch> contextList,
		int worksheetPosition,
		WvExcelRange range)
	{
		if (range.FirstRow <= 0 || range.LastRow <= 0
			|| range.FirstColumn <= 0 || range.LastColumn <= 0)
			throw new ArgumentException("range positions should be > 0", nameof(range));

		if (range.FirstRow > range.LastRow)
			throw new ArgumentException("first row should be <= than last row", nameof(range));
		if (range.FirstColumn > range.LastColumn)
			throw new ArgumentException("first column should be <= than last row", nameof(range));

		var result = new List<WvExcelFileTemplateContextBranch>();

		foreach (var context in contextList)
		{
			if (context.Worksheet is null || context.Worksheet.Position != worksheetPosition) continue;

			if (context.Range is null) continue;

			var contextRange = new WvExcelRange()
			{
				FirstRow = context.Range!.FirstRow().RowNumber(),
				LastRow = context.Range!.LastRow().RowNumber(),
				FirstColumn = context.Range!.FirstColumn().ColumnNumber(),
				LastColumn = context.Range!.LastColumn().ColumnNumber()
			};
			if (CheckIntersection(range, contextRange))
				result.Add(context);
		}
		return result;
	}
	private static bool CheckIntersection(WvExcelRange range1, WvExcelRange range2)
	{
		// Check if the ranges intersect
		bool intersects = !(range1.LastRow < range2.FirstRow ||
						   range1.LastColumn < range2.FirstColumn ||
						   range1.FirstRow > range2.LastRow ||
						   range1.FirstColumn > range2.LastColumn);

		return intersects;
	}
}

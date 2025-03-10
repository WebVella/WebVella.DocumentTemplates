using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Engines.SpreadsheetFile.Models;

namespace WebVella.DocumentTemplates.Engines.SpreadsheetFile.Utility;
public static class WvSpreadsheetFileTemplateProcessContextExtensions
{
	public static List<WvSpreadsheetFileTemplateProcessContext> GetByAddress(
		this List<WvSpreadsheetFileTemplateProcessContext> contextList,
		int row,
		int column)
	{
		var result = new List<WvSpreadsheetFileTemplateProcessContext>();
		if (row < 1 || column < 1) return result;
		foreach (var context in contextList)
		{
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
	
}

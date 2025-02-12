using WebVella.DocumentTemplates.Core.Utility;

namespace WebVella.DocumentTemplates.Engines.Excel.Utility;
public static class TemplateContextExtensions
{
	public static List<WvExcelFileTemplateContext> GetByAddress(
		this List<WvExcelFileTemplateContext> contextList,
		int worksheetPosition,
		int row,
		int column,
		WvExcelFileTemplateContextType? type = null)
	{
		var result = new List<WvExcelFileTemplateContext>();
		if (row < 1 || column < 1) return result;
		foreach (var context in contextList)
		{
			if (type is not null && context.Type != type.Value) continue;
			if (context.WorksheetPosition != worksheetPosition) continue;

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

	public static void CalculateDependencies(this WvExcelFileTemplateContext context, List<WvExcelFileTemplateContext> contextList)
	{
		if (context is null) throw new ArgumentNullException(nameof(context));
		if (contextList is null || contextList.Count == 0) return;
		if (context.Range is null) return;
		if (context.Type != WvExcelFileTemplateContextType.CellRange) return;
		var rangeValues = context.Range.CellsUsed().Select(x => x.Value).ToList();
		if (rangeValues.Count == 0) return;
		foreach (var value in rangeValues)
		{
			var tags = WvTemplateUtility.GetTagsFromTemplate(value.ToString());
			foreach (var tag in tags) {
				if(tag.Type != Core.WvTemplateTagType.Function) continue;
				if(String.IsNullOrWhiteSpace(tag.FunctionName)) continue;
			}
		}
	}

}

using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using System.Data;
using System.Globalization;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Extensions;

namespace WebVella.DocumentTemplates.Engines.SpreadsheetFile.Utility;
public partial class WvSpreadsheetFileEngineUtility
{
	public WvTemplateTagResultList ProcessTemplateTag(
		string? template,
		DataTable dataSource,
		WvSpreadsheetFileTemplateContext templateContext,
		WvSpreadsheetFileTemplateProcessResultItem resultItem)
	{
		var result = new WvTemplateTagResultList();
		result.Tags = new WvTemplateUtility().GetTagsFromTemplate(template);
		result.ExpandCount = 1;
		foreach (var tag in result.Tags)
		{
			var hasSeparator = tag.ParamGroups.Any(g => g.Parameters.Any(x => x.Type.FullName == typeof(WvTemplateTagSeparatorParameterProcessor).FullName));
			if (tag.Type == WvTemplateTagType.Data)
			{
				//If not indexed and no separator is defined return the row count
				if (tag.IndexList.Count == 0 && !hasSeparator)
				{
					if (result.ExpandCount < dataSource.Rows.Count)
						result.ExpandCount = dataSource.Rows.Count;
				}
			}
			else if (tag.Type == WvTemplateTagType.Function
			|| tag.Type == WvTemplateTagType.SpreadsheetFunction)
			{
				//Expand with the parent if:
				//Parent is LeftContext
				//Parent is Forced
				if (
					(templateContext.LeftContext is not null
						&& templateContext.ParentContext is not null
						&& templateContext.LeftContext.Id == templateContext.ParentContext.Id)
					|| (templateContext.ForcedContext is not null
						&& templateContext.ParentContext is not null
						&& templateContext.ForcedContext.Id == templateContext.ParentContext.Id)
						)
				{
					var parentResult = resultItem.Contexts.Single(x => x.TemplateContextId == templateContext.ParentContext.Id);
					if (result.ExpandCount < parentResult.ExpandCount)
						result.ExpandCount = parentResult.ExpandCount;
				}
				//Do not expand in other cases
			}
		}

		return result;
	}
	private string getCellAddress(int worksheetPosition, int row, int col) => $"{worksheetPosition}:{row}:{col}";

	private void CopyCellProperties(IXLCell template, IXLRange result)
	{
		if (!String.IsNullOrWhiteSpace(template.FormulaA1))
			result.FormulaA1 = template.FormulaA1;
		if (!String.IsNullOrWhiteSpace(template.FormulaR1C1))
			result.FormulaR1C1 = template.FormulaR1C1;

		result.ShareString = template.ShareString;
		result.Style = template.Style;
		//We need to copy theme color differently as themed colors as the palletes can differ
		//for different types of cells - text or number
		//We need to copy theme color differently as themed colors as the palletes can differ
		//for different types of cells - text or number
		var templateColor = template.GetColor();
		var templateBackgroundColor = template.GetBackgroundColor();
		if (templateColor is not null)
		{
			result.Style.Font.SetFontColor(XLColor.FromColor(templateColor.Value));
		}
		if (templateBackgroundColor is not null)
		{
			result.Style.Fill.SetBackgroundColor(XLColor.FromColor(templateBackgroundColor.Value));
		}

	}

	private void CopyRowProperties(IXLRow origin, IXLRow result)
	{
		result.OutlineLevel = origin.OutlineLevel;
		result.Height = origin.Height;
	}

	private void CopyRowProperties(IXLRange templateRange, IXLRange resultRange,
		IXLWorksheet templateWs, IXLWorksheet resultWs,
		HashSet<long> processedRows)
	{
		var resultRangeRowsCount = resultRange.RowCount();
		for (var i = 0; i < templateRange.RowCount(); i++)
		{
			if (processedRows.Contains(i + 1)) continue;
			processedRows.Add(i + 1);
			if (resultRangeRowsCount - 1 < i) continue;
			var tempRow = templateWs.Row(templateRange.RangeAddress.FirstAddress.RowNumber + i);
			var resultRow = resultWs.Row(resultRange.RangeAddress.FirstAddress.RowNumber + i);
			CopyRowProperties(tempRow, resultRow);
		}
	}
	private void CopyColumnsProperties(IXLRange templateRange, IXLRange resultRange,
		IXLWorksheet templateWs, IXLWorksheet resultWs,
		HashSet<long> processedColumns)
	{
		var dsRowResultColumnsCount = resultRange.ColumnCount();
		for (var i = 0; i < templateRange.ColumnCount(); i++)
		{
			if (processedColumns.Contains(i + 1)) continue;
			processedColumns.Add(i + 1);
			if (dsRowResultColumnsCount - 1 < i) continue;
			var tempColumn = templateWs.Column(templateRange.RangeAddress.FirstAddress.ColumnNumber + i);
			var resultColumn = resultWs.Column(resultRange.RangeAddress.FirstAddress.ColumnNumber + i);
			resultColumn.OutlineLevel = tempColumn.OutlineLevel;
			resultColumn.Width = tempColumn.Width;
		}
	}

	private void addRangeToGrid(int firstRow, int firstColumn, int lastRow, int lastColumn, Guid contextId,
		WvSpreadsheetFileTemplateProcessResultItemRow resultRow, int resultFirstRow)
	{
		if (firstRow > 1)
		{
			firstRow -= resultFirstRow - 1;
			lastRow -= resultFirstRow - 1;
		}
		for (int rowNum = firstRow; rowNum <= lastRow; rowNum++)
		{
			for (int colNum = firstColumn; colNum <= lastColumn; colNum++)
			{
				if (resultRow.RelativeGridAddressHs.Contains($"{rowNum}:{colNum}"))
					throw new Exception($"Context overlap at result row: {rowNum} column: {colNum}");

				resultRow.RelativeGrid.Add(new WvSpreadsheetFileTemplateProcessResultItemRowGridItem
				{
					Row = rowNum,
					Column = colNum,
					ContextId = contextId,
				});
				resultRow.RelativeGridAddressHs.Add($"{rowNum}:{colNum}");
			}
		}

		if (resultRow.RelativeGridMaxRow < lastRow) resultRow.RelativeGridMaxRow = lastRow;
		if (resultRow.RelativeGridMaxColumn < lastColumn) resultRow.RelativeGridMaxColumn = lastColumn;
	}

}

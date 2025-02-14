using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using System.Data;
using System.Globalization;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Extensions;

namespace WebVella.DocumentTemplates.Engines.Excel.Utility;
public static partial class WvExcelFileEngineUtility
{
	private static string getCellAddress(int worksheetPosition, int row, int col) => $"{worksheetPosition}:{row}:{col}";

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
	private static void addRangeToGrid(int firstRow, int firstColumn, int lastRow, int lastColumn, Guid contextId,
		WvExcelFileTemplateProcessResultItemRow resultRow, int resultFirstRow)
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

				resultRow.RelativeGrid.Add(new WvExcelFileTemplateProcessResultItemRowGridItem
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

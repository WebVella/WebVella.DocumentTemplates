﻿using ClosedXML.Excel;
using System.Data;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Engines.SpreadsheetFile.Models;
using WebVella.DocumentTemplates.Engines.SpreadsheetFile.Utility;

namespace WebVella.DocumentTemplates.Engines.SpreadsheetFile.Functions;
public class AbsSpreadsheetFileTemplateFunction : IWvSpreadsheetFileTemplateFunctionProcessor
{
	public string Name { get; } = "abs";
	public int Priority { get; } = 10000;
	public bool HasError { get; set; }
	public string? ErrorMessage { get; set; }
	public object? Process(
			string? value,
			WvTemplateTag tag,
			WvSpreadsheetFileTemplateContext templateContext,
			int expandPosition,
			int expandPositionMax,
			DataTable dataSource,
			WvSpreadsheetFileTemplateProcessResult result,
			WvSpreadsheetFileTemplateProcessResultItem resultItem,
			IXLRange resultRange,
			IXLWorksheet worksheet
		)
	{
		if (tag is null) throw new ArgumentException(nameof(tag));
		if (dataSource is null) throw new ArgumentException(nameof(dataSource));
		if (result is null) throw new ArgumentException(nameof(result));
		if (resultItem is null) throw new ArgumentException(nameof(resultItem));
		if (worksheet is null) throw new ArgumentException(nameof(worksheet));
		if (tag.Type != WvTemplateTagType.Function)
			throw new ArgumentException("Template tag is not Function type", nameof(tag));
		if (String.IsNullOrWhiteSpace(tag.FullString)) return value;

		object? resultValue = null;
		decimal sum = 0;
		var rangeList = new WvSpreadsheetRangeHelpers().GetRangeAddressesForTag(
			tag: tag,
			templateContext: templateContext,
			expandPosition: expandPosition,
			expandPositionMax: expandPositionMax,
			result: result,
			resultItem: resultItem,
			worksheet: worksheet
		);
		var processedCellsHS = new HashSet<string>();
		foreach (var rangeAddress in rangeList)
		{
			var cellRange = worksheet.Range(rangeAddress);
			var firstAddress = cellRange.RangeAddress.FirstAddress;
			var lastAddress = cellRange.RangeAddress.LastAddress;
			for (var rowNum = firstAddress.RowNumber; rowNum <= lastAddress.RowNumber; rowNum++)
			{
				for (var colNum = firstAddress.ColumnNumber; colNum <= lastAddress.ColumnNumber; colNum++)
				{
					var resultCell = worksheet.Cell(rowNum, colNum);
					if (processedCellsHS.Contains(resultCell.Address.ToString() ?? "")) continue;
					var mergedRange = resultCell.MergedRange();
					var mergedRows = 1;
					var mergedCols = 1;
					if (mergedRange != null)
					{
						foreach (var cell in mergedRange.Cells())
						{
							processedCellsHS.Add(cell.Address.ToString() ?? "");
						}
						mergedRows = mergedRange.RowCount();
						mergedCols = mergedRange.ColumnCount();
					}
					else
					{
						processedCellsHS.Add(resultCell.Address.ToString() ?? "");
					}

					if (!resultCell.Value.IsNumber)
					{
						HasError = true;
						ErrorMessage = $"non numeric value in range";
						return null;
					}

					sum += (decimal)resultCell.Value.GetNumber();
				}
			}
		}

		sum = Math.Abs(sum);
		if (value is not null && value is string)
		{
			if (value == tag.FullString)
				resultValue = sum;
			else
				resultValue = ((string)value).Replace(tag.FullString ?? String.Empty, sum.ToString());
		}
		return resultValue;
	}
}

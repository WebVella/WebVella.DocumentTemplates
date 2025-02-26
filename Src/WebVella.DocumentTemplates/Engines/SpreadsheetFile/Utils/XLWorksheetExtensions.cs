using ClosedXML.Excel;

namespace WebVella.DocumentTemplates.Engines.SpreadsheetFile.Utility;
public static class XLWorksheetExtensions
{
	public static (int, int) GetGrossUsedRange(this IXLWorksheet ws)
	{
		//Process cells
		var usedRowsCount = ws.LastRowUsed()?.RowNumber() ?? 1;
		var usedColumnsCount = ws.LastColumnUsed()?.ColumnNumber() ?? 1;
		//Process pictures
		foreach (var picture in ws.Pictures)
		{
			int rowPosition = picture.TopLeftCell.Address.RowNumber;
			int colPosition = picture.TopLeftCell.Address.ColumnNumber;
			if (rowPosition > usedRowsCount) usedRowsCount = rowPosition;
			if (colPosition > usedColumnsCount) usedColumnsCount = colPosition;
		}

		//Styles check 50 rows and columns outside the current range for styles
		var checkRange = 50;
		var maxRows = usedRowsCount + checkRange;
		var maxColumns = usedColumnsCount + checkRange;
		for (int rowNum = 1; rowNum <= maxRows; rowNum++)
		{
			for (int colNum = 1; colNum <= maxColumns; colNum++)
			{
				if (rowNum <= usedRowsCount && colNum <= usedColumnsCount) continue;
				try
				{
					var cell = ws.Cell(rowNum, colNum);
					if (cell.HasContent())
					{
						if (rowNum > usedRowsCount) usedRowsCount = rowNum;
						if (colNum > usedColumnsCount) usedColumnsCount = colNum;
					}
				}
				catch
				{
					break;
				}
			}
		}

		return (usedRowsCount, usedColumnsCount);
	}

}

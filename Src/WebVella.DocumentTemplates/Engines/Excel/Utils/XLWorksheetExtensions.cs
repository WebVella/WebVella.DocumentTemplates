using ClosedXML.Excel;
using System.Data;

namespace WebVella.DocumentTemplates.Engines.Excel.Utility;
public static class XLWorksheetExtensions
{
	public static (int, int) GetUsedRangeWithEmbeds(this IXLWorksheet ws)
	{
		//Process cells
		var usedRowsCount = ws.LastRowUsed()?.RowNumber() ?? 1;
		var usedColumnsCount = ws.LastColumnUsed()?.ColumnNumber() ?? 1;
		//Process pictures
		foreach (var picture in ws.Pictures)
		{
			int rowPosition = picture.TopLeftCell.Address.RowNumber;
			int colPosition = picture.TopLeftCell.Address.ColumnNumber;
			if(rowPosition > usedRowsCount) usedRowsCount = rowPosition;
			if(colPosition > usedColumnsCount) usedColumnsCount = colPosition;
		}


		return (usedRowsCount, usedColumnsCount);
	}

}

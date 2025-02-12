using ClosedXML.Excel;
using System.Data;
using System.Text.RegularExpressions;
using WebVella.DocumentTemplates.Engines.Excel.Models;

namespace WebVella.DocumentTemplates.Engines.Excel.Utility;
public static class WvExcelRangeHelpers
{
	public static WvExcelRange? GetRangeFromString(string address)
	{
		if (String.IsNullOrEmpty(address)) return null;

		var result = new WvExcelRange();
		address = address.Trim();
		if (address.Contains('!'))
		{
			var split = address.Split('!', StringSplitOptions.RemoveEmptyEntries);
			if (split.Length != 2) return null;
			result.Worksheet = split[0];
			address = split[1];
		}

		//if ranged
		if (address.Contains(":"))
		{
			var split = address.Split(':', StringSplitOptions.RemoveEmptyEntries);
			(result.FirstRow, result.FirstColumn) = ExtractRowColumnFromAddress(split[0]);
			(result.LastRow, result.LastColumn) = ExtractRowColumnFromAddress(split[1]);
		}
		else
		{
			(result.FirstRow, result.FirstColumn) = ExtractRowColumnFromAddress(address);
			result.LastRow = result.FirstRow;
			result.LastColumn = result.FirstColumn;
		}

		if (result.FirstRow == 0
		|| result.FirstColumn == 0
		|| result.LastRow == 0
		|| result.LastColumn == 0) return null;

		return result;
	}

	private static (int, int) ExtractRowColumnFromAddress(string cellAddress)
	{
		int row = 0;
		int column = 0;
		if (String.IsNullOrWhiteSpace(cellAddress)) return (row, column);
		cellAddress = cellAddress.Trim();
		if (!XLHelper.IsValidA1Address(cellAddress)) return (row, column);
		if (!cellAddress.Any(i => char.IsDigit(i))) return (row, column);
		int indexOfFirstDigit = Enumerable.Range(0, cellAddress.Length).FirstOrDefault(i => char.IsDigit(cellAddress[i]));

		var columnString = cellAddress.Substring(0,indexOfFirstDigit);
		var rowString = cellAddress.Substring(indexOfFirstDigit);

		if (XLHelper.IsValidRow(rowString))
		{
			if(int.TryParse(rowString, out int outInt)) row = outInt;
		}
		if (XLHelper.IsValidColumn(columnString))
		{
			column = XLHelper.GetColumnNumberFromLetter(columnString);
		}

		return (row, column);
	}
}

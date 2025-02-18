using ClosedXML.Excel;
using System.Data;
using System.Drawing;
using System.Text.RegularExpressions;
using WebVella.DocumentTemplates.Engines.Excel.Models;

namespace WebVella.DocumentTemplates.Engines.Excel.Utility;
public class WvExcelRangeHelpers
{
	public WvExcelRange? GetRangeFromString(string address)
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

		if (result.FirstRow > result.LastRow)
		{
			result = result with
			{
				FirstRow = result.LastRow,
				LastRow = result.FirstRow,
			};
		}

		if (result.FirstColumn > result.LastColumn)
		{
			result = result with
			{
				FirstColumn = result.LastColumn,
				LastColumn = result.FirstColumn,
			};
		}

		return result;
	}

	private (int, int) ExtractRowColumnFromAddress(string cellAddress)
	{
		int row = 0;
		int column = 0;
		if (String.IsNullOrWhiteSpace(cellAddress)) return (row, column);
		cellAddress = cellAddress.Trim();
		if (!XLHelper.IsValidA1Address(cellAddress)) return (row, column);
		if (!cellAddress.Any(i => char.IsDigit(i))) return (row, column);
		int indexOfFirstDigit = Enumerable.Range(0, cellAddress.Length).FirstOrDefault(i => char.IsDigit(cellAddress[i]));

		var columnString = cellAddress.Substring(0, indexOfFirstDigit);
		var rowString = cellAddress.Substring(indexOfFirstDigit);

		if (XLHelper.IsValidRow(rowString))
		{
			if (int.TryParse(rowString, out int outInt)) row = outInt;
		}
		if (XLHelper.IsValidColumn(columnString))
		{
			column = XLHelper.GetColumnNumberFromLetter(columnString);
		}

		return (row, column);
	}

	public Color ApplyTint(Color color, double tint)
	{
		if (tint < -1.0) tint = -1.0;
		if (tint > 1.0) tint = 1.0;

		Color colorRgb = color;
		double fHue = colorRgb.GetHue();
		double fSat = colorRgb.GetSaturation();
		double fLum = colorRgb.GetBrightness();
		if (tint < 0)
		{
			fLum = fLum * (1.0 + tint);
		}
		else
		{
			fLum = fLum * (1.0 - tint) + (1.0 - 1.0 * (1.0 - tint));
		}
		return ToColor(fHue, fSat, fLum);
	}

	public Color ToColor(double hue, double saturation, double luminance)
	{
		double chroma = (1.0 - Math.Abs(2.0 * luminance - 1.0)) * saturation;
		double fHue = hue / 60.0;
		double fHueMod2 = fHue;
		while (fHueMod2 >= 2.0) fHueMod2 -= 2.0;
		double fTemp = chroma * (1.0 - Math.Abs(fHueMod2 - 1.0));

		double fRed, fGreen, fBlue;
		if (fHue < 1.0)
		{
			fRed = chroma;
			fGreen = fTemp;
			fBlue = 0;
		}
		else if (fHue < 2.0)
		{
			fRed = fTemp;
			fGreen = chroma;
			fBlue = 0;
		}
		else if (fHue < 3.0)
		{
			fRed = 0;
			fGreen = chroma;
			fBlue = fTemp;
		}
		else if (fHue < 4.0)
		{
			fRed = 0;
			fGreen = fTemp;
			fBlue = chroma;
		}
		else if (fHue < 5.0)
		{
			fRed = fTemp;
			fGreen = 0;
			fBlue = chroma;
		}
		else if (fHue < 6.0)
		{
			fRed = chroma;
			fGreen = 0;
			fBlue = fTemp;
		}
		else
		{
			fRed = 0;
			fGreen = 0;
			fBlue = 0;
		}

		double fMin = luminance - 0.5 * chroma;
		fRed += fMin;
		fGreen += fMin;
		fBlue += fMin;

		fRed *= 255.0;
		fGreen *= 255.0;
		fBlue *= 255.0;

		var red = Convert.ToInt32(Math.Truncate(fRed));
		var green = Convert.ToInt32(Math.Truncate(fGreen));
		var blue = Convert.ToInt32(Math.Truncate(fBlue));

		red = Math.Min(255, Math.Max(red, 0));
		green = Math.Min(255, Math.Max(green, 0));
		blue = Math.Min(255, Math.Max(blue, 0));

		return Color.FromArgb(red, green, blue);
	}

}

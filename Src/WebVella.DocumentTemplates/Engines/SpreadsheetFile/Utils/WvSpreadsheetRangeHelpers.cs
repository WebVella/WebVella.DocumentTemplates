﻿using ClosedXML.Excel;
using System.Data;
using System.Drawing;
using System.Text.RegularExpressions;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Engines.SpreadsheetFile.Models;

namespace WebVella.DocumentTemplates.Engines.SpreadsheetFile.Utility;
public class WvSpreadsheetRangeHelpers
{
	public WvSpreadsheetRange? GetRangeFromString(string address)
	{
		if (String.IsNullOrEmpty(address)) return null;

		var result = new WvSpreadsheetRange();
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
			(result.FirstRow, result.FirstRowLocked, result.FirstColumn, result.FirstColumnLocked) = ExtractRowColumnFromAddress(split[0]);
			(result.LastRow, result.LastRowLocked, result.LastColumn, result.LastColumnLocked) = ExtractRowColumnFromAddress(split[1]);
		}
		else
		{
			(result.FirstRow, result.FirstRowLocked, result.FirstColumn, result.FirstColumnLocked) = ExtractRowColumnFromAddress(address);
			result.LastRow = result.FirstRow;
			result.LastRowLocked = result.FirstRowLocked;
			result.LastColumn = result.FirstColumn;
			result.LastColumnLocked = result.FirstColumnLocked;
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
				FirstRowLocked = result.LastRowLocked,
				LastRow = result.FirstRow,
				LastRowLocked = result.FirstRowLocked
			};
		}

		if (result.FirstColumn > result.LastColumn)
		{
			result = result with
			{
				FirstColumn = result.LastColumn,
				FirstColumnLocked = result.LastColumnLocked,
				LastColumn = result.FirstColumn,
				LastColumnLocked = result.FirstColumnLocked
			};
		}

		return result;
	}

	private (int, bool, int, bool) ExtractRowColumnFromAddress(string cellAddress)
	{
		int row = 0;
		bool rowLocked = false;
		int column = 0;
		bool columnLocked = false;
		if (String.IsNullOrWhiteSpace(cellAddress)) return (row, rowLocked, column, columnLocked);
		cellAddress = cellAddress.Trim();
		if (!XLHelper.IsValidA1Address(cellAddress)) return (row, rowLocked, column, columnLocked);
		if (!cellAddress.Any(i => char.IsDigit(i))) return (row, rowLocked, column, columnLocked);
		int indexOfFirstDigit = Enumerable.Range(0, cellAddress.Length).FirstOrDefault(i => char.IsDigit(cellAddress[i]));

		var columnString = cellAddress.Substring(0, indexOfFirstDigit);
		var rowString = cellAddress.Substring(indexOfFirstDigit);
		if (columnString.EndsWith("$"))
		{
			columnString = columnString.Substring(0, columnString.Length - 1);
			rowString = "$" + rowString;
		}

		if (rowString.StartsWith("$"))
		{
			rowLocked = true;
			rowString = rowString.Substring(1);
		}
		if (columnString.StartsWith("$"))
		{
			columnLocked = true;
			columnString = columnString.Substring(1);
		}

		if (XLHelper.IsValidRow(rowString))
		{
			if (int.TryParse(rowString, out int outInt)) row = outInt;
		}

		if (XLHelper.IsValidColumn(columnString))
		{
			column = XLHelper.GetColumnNumberFromLetter(columnString);
		}

		return (row, rowLocked, column, columnLocked);
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

	public List<string> GetRangeAddressesForTag(
		WvTemplateTag tag,
		WvSpreadsheetFileTemplateContext templateContext,
		int expandPosition,
		int expandPositionMax,
		WvSpreadsheetFileTemplateProcessResult result,
		WvSpreadsheetFileTemplateProcessResultItem resultItem,
		IXLWorksheet worksheet
		)
	{
		var rangeList = new List<string>();
		foreach (var paramGroup in tag.ParamGroups)
		{
			foreach (var param in paramGroup.Parameters)
			{
				if (String.IsNullOrWhiteSpace(param.ValueString)) continue;
				var range = new WvSpreadsheetRangeHelpers().GetRangeFromString(param.ValueString ?? String.Empty);
				if (range is not null)
				{
					var rangeTemplateContexts = result.TemplateContexts.GetIntersections(
						worksheetPosition: worksheet.Position,
						range: range,
						type: WvSpreadsheetFileTemplateContextType.CellRange
					);
					var resultContexts = resultItem.Contexts.Where(x => rangeTemplateContexts.Any(y => y.Id == x.TemplateContextId));
					if (resultContexts.Count() == 0) continue;

					WvSpreadsheetFileTemplateProcessContext firstResultContext = resultContexts.First();
					WvSpreadsheetFileTemplateProcessContext lastResultContext = resultContexts.Last();
					var processedContexts = new List<WvSpreadsheetFileTemplateProcessContext> { firstResultContext };
					if(firstResultContext.TemplateContextId !=  lastResultContext.TemplateContextId)
						processedContexts.Add(lastResultContext);

					foreach (var resultCtx in processedContexts)
					{
						if (resultCtx.Range?.RangeAddress is null) continue;

						if (expandPositionMax <= 1)
						{
							//the function row will not  expends
							rangeList.Add(resultCtx.Range.RangeAddress.ToString() ?? String.Empty);
						}
						else
						{
							//the function row expends
							var relativeRangeFirstRow = resultCtx.Range.RangeAddress.FirstAddress.RowNumber;
							var relativeRangeFirstColumn = resultCtx.Range.RangeAddress.FirstAddress.ColumnNumber;
							var relativeRangeLastRow = relativeRangeFirstRow;
							var relativeRangeLastColumn = relativeRangeFirstColumn;

							var restCtxCell = worksheet.Cell(relativeRangeFirstRow, relativeRangeFirstColumn);
							var mergedRange = restCtxCell.MergedRange();
							if (mergedRange is not null)
							{
								relativeRangeLastRow += mergedRange.RowCount() - 1;
								relativeRangeLastColumn = mergedRange.ColumnCount() - 1;
							}

							if (resultCtx.TemplateContextId == firstResultContext.TemplateContextId)
							{
								if (templateContext.Flow == WvTemplateTagDataFlow.Vertical)
								{
									if (!range.FirstRowLocked)
									{
										relativeRangeFirstRow += (expandPosition - 1);
										relativeRangeLastRow += (expandPosition - 1);
									}
								}
								else
								{
									if (!range.FirstColumnLocked)
									{
										relativeRangeFirstColumn += (expandPosition - 1);
										relativeRangeLastColumn += (expandPosition - 1);
									}
								}
							}
							else
							{
								if (templateContext.Flow == WvTemplateTagDataFlow.Vertical)
								{
									if (!range.LastRowLocked)
									{
										relativeRangeFirstRow += (expandPosition - 1);
										relativeRangeLastRow += (expandPosition - 1);
									}
								}
								else
								{
									if (!range.LastColumnLocked)
									{
										relativeRangeFirstColumn += (expandPosition - 1);
										relativeRangeLastColumn += (expandPosition - 1);
									}
								}
							}
							rangeList.Add(worksheet.Range(relativeRangeFirstRow, relativeRangeFirstColumn, relativeRangeLastRow, relativeRangeLastColumn).RangeAddress.ToString() ?? String.Empty);
						}
					}
				}
			}
		}

		return rangeList;
	}
}

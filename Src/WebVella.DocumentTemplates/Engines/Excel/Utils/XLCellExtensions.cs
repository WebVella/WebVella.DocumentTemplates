using ClosedXML.Excel;
using System.Data;
using System.Drawing;

namespace WebVella.DocumentTemplates.Engines.Excel.Utility;
public static class XLCellExtension
{
	public static bool HasContent(this IXLCell cell)
	{
		if (!String.IsNullOrWhiteSpace(cell.Value.ToString())) return true;
		if (hasStyle(cell)) return true;
		return false;
	}

	private static bool hasStyle(this IXLCell cell)
	{
		//Alignment does not matter here as this style checks after the content is checked
		if (hasBorder(cell)) return true;
		if (hasColor(cell)) return true;

		return false;
	}


	private static bool hasBorder(this IXLCell cell)
	{
		if (cell.Style.Border.LeftBorder != XLBorderStyleValues.None) return true;
		if (cell.Style.Border.RightBorder != XLBorderStyleValues.None) return true;
		if (cell.Style.Border.TopBorder != XLBorderStyleValues.None) return true;
		if (cell.Style.Border.BottomBorder != XLBorderStyleValues.None) return true;
		if (cell.Style.Border.DiagonalBorder != XLBorderStyleValues.None) return true;
		return false;
	}
	private static bool hasColor(this IXLCell cell)
	{
		if (cell.Style.Fill.PatternType != XLFillPatternValues.None) return true;

		return false;
	}

	public static Color? GetBackgroundColor(this IXLCell cell)
	{
		Color? color = null;
		if (cell.Style.Fill.BackgroundColor.ColorType == XLColorType.Color)
		{
			return cell.Style.Fill.BackgroundColor.Color;
		}
		if (cell.Style.Fill.BackgroundColor.ColorType == XLColorType.Indexed)
		{
			XLColor c = XLColor.FromIndex(cell.Style.Fill.BackgroundColor.Indexed);
			color = c.Color;
			return color;
		}
		if (cell.Style.Fill.BackgroundColor.ColorType == XLColorType.Theme)
		{
			XLWorkbook workbook = cell.Worksheet.Workbook;

			double tint = cell.Style.Fill.BackgroundColor.ThemeTint;
			XLThemeColor themeColor = cell.Style.Fill.BackgroundColor.ThemeColor;
			if (themeColor == XLThemeColor.Accent1)
			{
				color = workbook.Theme.Accent1.Color;
			}
			else if (themeColor == XLThemeColor.Accent2)
			{
				color = workbook.Theme.Accent2.Color;
			}
			else if (themeColor == XLThemeColor.Accent3)
			{
				color = workbook.Theme.Accent3.Color;
			}
			else if (themeColor == XLThemeColor.Accent4)
			{
				color = workbook.Theme.Accent4.Color;
			}
			else if (themeColor == XLThemeColor.Accent5)
			{
				color = workbook.Theme.Accent5.Color;
			}
			else if (themeColor == XLThemeColor.Accent6)
			{
				color = workbook.Theme.Accent6.Color;
			}
			else if (themeColor == XLThemeColor.Background1)
			{
				color = workbook.Theme.Background1.Color;
			}
			else if (themeColor == XLThemeColor.Background2)
			{
				color = workbook.Theme.Background2.Color;
			}
			else if (themeColor == XLThemeColor.FollowedHyperlink)
			{
				color = workbook.Theme.FollowedHyperlink.Color;
			}
			else if (themeColor == XLThemeColor.Hyperlink)
			{
				color = workbook.Theme.Hyperlink.Color;
			}
			else if (themeColor == XLThemeColor.Text1)
			{
				color = workbook.Theme.Text1.Color;
			}
			else if (themeColor == XLThemeColor.Text2)
			{
				color = workbook.Theme.Text2.Color;
			}
			if (color is not null)
				color = new WvExcelRangeHelpers().ApplyTint(color.Value,tint);

		}
		return color;
	}

	public static Color? GetColor(this IXLCell cell)
	{
		Color? color = null;
		if (cell.Style.Font.FontColor.ColorType == XLColorType.Color)
		{
			return cell.Style.Font.FontColor.Color;
		}
		if (cell.Style.Font.FontColor.ColorType == XLColorType.Indexed)
		{
			XLColor c = XLColor.FromIndex(cell.Style.Font.FontColor.Indexed);
			color = c.Color;
			return color;
		}
		if (cell.Style.Font.FontColor.ColorType == XLColorType.Theme)
		{
			XLWorkbook workbook = cell.Worksheet.Workbook;


			double tint = cell.Style.Font.FontColor.ThemeTint;
			XLThemeColor themeColor = cell.Style.Font.FontColor.ThemeColor;
			if (themeColor == XLThemeColor.Accent1)
			{
				color = workbook.Theme.Accent1.Color;
			}
			else if (themeColor == XLThemeColor.Accent2)
			{
				color = workbook.Theme.Accent2.Color;
			}
			else if (themeColor == XLThemeColor.Accent3)
			{
				color = workbook.Theme.Accent3.Color;
			}
			else if (themeColor == XLThemeColor.Accent4)
			{
				color = workbook.Theme.Accent4.Color;
			}
			else if (themeColor == XLThemeColor.Accent5)
			{
				color = workbook.Theme.Accent5.Color;
			}
			else if (themeColor == XLThemeColor.Accent6)
			{
				color = workbook.Theme.Accent6.Color;
			}
			else if (themeColor == XLThemeColor.Background1)
			{
				color = workbook.Theme.Background1.Color;
			}
			else if (themeColor == XLThemeColor.Background2)
			{
				color = workbook.Theme.Background2.Color;
			}
			else if (themeColor == XLThemeColor.FollowedHyperlink)
			{
				color = workbook.Theme.FollowedHyperlink.Color;
			}
			else if (themeColor == XLThemeColor.Hyperlink)
			{
				color = workbook.Theme.Hyperlink.Color;
			}
			else if (themeColor == XLThemeColor.Text1)
			{
				color = workbook.Theme.Text1.Color;
			}
			else if (themeColor == XLThemeColor.Text2)
			{
				color = workbook.Theme.Text2.Color;
			}
			if (color is not null)
				color = new WvExcelRangeHelpers().ApplyTint(color.Value,tint);
		}
		return color;
	}
}

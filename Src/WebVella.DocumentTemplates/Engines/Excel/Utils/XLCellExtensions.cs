using ClosedXML.Excel;
using System.Data;

namespace WebVella.DocumentTemplates.Engines.Excel.Utility;
public static class XLCellExtension
{
	public static bool HasContent(this IXLCell cell)
	{
		if(!String.IsNullOrWhiteSpace(cell.Value.ToString())) return true;
		if(hasStyle(cell)) return true;
		return false;
	}

	private static bool hasStyle(this IXLCell cell){ 
		//Alignment does not matter here as this style checks after the content is checked
		if(hasBorder(cell)) return true;
		if(hasColor(cell)) return true;

		return false;
	}


	private static bool hasBorder(this IXLCell cell){ 
		if(cell.Style.Border.LeftBorder != XLBorderStyleValues.None) return true;
		if(cell.Style.Border.RightBorder != XLBorderStyleValues.None) return true;
		if(cell.Style.Border.TopBorder != XLBorderStyleValues.None) return true;
		if(cell.Style.Border.BottomBorder != XLBorderStyleValues.None) return true;
		if(cell.Style.Border.DiagonalBorder != XLBorderStyleValues.None) return true;
		return false;
	}
	private static bool hasColor(this IXLCell cell){ 
		if(cell.Style.Fill.PatternType != XLFillPatternValues.None) return true;

		return false;
	}
}

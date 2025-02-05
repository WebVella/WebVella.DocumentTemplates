﻿using ClosedXML.Excel;
using System.Data;
using System.Globalization;
using System.Text;
using WebVella.DocumentTemplates.Engines.Excel;

namespace WebVella.DocumentTemplates.Tests.Models;
public class TestBase
{
	public DataTable SampleData;
	public DataTable TypedData;
	public CultureInfo DefaultCulture = new CultureInfo("en-US");
	public TestBase()
	{
		//Data
		{
			var ds = new DataTable();
			ds.Columns.Add("position", typeof(int));
			ds.Columns.Add("sku", typeof(string));
			ds.Columns.Add("name", typeof(string));
			ds.Columns.Add("price", typeof(decimal));

			for (int i = 0; i < 5; i++)
			{
				var position = i + 1;
				var dsrow = ds.NewRow();
				dsrow["position"] = position;
				dsrow["sku"] = $"sku{position}";
				dsrow["name"] = $"item{position}";
				dsrow["price"] = (decimal)(position * 10);
				ds.Rows.Add(dsrow);
			}
			SampleData = ds;
		}
		//TypedData
		{
			var ds = new DataTable();
			ds.Columns.Add("short", typeof(short));
			ds.Columns.Add("int", typeof(int));
			ds.Columns.Add("long", typeof(long));
			ds.Columns.Add("number", typeof(decimal));
			ds.Columns.Add("date", typeof(DateOnly));
			ds.Columns.Add("datetime", typeof(DateTime));
			ds.Columns.Add("shorttext", typeof(string));
			ds.Columns.Add("text", typeof(string));
			ds.Columns.Add("guid", typeof(Guid));
			for (short i = 0; i < 5; i++)
			{
				var position = i + 1;
				var dsrow = ds.NewRow();
				dsrow["short"] = (short)position;
				dsrow["int"] = (int)(position + 100);
				dsrow["long"] = (long)(position + 1000);
				dsrow["number"] = (decimal)(position * 10);
				dsrow["date"] = DateOnly.FromDateTime(DateTime.Now.AddDays(i));
				dsrow["datetime"] = DateTime.Now.AddDays(i + 10);
				dsrow["shorttext"] = $"short text {i}";
				dsrow["text"] = $"text {i}";
				dsrow["guid"] = Guid.NewGuid();
				ds.Rows.Add(dsrow);
			}
			TypedData = ds;
		}
	}

	public static DataTable CreateDataTableFromExisting(DataTable originalTable, List<int>? rowIndices = null)
	{
		// Create a new DataTable with the same structure as the original table
		DataTable newTable = new DataTable();

		foreach (DataColumn column in originalTable.Columns)
		{
			newTable.Columns.Add(column.ColumnName, column.DataType);
		}

		// Check if there is at least one row in the original table
		if (originalTable.Rows.Count > 0)
		{
			var currentIndex = -1;
			foreach (DataRow row in originalTable.Rows)
			{
				currentIndex++;
				if (rowIndices is not null && !rowIndices.Contains(currentIndex)) continue;
				DataRow newRow = newTable.NewRow();

				foreach (DataColumn column in originalTable.Columns)
				{
					newRow[column.ColumnName] = row[column.ColumnName];
				}
				newTable.Rows.Add(newRow);
			}
		}
		return newTable;
	}

	public static List<string> GetLines(string content)
	{
		if (string.IsNullOrEmpty(content)) return new List<string>();
		return content.Split("\r\n".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();
	}

	public static byte[] LoadFile(string fileName)
	{
		var path = Path.Combine(Environment.CurrentDirectory, $"Files\\{fileName}");
		var fi = new FileInfo(path);
		if (!fi.Exists) throw new FileNotFoundException();
		return File.ReadAllBytes(path);
	}
	public static void SaveFile(string content, string fileName)
	{
		DirectoryInfo? debugFolder = Directory.GetParent(Environment.CurrentDirectory);
		if (debugFolder is null || !debugFolder.Exists) throw new Exception("Debug Folder not found");
		var projectFolder = debugFolder.Parent?.Parent;
		if (projectFolder is null || !projectFolder.Exists) throw new Exception("Project Folder not found");

		string path = Path.Combine(projectFolder.FullName, $"FileResults\\result-{fileName}");
		StreamWriter sw = new StreamWriter(path, false, Encoding.ASCII);
		sw.Write(content);
		sw.Close();
	}

	private static void GeneralResultChecks(WvExcelFileTemplateResult result)
	{
		Assert.NotNull(result);
		Assert.NotNull(result.Template);
		Assert.NotNull(result.Result);
		Assert.NotNull(result.Contexts);

		Assert.NotNull(result.Template.Worksheets);
		Assert.True(result.Template.Worksheets.Count > 0);

		Assert.NotNull(result.Result.Worksheets);
		Assert.True(result.Result.Worksheets.Count > 0);
	}

	private static void CheckRangeDimensions(IXLRange range, int startRowNumber, int startColumnNumber, int lastRowNumber, int lastColumnNumber)
	{
		Assert.NotNull(range);
		Assert.NotNull(range.RangeAddress);
		Assert.NotNull(range.RangeAddress.FirstAddress);
		Assert.NotNull(range.RangeAddress.LastAddress);
		Assert.Equal(startRowNumber, range.RangeAddress.FirstAddress.RowNumber);
		Assert.Equal(startColumnNumber, range.RangeAddress.FirstAddress.ColumnNumber);

		Assert.Equal(lastRowNumber, range.RangeAddress.LastAddress.RowNumber);
		Assert.Equal(lastColumnNumber, range.RangeAddress.LastAddress.ColumnNumber);
	}

	private static void CheckCellPropertiesCopy(WvExcelFileTemplateResult result)
	{
		foreach (var context in result.Contexts)
		{
			var firstTemplateCell = context.TemplateRange?.Cell(1, 1);
			var firstResultCell = context.ResultRange?.Cell(1, 1);
			Assert.NotNull(firstTemplateCell);
			Assert.NotNull(firstResultCell);

			Assert.Equal(firstTemplateCell.ShowPhonetic, firstResultCell.ShowPhonetic);
			Assert.Equal(firstTemplateCell.FormulaA1, firstResultCell.FormulaA1);
			Assert.Equal(firstTemplateCell.FormulaR1C1, firstResultCell.FormulaR1C1);
			Assert.Equal(firstTemplateCell.ShareString, firstResultCell.ShareString);
			Assert.Equal(firstTemplateCell.Active, firstResultCell.Active);

			CompareStyle(firstTemplateCell, firstResultCell);
		}
	}
	private static void CompareStyle(IXLCell template, IXLCell result)
	{
		if (template.Style is null)
		{
			Assert.Null(result.Style);
		}
		else
		{
			Assert.NotNull(result.Style);
			CompareStyleAlignment(template.Style.Alignment, result.Style.Alignment);
			CompareStyleBorder(template.Style.Border, result.Style.Border);
			CompareStyleFormat(template.Style.DateFormat, result.Style.DateFormat);
			CompareStyleFill(template.Style.Fill, result.Style.Fill);
			CompareStyleFont(template.Style.Font, result.Style.Font);
			Assert.Equal(template.Style.IncludeQuotePrefix, result.Style.IncludeQuotePrefix);
			CompareStyleFormat(template.Style.NumberFormat, result.Style.NumberFormat);
			CompareProtection(template.Style.Protection, result.Style.Protection);
		}
	}

	private static void CompareStyleAlignment(IXLAlignment template, IXLAlignment result)
	{
		Assert.Equal(template.TopToBottom, result.TopToBottom);
		Assert.Equal(template.TextRotation, result.TextRotation);
		Assert.Equal(template.ShrinkToFit, result.ShrinkToFit);
		Assert.Equal(template.RelativeIndent, result.RelativeIndent);
		Assert.Equal(template.ReadingOrder, result.ReadingOrder);
		Assert.Equal(template.JustifyLastLine, result.JustifyLastLine);
		Assert.Equal(template.Indent, result.Indent);
		Assert.Equal(template.Vertical, result.Vertical);
		Assert.Equal(template.Horizontal, result.Horizontal);
		Assert.Equal(template.WrapText, result.WrapText);
	}
	private static void CompareStyleBorder(IXLBorder template, IXLBorder result)
	{
		CompareStyleColor(template.DiagonalBorderColor, result.DiagonalBorderColor);
		Assert.Equal(template.LeftBorder, result.LeftBorder);
		CompareStyleColor(template.LeftBorderColor, result.LeftBorderColor);
		Assert.Equal(template.DiagonalBorder, result.DiagonalBorder);
		Assert.Equal(template.RightBorder, result.RightBorder);
		Assert.Equal(template.TopBorder, result.TopBorder);
		CompareStyleColor(template.TopBorderColor, result.TopBorderColor);
		Assert.Equal(template.BottomBorder, result.BottomBorder);
		CompareStyleColor(template.BottomBorderColor, result.BottomBorderColor);
		Assert.Equal(template.DiagonalUp, result.DiagonalUp);
		Assert.Equal(template.DiagonalDown, result.DiagonalDown);
		CompareStyleColor(template.RightBorderColor, result.RightBorderColor);
	}
	private static void CompareStyleFormat(IXLNumberFormat template, IXLNumberFormat result)
	{
		Assert.Equal(template.NumberFormatId, result.NumberFormatId);
		Assert.Equal(template.Format, result.Format);
	}
	private static void CompareStyleFill(IXLFill template, IXLFill result)
	{
		CompareStyleColor(template.BackgroundColor, result.BackgroundColor);
		CompareStyleColor(template.PatternColor, result.PatternColor);
		Assert.Equal(template.PatternType, result.PatternType);
	}
	private static void CompareStyleFont(IXLFont template, IXLFont result)
	{
		Assert.Equal(template.Bold, result.Bold);
		Assert.Equal(template.Italic, result.Italic);
		Assert.Equal(template.Underline, result.Underline);
		Assert.Equal(template.Strikethrough, result.Strikethrough);
		Assert.Equal(template.VerticalAlignment, result.VerticalAlignment);
		Assert.Equal(template.Shadow, result.Shadow);
		Assert.Equal(template.FontSize, result.FontSize);
		CompareStyleColor(template.FontColor, result.FontColor);
		Assert.Equal(template.FontName, result.FontName);
		Assert.Equal(template.FontFamilyNumbering, result.FontFamilyNumbering);
		Assert.Equal(template.FontCharSet, result.FontCharSet);
		Assert.Equal(template.FontScheme, result.FontScheme);
	}
	private static void CompareProtection(IXLProtection template, IXLProtection result)
	{
		Assert.Equal(template.Locked, result.Locked);
		Assert.Equal(template.Hidden, result.Hidden);
	}
	private static void CompareStyleColor(XLColor template, XLColor result)
	{
		if (template is null)
		{
			Assert.Null(result);
		}
		else
		{
			Assert.NotNull(result);
			Assert.Equal(template.ToString(), result.Color.ToString());
		}
	}

	private static void CompareRowProperties(IXLRow template, IXLRow result)
	{
		Assert.Equal(template.OutlineLevel,result.OutlineLevel);
		Assert.Equal(template.Height,result.Height);
	}
	private static void CompareColumnProperties(IXLColumn template, IXLColumn result)
	{
		Assert.Equal(template.OutlineLevel,result.OutlineLevel);
		Assert.Equal(template.Width,result.Width);
	}

	private static XLWorkbook LoadWorkbook(string fileName)
	{
		var path = Path.Combine(Environment.CurrentDirectory, $"Files\\{fileName}");
		var fi = new FileInfo(path);
		Assert.NotNull(fi);
		Assert.True(fi.Exists);
		var templateWB = new XLWorkbook(path);
		Assert.NotNull(templateWB);
		return templateWB;
	}
	private static void SaveWorkbook(XLWorkbook workbook, string fileName)
	{
		DirectoryInfo? debugFolder = Directory.GetParent(Environment.CurrentDirectory);
		if(debugFolder is null) throw new Exception("debugFolder not found");
		var projectFolder = debugFolder.Parent?.Parent;
		if(projectFolder is null) throw new Exception("projectFolder not found");

		var path = Path.Combine(projectFolder.FullName, $"FileResults\\result-{fileName}");
		workbook.SaveAs(path);
	}
}

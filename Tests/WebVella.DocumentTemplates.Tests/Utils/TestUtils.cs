using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WebVella.DocumentTemplates.Engines.DocumentFile;
using WebVella.DocumentTemplates.Engines.SpreadsheetFile;

namespace WebVella.DocumentTemplates.Tests.Utils;
public class TestUtils
{

	public List<string> GetLines(string content)
	{
		if (string.IsNullOrEmpty(content)) return new List<string>();
		return content.Split("\r\n".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();
	}

	public byte[] LoadFile(string fileName)
	{
		var path = Path.Combine(Environment.CurrentDirectory, $"Files\\{fileName}");
		var fi = new FileInfo(path);
		if (!fi.Exists) throw new FileNotFoundException();
		return File.ReadAllBytes(path);
	}

	public MemoryStream LoadFileStream(string fileName)
	{
		return new MemoryStream(LoadFile(fileName));
	}

	public void SaveFile(string content, string fileName)
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

	public MemoryStream LoadFileAsStream(string fileName)
	{
		var path = Path.Combine(Environment.CurrentDirectory, $"Files\\{fileName}");
		var fi = new FileInfo(path);
		if (!fi.Exists) throw new FileNotFoundException();
		MemoryStream ms = new MemoryStream();
		using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
			file.CopyTo(ms);
		return ms;
	}
	public void SaveFileFromStream(MemoryStream content, string fileName)
	{
		DirectoryInfo? debugFolder = Directory.GetParent(Environment.CurrentDirectory);
		if (debugFolder is null) throw new Exception("debugFolder not found");
		var projectFolder = debugFolder.Parent?.Parent;
		if (projectFolder is null) throw new Exception("projectFolder not found");

		var path = Path.Combine(projectFolder.FullName, $"FileResults\\result-{fileName}");
		using (FileStream fs = new FileStream(path, FileMode.Create))
		{
			content.WriteTo(fs);
		}
	}


	public void GeneralResultChecks(WvSpreadsheetFileTemplateProcessResult? result)
	{
		Assert.NotNull(result);
		Assert.NotNull(result.Workbook);
		Assert.NotNull(result.ResultItems);
		Assert.True(result.ResultItems.Count > 0);

		Assert.NotNull(result.Workbook.Worksheets);
		Assert.True(result.Workbook.Worksheets.Count > 0);

		Assert.NotNull(result.ResultItems[0].Result);
		Assert.NotNull(result.ResultItems[0].Workbook);
		Assert.NotNull(result.ResultItems[0].Workbook!.Worksheets);
		Assert.True(result.ResultItems[0].Workbook!.Worksheets.Count > 0);
	}

	public void GeneralResultChecks(WvDocumentFileTemplateProcessResult? result)
	{
		Assert.NotNull(result);
		Assert.NotNull(result.WordDocument);
		Assert.NotNull(result.ResultItems);
		Assert.True(result.ResultItems.Count > 0);

		Assert.NotNull(result.WordDocument.MainDocumentPart);
		Assert.NotNull(result.WordDocument.MainDocumentPart.Document);
		Assert.NotNull(result.WordDocument.MainDocumentPart.Document.Body);

		Assert.NotNull(result.ResultItems[0].Result);
		Assert.NotNull(result.ResultItems[0].WordDocument);
		Assert.NotNull(result.ResultItems[0].WordDocument.MainDocumentPart);
		Assert.NotNull(result.ResultItems[0].WordDocument.MainDocumentPart.Document);
		Assert.NotNull(result.ResultItems[0].WordDocument.MainDocumentPart.Document.Body);
	}

	public void CheckRangeDimensions(IXLRange range, int startRowNumber, int startColumnNumber, int lastRowNumber, int lastColumnNumber)
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

	public void CheckCellPropertiesCopy(List<WvSpreadsheetFileTemplateContext> templateContexts,
		WvSpreadsheetFileTemplateProcessResultItem resultItem)
	{
		foreach (var tempContext in templateContexts)
		{
			var resultContext = resultItem.Contexts.FirstOrDefault(x => x.TemplateContextId == tempContext.Id);
			Assert.NotNull(resultContext);
			var firstTemplateCell = tempContext.Range?.Cell(1, 1);
			var firstResultCell = resultContext.Range?.Cell(1, 1);
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

	public void CompareStyle(IXLCell template, IXLCell result)
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

	public void CompareStyleAlignment(IXLAlignment template, IXLAlignment result)
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

	public void CompareStyleBorder(IXLBorder template, IXLBorder result)
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

	public void CompareStyleFormat(IXLNumberFormat template, IXLNumberFormat result)
	{
		Assert.Equal(template.NumberFormatId, result.NumberFormatId);
		Assert.Equal(template.Format, result.Format);
	}

	public void CompareStyleFill(IXLFill template, IXLFill result)
	{
		CompareStyleColor(template.BackgroundColor, result.BackgroundColor);
		CompareStyleColor(template.PatternColor, result.PatternColor);
		Assert.Equal(template.PatternType, result.PatternType);
	}

	public void CompareStyleFont(IXLFont template, IXLFont result)
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

	public void CompareProtection(IXLProtection template, IXLProtection result)
	{
		Assert.Equal(template.Locked, result.Locked);
		Assert.Equal(template.Hidden, result.Hidden);
	}

	public void CompareStyleColor(XLColor template, XLColor result)
	{
		if (template is null)
		{
			Assert.Null(result);
		}
		else
		{
			Assert.NotNull(result);
			if (template.ColorType == XLColorType.Color && result.ColorType == XLColorType.Color)
				Assert.Equal(template.ToString(), result.ToString());

			if (template.ColorType == XLColorType.Theme && result.ColorType == XLColorType.Theme)
				Assert.Equal(template.ToString(), result.ToString());
		}
	}

	public void CompareRowProperties(IXLRow template, IXLRow result)
	{
		Assert.Equal(template.OutlineLevel, result.OutlineLevel);
		Assert.Equal(template.Height, result.Height);
	}

	public void CompareColumnProperties(IXLColumn template, IXLColumn result)
	{
		Assert.Equal(template.OutlineLevel, result.OutlineLevel);
		Assert.Equal(template.Width, result.Width);
	}

	public void SaveWorkbook(XLWorkbook workbook, string fileName)
	{
		DirectoryInfo? debugFolder = Directory.GetParent(Environment.CurrentDirectory);
		if (debugFolder is null) throw new Exception("debugFolder not found");
		var projectFolder = debugFolder.Parent?.Parent;
		if (projectFolder is null) throw new Exception("projectFolder not found");

		var path = Path.Combine(projectFolder.FullName, $"FileResults\\result-{fileName}");
		workbook.SaveAs(path);
	}



}

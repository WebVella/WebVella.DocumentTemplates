using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using System.Data;
using System.Globalization;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Engines.Text;
using WebVella.DocumentTemplates.Extensions;
using System;
using System.Linq;
using WordText = DocumentFormat.OpenXml.Wordprocessing.Text;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using DocumentFormat.OpenXml.Spreadsheet;
using WebVella.DocumentTemplates.Engines.SpreadsheetFile;
using WebVella.DocumentTemplates.Engines.SpreadsheetFile.Utility;

namespace WebVella.DocumentTemplates.Engines.DocumentFile.Utility;
public partial class WvDocumentFileEngineUtility
{
	public void ProcessDocumentTemplate(WvDocumentFileTemplateProcessResult result,
		WvDocumentFileTemplateProcessResultItem resultItem, DataTable dataSource, CultureInfo culture)
	{
		//Validate
		if (resultItem is null) throw new Exception("No result provided!");
		if (dataSource is null) throw new Exception("No datasource provided!");
		if (result.Template is null) throw new Exception("No Template provided!");
		if (result.WordDocument is null) throw new Exception("No result.WordDocument provided!");
		if (result.WordDocument.MainDocumentPart is null) throw new Exception("No  result.WordDocument.MainDocumentPart provided!");
		if (resultItem.WordDocument is null) throw new Exception("No resultItem.WordDocument provided!");
		if (resultItem.WordDocument.MainDocumentPart is null)
		{
			var mainPart = resultItem.WordDocument.AddMainDocumentPart();
			mainPart.Document = new Document(new Body());
		}


		Body templateBody = result.WordDocument.MainDocumentPart.Document!.Body!;
		Body resultBody = resultItem.WordDocument.MainDocumentPart!.Document!.Body!;
		foreach (var childEl in templateBody.Elements())
		{
			var resultEl = _processDocumentElement(childEl, dataSource, culture);
			resultBody.Append(resultEl);
		}
	}

	private OpenXmlElement _processDocumentElement(
		OpenXmlElement template,
		DataTable dataSource, CultureInfo culture
		)
	{
		if (template.GetType().FullName == typeof(Paragraph).FullName)
			return _processDocumentParagarph((Paragraph)template, dataSource, culture);
		else if (template.GetType().FullName == typeof(Table).FullName)
			return _processDocumentTable((Table)template, dataSource, culture);
		return template.CloneNode(true);
	}

	private Paragraph _processDocumentParagarph(Paragraph template,
		DataTable dataSource, CultureInfo culture)
	{
		Paragraph resultEl = new Paragraph();

		//Copy properties
		ParagraphProperties? templateProperties = template.Elements<ParagraphProperties>().FirstOrDefault();
		if (templateProperties is not null)
		{
			resultEl.AppendChild(templateProperties.CloneNode(true));
		}

		if (!String.IsNullOrWhiteSpace(template.InnerText)
			&& resultEl.GetType().InheritsClass(typeof(OpenXmlCompositeElement)))
		{
			var textTemplate = new WvTextTemplate
			{
				Template = template.InnerText
			};
			var textTemplateResults = textTemplate.Process(dataSource, culture);
			if (textTemplateResults.ResultItems.Count == 0
				|| textTemplateResults.ResultItems[0].Result is null) return resultEl;

			var run = new Run(new WordText(textTemplateResults.ResultItems[0].Result!));
			RunProperties? templateRunProperties = template.Descendants<RunProperties>().FirstOrDefault();
			if (templateRunProperties is not null)
			{
				run.PrependChild(templateRunProperties.CloneNode(true));
			}
			((OpenXmlCompositeElement)resultEl).Append(run);

		}
		return resultEl;
	}

	private Table _processDocumentTable(Table template,
		DataTable dataSource, CultureInfo culture)
	{
		Table resultEl = new Table();

		TableProperties? templateProperties = template.Elements<TableProperties>().FirstOrDefault();
		if (templateProperties is not null)
		{
			resultEl.AppendChild(templateProperties.CloneNode(true));
		}

		TableGrid? templateGrid = template.Elements<TableGrid>().FirstOrDefault();
		if (templateGrid is not null)
		{
			resultEl.AppendChild(templateGrid.CloneNode(true));
		}

		var xlToTableRowContextDict = new Dictionary<int, TableRow>();
		var xlToTableCellContextDict = new Dictionary<string, TableCell>();
		var tempWb = new XLWorkbook();
		var tempWs = tempWb.Worksheets.Add();
		var tempRowPosition = 1;
		var tempColPosition = 1;
		foreach (var rowEl in template.Elements<TableRow>())
		{
			tempColPosition = 1;
			xlToTableRowContextDict[tempRowPosition] = rowEl;
			foreach (var cellEl in rowEl.Elements<TableCell>())
			{
				tempWs.Cell(tempRowPosition, tempColPosition).SetValue(cellEl.InnerText);
				xlToTableCellContextDict[$"{tempRowPosition}-{tempColPosition}"] = cellEl;
				tempColPosition++;
			}
			tempRowPosition++;
		}

		var templMS = new MemoryStream();
		tempWb.SaveAs(templMS);

		var spreadSheetTemplate = new WvSpreadsheetFileTemplate
		{
			Template = templMS
		};
		var spreadSheetTemplateResult = spreadSheetTemplate.Process(dataSource, culture);

		if (spreadSheetTemplateResult.ResultItems.Count == 0
			|| spreadSheetTemplateResult.ResultItems[0].Workbook is null
			|| spreadSheetTemplateResult.ResultItems[0].Workbook!.Worksheets!.Count == 0)
			return resultEl;

		var ws = spreadSheetTemplateResult.ResultItems[0].Workbook!.Worksheets.First();

		var testRowProps = xlToTableRowContextDict[1];
		var testCellProps = xlToTableCellContextDict[$"{1}-{1}"];

		for (var rowPosition = 1; rowPosition <= ws.LastRowUsed()!.RowNumber(); rowPosition++)
		{
			var rowEl = new TableRow();
			var rowProps = testCellProps.Descendants<TableRowProperties>().FirstOrDefault();
			if (rowProps is not null)
				rowEl.PrependChild(rowProps.CloneNode(true));

			for (var colPosition = 1; colPosition <= ws.LastColumnUsed()!.ColumnNumber(); colPosition++)
			{
				var cellEl = new TableCell();
				var cellProps = testCellProps.Descendants<TableCellProperties>().FirstOrDefault();
				if (cellProps is not null)
					cellEl.PrependChild(cellProps.CloneNode(true));

				var cell = ws.Cell(rowPosition, colPosition);
				Run run = new Run(new WordText(cell.Value.ToString()));
				var paragraph = new Paragraph();
				var cellParagraph = testCellProps.Descendants<Paragraph>().FirstOrDefault();
				if (cellParagraph is not null)
				{
					var cellParagraphProperties = cellParagraph.Descendants<ParagraphProperties>().FirstOrDefault();
					if (cellParagraphProperties is not null)
						paragraph.PrependChild(cellParagraphProperties.CloneNode(true));
				}
				paragraph.AppendChild(run);
				cellEl.AppendChild(paragraph);

				rowEl.AppendChild(cellEl);
			}
			resultEl.AppendChild(rowEl);
		}


		return resultEl;
	}

	private List<TableRow> _processDocumentTableRow(TableRow row,
		DataTable dataSource, CultureInfo culture)
	{
		var result = new List<TableRow>();

		return result;
	}
}

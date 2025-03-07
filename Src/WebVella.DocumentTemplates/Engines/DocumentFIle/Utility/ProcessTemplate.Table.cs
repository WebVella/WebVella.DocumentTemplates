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
using Word = DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;
using WebVella.DocumentTemplates.Engines.SpreadsheetFile;
using WebVella.DocumentTemplates.Engines.SpreadsheetFile.Utility;
using System.Text;

namespace WebVella.DocumentTemplates.Engines.DocumentFile.Utility;
public partial class WvDocumentFileEngineUtility
{
private Word.Table _processDocumentTable(Word.Table template,
		DataTable dataSource, CultureInfo culture)
	{
		Word.Table resultEl = new Word.Table();

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
				Word.Run run = new Word.Run(new Word.Text(cell.Value.ToString()));
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
}

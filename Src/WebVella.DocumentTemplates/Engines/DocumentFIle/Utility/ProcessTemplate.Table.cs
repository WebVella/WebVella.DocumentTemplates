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
using WebVella.DocumentTemplates.Engines.SpreadsheetFile.Models;

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


		var testCellProps = xlToTableCellContextDict[$"{1}-{1}"];

		for (var rowPosition = 1; rowPosition <= ws.LastRowUsed()!.RowNumber(); rowPosition++)
		{
			var rowEl = createTableRowEl(rowPosition, xlToTableRowContextDict, spreadSheetTemplateResult);
			for (var colPosition = 1; colPosition <= ws.LastColumnUsed()!.ColumnNumber(); colPosition++)
			{
				var cell = ws.Cell(rowPosition, colPosition);
				var cellEl = createTableCellEl(rowPosition, colPosition, xlToTableCellContextDict,
					spreadSheetTemplateResult, cell.Value.ToString());
				rowEl.AppendChild(cellEl);
			}
			resultEl.AppendChild(rowEl);
		}


		return resultEl;
	}

	private Word.TableRow createTableRowEl(int rowPosition,
	Dictionary<int, TableRow> originalRowDict,
	WvSpreadsheetFileTemplateProcessResult wsResult)
	{
		var range = new WvSpreadsheetRange
		{
			FirstRow = rowPosition,
			LastRow = rowPosition,
			FirstColumn = 1,
			LastColumn = 1,
		};
		var wsResultRows = wsResult.ResultItems[0].Contexts.GetByAddress(rowPosition, 1);
		if(wsResultRows.Count == 0) return new Word.TableRow();
		var wsResultRow = wsResultRows.First();
		var wsTemplateRow = wsResult.TemplateContexts.Single(x=> x.Id == wsResultRow.TemplateContextId);
		var wsRow = wsTemplateRow.Range?.RangeAddress.FirstAddress.RowNumber ?? 1;
		if (!originalRowDict.ContainsKey(wsRow)) return new Word.TableRow();

		var originalRow = originalRowDict[wsRow];
		var rowEl = (Word.TableRow)originalRow.CloneNode(true);
		rowEl.RemoveAllChildren<Word.TableCell>();
		return rowEl;
	}

	private Word.TableCell createTableCellEl(int rowPosition, int columnPosition,
	Dictionary<string, TableCell> originalCellDict,
	WvSpreadsheetFileTemplateProcessResult wsResult, string value)
	{
		var wsResultRows = wsResult.ResultItems[0].Contexts.GetByAddress(rowPosition, columnPosition);
		if(wsResultRows.Count == 0) return new Word.TableCell();
		var wsResultRow = wsResultRows.First();
		var wsTemplateRow = wsResult.TemplateContexts.Single(x=> x.Id == wsResultRow.TemplateContextId);
		var wsRow = wsTemplateRow.Range?.RangeAddress.FirstAddress.RowNumber ?? 1;
		var wsCol = wsTemplateRow.Range?.RangeAddress.FirstAddress.ColumnNumber ?? 1;
		var dictKey = $"{wsRow}-{wsCol}";
		if (!originalCellDict.ContainsKey(dictKey)) return new Word.TableCell();
		var originalCell = originalCellDict[dictKey];
		var cellEl = (Word.TableCell)originalCell.CloneNode(true);
		cellEl.RemoveAllChildren<Word.Paragraph>();

		var cellParagraph = new Paragraph(new Word.Run(new Word.Text(value)));
		var originalCellParagraph = originalCell.Descendants<Word.Paragraph>().FirstOrDefault();
		if (originalCellParagraph is not null)
		{
			cellParagraph = (Word.Paragraph)originalCellParagraph.CloneNode(true);
			cellParagraph.RemoveAllChildren<Word.Run>();

			var originalCellRun = originalCellParagraph.Descendants<Word.Run>().FirstOrDefault();
			if (originalCellRun is not null)
			{
				var paragraphRun = (Word.Run)originalCellRun.CloneNode(true);
				paragraphRun.RemoveAllChildren<Word.Text>();
				paragraphRun.AppendChild(new Word.Text(value));
				cellParagraph.AppendChild(paragraphRun);
			}
			else
			{
				cellParagraph.AppendChild(new Word.Run(new Word.Text(value)));
			}
		}
		cellEl.AppendChild(cellParagraph);

		return cellEl;
	}
}

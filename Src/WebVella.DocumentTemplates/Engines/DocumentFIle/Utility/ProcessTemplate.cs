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
		if (result.WordDocument.MainDocumentPart.Document is null) throw new Exception("No  result.WordDocument.MainDocumentPart.Document provided!");
		if (result.WordDocument.MainDocumentPart.Document!.Body is null) throw new Exception("No  result.WordDocument.MainDocumentPart.Document.Body provided!");
		if (resultItem.WordDocument is null) throw new Exception("No resultItem.WordDocument provided!");
		if (resultItem.WordDocument.MainDocumentPart is null)
		{
			var mainPart = resultItem.WordDocument.AddMainDocumentPart();
			mainPart.Document = new Document(new Word.Body());
		}

		//Process Body
		var templateBody = result.WordDocument.MainDocumentPart.Document!.Body!;
		var resultBody = resultItem.WordDocument.MainDocumentPart!.Document!.Body!;
		foreach (var childEl in templateBody.ChildElements)
		{
			var resultChiledEl = _processDocumentElement(childEl, dataSource, culture);
			if (resultChiledEl is not null)
				resultBody.AppendChild(resultChiledEl);
		}

		_copyDocumentStylesAndSettings(result.WordDocument, resultItem.WordDocument);
		_copyHeadersAndFooters(result.WordDocument, resultItem.WordDocument);
		_copyFootnotesAndEndnotes(result.WordDocument, resultItem.WordDocument);
		_copyImages(result.WordDocument, resultItem.WordDocument);
	}

}

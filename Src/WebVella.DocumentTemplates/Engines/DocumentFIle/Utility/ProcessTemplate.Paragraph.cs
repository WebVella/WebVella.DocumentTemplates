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
	private Word.Paragraph _processDocumentParagaraph(Word.Paragraph template,
		DataTable dataSource, CultureInfo culture)
	{
		Word.Paragraph resultEl = (Word.Paragraph)template.CloneNode(true);
		if (String.IsNullOrWhiteSpace(template.InnerText)) return resultEl;
		resultEl.RemoveAllChildren();

		//Runs Need to be preprocessed as there is often a problem where the tag is 
		//delivered in different runs
		var queue = new Queue<Word.Run>();
		
		foreach (var childEl in template.ChildElements)
		{
			
			var resultChildEl =  _processDocumentElement(childEl, dataSource, culture);
			if(resultChildEl is not null)
				resultEl.AppendChild(resultChildEl);
		}
		return resultEl;
	}

}

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
	private Word.Text _processDocumentText(Word.Text template,
	DataTable dataSource, CultureInfo culture)
	{
		Word.Text resultEl = (Word.Text)template.CloneNode(true);
		resultEl.Text = String.Empty;

		if (String.IsNullOrWhiteSpace(template.InnerText)) return resultEl;

		var textTemplate = new WvTextTemplate
		{
			Template = template.InnerText
		};
		var textTemplateResults = textTemplate.Process(dataSource, culture);
		var sb = new StringBuilder();
		foreach (var item in textTemplateResults.ResultItems)
		{
			sb.Append(item.Result ?? String.Empty);
		}
		resultEl.Text = sb.ToString();
		return resultEl;
	}
}

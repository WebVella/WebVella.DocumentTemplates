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
	private OpenXmlElement? _processDocumentElement(
		OpenXmlElement template,
		DataTable dataSource, CultureInfo culture
		)
	{
		if (template.GetType().FullName == typeof(Paragraph).FullName)
			return _processDocumentParagaraph((Paragraph)template, dataSource, culture);

		if (template.GetType().FullName == typeof(Word.Run).FullName)
			return _processDocumentRun((Word.Run)template, dataSource, culture);

		if (template.GetType().FullName == typeof(Word.Table).FullName)
			return _processDocumentTable((Word.Table)template, dataSource, culture);

		if (template.GetType().FullName == typeof(Word.TableCell).FullName)
			return _processDocumentTableCell((Word.TableCell)template, dataSource, culture);

		if (template.GetType().FullName == typeof(Word.TableRow).FullName)
			return _processDocumentTableRow((Word.TableRow)template, dataSource, culture);

		if (template.GetType().FullName == typeof(Word.Text).FullName)
			return _processDocumentText((Word.Text)template, dataSource, culture);

		if (template.GetType().FullName == typeof(Word.Comment).FullName)
			return null;

		if (template.GetType().FullName == typeof(Word.ProofError).FullName)
			return null;

		if (template.GetType().FullName == typeof(Word.SectionProperties).FullName)
			return null;

		if (template.GetType().FullName == typeof(Word.ParagraphProperties).FullName)
			return template.CloneNode(true);

		if (template.GetType().FullName == typeof(Word.RunProperties).FullName)
			return template.CloneNode(true);

		return template.CloneNode(true);
	}
}

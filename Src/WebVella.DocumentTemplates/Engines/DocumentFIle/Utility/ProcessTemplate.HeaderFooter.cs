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
	private void _copyHeadersAndFooters(WordprocessingDocument original, WordprocessingDocument target)
	{
		var originalMainPart = original.MainDocumentPart;
		var targetMainPart = target.MainDocumentPart;
		if (originalMainPart == null || targetMainPart == null)
			return;

		var body = targetMainPart.Document.Body!;
		var originalSectPr = originalMainPart.Document.Body!.Elements<SectionProperties>().LastOrDefault();
		if (originalSectPr == null)
			return;

		var targetSectPr = body.Elements<SectionProperties>().LastOrDefault();
		if (targetSectPr == null)
		{
			targetSectPr = new SectionProperties();
			body.Append(targetSectPr);
		}

		foreach (var headerRef in originalSectPr.Elements<HeaderReference>())
		{
			if (headerRef.Id is null) continue;
			var originalHeader = originalMainPart.GetPartById(headerRef.Id!) as HeaderPart;
			if (originalHeader is null) continue;
			var targetHeader =  _getOrCreateHeader(targetMainPart);
			_copyPartContent(originalHeader, targetHeader);

			var newHeaderReference = new HeaderReference
			{
				Type = headerRef.Type,
				Id = targetMainPart.GetIdOfPart(targetHeader)
			};
			targetSectPr.AppendChild(newHeaderReference);
		}

		foreach (var footerRef in originalSectPr.Elements<FooterReference>())
		{
			if (footerRef.Id is null) continue;
			var originalFooter = originalMainPart.GetPartById(footerRef.Id!) as FooterPart;
			if (originalFooter is null) continue;
			var targetFooter = _getOrCreateFooter(targetMainPart);
			_copyPartContent(originalFooter, targetFooter);

			var newFooterReference = new FooterReference
			{
				Type = footerRef.Type,
				Id = targetMainPart.GetIdOfPart(targetFooter)
			};
			targetSectPr.AppendChild(newFooterReference);
		}
	}
	private void _copyFootnotesAndEndnotes(WordprocessingDocument original, WordprocessingDocument target)
	{
		var originalMainPart = original.MainDocumentPart;
		var targetMainPart = target.MainDocumentPart;
		if (originalMainPart == null || targetMainPart == null)
			return;

		if (originalMainPart.FootnotesPart != null)
		{
			if (targetMainPart.FootnotesPart == null)
				targetMainPart.AddNewPart<FootnotesPart>();

			_copyPartContent(originalMainPart.FootnotesPart, targetMainPart.FootnotesPart!);
		}

		if (originalMainPart.EndnotesPart != null)
		{
			if (targetMainPart.EndnotesPart == null)
				targetMainPart.AddNewPart<EndnotesPart>();

			_copyPartContent(originalMainPart.EndnotesPart, targetMainPart.EndnotesPart!);
		}
	}
	private void _copyPartContent(OpenXmlPart originalPart, OpenXmlPart targetPart)
	{
		using (Stream sourceStream = originalPart.GetStream(FileMode.Open, FileAccess.Read))
		using (Stream targetStream = targetPart.GetStream(FileMode.Create, FileAccess.Write))
		{
			sourceStream.CopyTo(targetStream);
		}
	}
}

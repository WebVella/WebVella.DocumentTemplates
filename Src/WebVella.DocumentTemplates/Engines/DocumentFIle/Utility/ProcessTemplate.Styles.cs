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
	public void _copyDocumentStylesAndSettings(WordprocessingDocument original, WordprocessingDocument target)
	{
		if (original == null || target == null)
			throw new ArgumentNullException("Both documents must be provided.");
		_copyOrCreatePart(original, target, "styles.xml");
		_copyOrCreatePart(original, target, "numbering.xml");
		_copyOrCreatePart(original, target, "theme/theme1.xml");
		_copyOrCreatePart(original, target, "settings.xml");
	}

	private void _copyOrCreatePart(WordprocessingDocument original, WordprocessingDocument target, string partPath)
	{
		var originalPart = _getPartByUri(original, partPath);
		if (originalPart == null) return;

		var targetPart = _getPartByUri(target, partPath) ?? _createPart(target, partPath);
		if(targetPart is null) return;

		using (Stream sourceStream = originalPart.GetStream(FileMode.Open, FileAccess.Read))
		using (Stream targetStream = targetPart.GetStream(FileMode.Create, FileAccess.Write))
		{
			sourceStream.CopyTo(targetStream);
		}
	}

	private OpenXmlPart _getPartByUri(WordprocessingDocument doc, string partPath)
	{
		return doc.MainDocumentPart!.Parts
			.FirstOrDefault(p => p.OpenXmlPart.Uri.ToString().EndsWith(partPath)).OpenXmlPart;
	}

	private OpenXmlPart? _createPart(WordprocessingDocument target, string partPath)
	{
		var mainPart = target.MainDocumentPart;
		if (mainPart == null)
			throw new InvalidOperationException("Target document does not have a MainDocumentPart.");
		switch (partPath)
		{
			case "styles.xml":
				return mainPart.AddNewPart<StyleDefinitionsPart>();
			case "numbering.xml":
				return mainPart.AddNewPart<NumberingDefinitionsPart>();
			case "theme/theme1.xml":
				return mainPart.AddNewPart<ThemePart>();
			case "settings.xml":
				return mainPart.AddNewPart<DocumentSettingsPart>();
			default:
				return null;

		}
	}
}

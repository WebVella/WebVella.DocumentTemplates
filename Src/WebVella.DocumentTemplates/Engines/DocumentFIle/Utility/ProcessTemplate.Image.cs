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
using DocumentFormat.OpenXml.Drawing;


namespace WebVella.DocumentTemplates.Engines.DocumentFile.Utility;
public partial class WvDocumentFileEngineUtility
{
	public void _copyImages(WordprocessingDocument original, WordprocessingDocument target)
	{
		var originalMainPart = original.MainDocumentPart;
		var targetMainPart = target.MainDocumentPart;

		if (originalMainPart == null || targetMainPart == null)
			throw new InvalidOperationException("Both documents must have a MainDocumentPart.");
		Dictionary<string, string> imageIdMap = new Dictionary<string, string>();
		_copyImageParts(originalMainPart, targetMainPart, imageIdMap);
		_copyImagesFromHeaderFooter(originalMainPart, targetMainPart, imageIdMap);
		_updateBlipReferences(targetMainPart, imageIdMap);
	}

	private void _copyImageParts(MainDocumentPart originalMainPart, MainDocumentPart targetMainPart, Dictionary<string, string> imageIdMap)
	{
		foreach (var imagePart in originalMainPart.ImageParts)
		{
			string originalImageId = originalMainPart.GetIdOfPart(imagePart);
			ImagePart newImagePart = targetMainPart.AddImagePart(imagePart.ContentType);
			var inputStream = imagePart.GetStream();
			var outputStream = newImagePart.GetStream();
			_copyStream(inputStream, outputStream);
			inputStream.Close();
			outputStream.Close();
			string newImageId = targetMainPart.GetIdOfPart(newImagePart);
			imageIdMap[originalImageId] = newImageId;
		}
	}

	private void _copyImagesFromHeaderFooter(MainDocumentPart originalMainPart, MainDocumentPart targetMainPart, Dictionary<string, string> imageIdMap)
	{
		foreach (var originalHeader in originalMainPart.HeaderParts)
		{
			var targetHeader = _getOrCreateHeader(targetMainPart);
			_copyHeaderImages(originalHeader, targetHeader, imageIdMap);
		}

		foreach (var originalFooter in originalMainPart.FooterParts)
		{
			var targetFooter = _getOrCreateFooter(targetMainPart);
			_copyFooterImages(originalFooter, targetFooter, imageIdMap);
		}
	}

	private HeaderPart _getOrCreateHeader(MainDocumentPart targetMainPart)
	{
		return targetMainPart.HeaderParts.FirstOrDefault() ?? targetMainPart.AddNewPart<HeaderPart>();
	}

	private FooterPart _getOrCreateFooter(MainDocumentPart targetMainPart)
	{
		return targetMainPart.FooterParts.FirstOrDefault() ?? targetMainPart.AddNewPart<FooterPart>();
	}

	private void _copyHeaderImages(HeaderPart originalPart, HeaderPart targetPart, Dictionary<string, string> imageIdMap)
	{
		foreach (var imagePart in originalPart.Parts.Select(p => p.OpenXmlPart).OfType<ImagePart>())
		{
			string originalImageId = originalPart.GetIdOfPart(imagePart);
			ImagePart newImagePart = targetPart.AddImagePart(imagePart.ContentType);
			var inputStream = imagePart.GetStream();
			var outputStream = newImagePart.GetStream();
			_copyStream(inputStream, outputStream);
			inputStream.Close();
			outputStream.Close();
			string newImageId = targetPart.GetIdOfPart(newImagePart);
			imageIdMap[originalImageId] = newImageId;
		}
	}

	private void _copyFooterImages(FooterPart originalPart, FooterPart targetPart, Dictionary<string, string> imageIdMap)
	{
		foreach (var imagePart in originalPart.Parts.Select(p => p.OpenXmlPart).OfType<ImagePart>())
		{
			string originalImageId = originalPart.GetIdOfPart(imagePart);
			ImagePart newImagePart = targetPart.AddImagePart(imagePart.ContentType);
			var inputStream = imagePart.GetStream();
			var outputStream = newImagePart.GetStream();
			_copyStream(inputStream, outputStream);
			inputStream.Close();
			outputStream.Close();
			string newImageId = targetPart.GetIdOfPart(newImagePart);
			imageIdMap[originalImageId] = newImageId;
		}
		foreach (var drawing in targetPart.Footer.Descendants<Word.Drawing>())
		{
			var blip = drawing.Descendants<Blip>().FirstOrDefault();
			if (blip != null && blip.Embed != null && imageIdMap.ContainsKey(blip.Embed.Value))
			{
				blip.Embed.Value = imageIdMap[blip.Embed.Value];
			}
		}
	}

	private void _updateBlipReferences(MainDocumentPart targetMainPart, Dictionary<string, string> imageIdMap)
	{
		foreach (var targetHeader in targetMainPart.HeaderParts)
		{
			foreach (var drawing in targetHeader.Header.Descendants<Word.Drawing>())
			{
				var blip = drawing.Descendants<Blip>().FirstOrDefault();
				if (blip != null && blip.Embed != null && imageIdMap.ContainsKey(blip.Embed.Value))
				{
					blip.Embed.Value = imageIdMap[blip.Embed.Value];
				}
			}
		}
		foreach (var targetFooter in targetMainPart.FooterParts)
		{
			foreach (var drawing in targetFooter.Footer.Descendants<Word.Drawing>())
			{
				var blip = drawing.Descendants<Blip>().FirstOrDefault();
				if (blip != null && blip.Embed != null && imageIdMap.ContainsKey(blip.Embed.Value))
				{
					blip.Embed.Value = imageIdMap[blip.Embed.Value];
				}
			}
		}
		foreach (var drawing in targetMainPart.Document.Body.Descendants<Word.Drawing>())
		{
			var blip = drawing.Descendants<Blip>().FirstOrDefault();
			if (blip != null && blip.Embed != null && imageIdMap.ContainsKey(blip.Embed.Value))
			{
				blip.Embed.Value = imageIdMap[blip.Embed.Value];
			}
		}
	}

	private void _copyStream(Stream input, Stream output)
	{
		input.CopyTo(output);
		output.Flush();
	}
}

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
		OpenXmlElement resultEl = _copyDocumentElement(template);
		if (!String.IsNullOrWhiteSpace(template.InnerText) 
			&& resultEl.GetType().InheritsClass(typeof(OpenXmlCompositeElement)))
		{
			var textTemplate = new WvTextTemplate
			{
				Template = template.InnerText
			};
			var textTemplateResults = textTemplate.Process(dataSource, culture);
			if(textTemplateResults.ResultItems.Count == 0
				|| textTemplateResults.ResultItems[0].Result is null){ 
				((OpenXmlCompositeElement)resultEl).Append(new Run(new WordText(String.Empty)));
			}
			else{ 
				((OpenXmlCompositeElement)resultEl).Append(new Run(new WordText(textTemplateResults.ResultItems[0].Result!)));
			}
		}


		foreach (OpenXmlElement childEl in template.ChildElements)
		{
			var resultChildEl = _processDocumentElement(childEl, dataSource, culture);
			if (resultChildEl is null) continue;

			resultEl!.ChildElements.Append(resultChildEl);
		}

		return resultEl;
	}

	private OpenXmlElement _copyDocumentElement(OpenXmlElement original)
	{
		OpenXmlElement? result = original.CloneNode(false);

		var copiedTypes = new List<Type>{
				typeof(Run),
				typeof(ParagraphProperties)
			};
		foreach (var childElement in original.ChildElements)
		{

			if (!copiedTypes.Any(x => x.FullName == childElement.GetType().FullName))
				continue;

			if (childElement.GetType().FullName == typeof(ParagraphProperties).FullName)
			{
				var clone = (ParagraphProperties)childElement.CloneNode(false);
				result.AppendChild(clone);
			}

		}
		return result;
	}
}

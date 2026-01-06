using DocumentFormat.OpenXml.Wordprocessing;
using System.Data;
using System.Globalization;
using Word = DocumentFormat.OpenXml.Wordprocessing;


namespace WebVella.DocumentTemplates.Engines.DocumentFile.Utility;
public partial class WvDocumentFileEngineUtility
{
	public void ProcessDocumentTemplate(WvDocumentFileTemplateProcessResult result,
		WvDocumentFileTemplateProcessResultItem resultItem, DataTable dataSource, 
		CultureInfo culture,
		Dictionary<string,WvDocumentFileTemplate> templateLibrary,
		int stackLevel = 0)
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
			var resultChiledElList = _processDocumentElement(
				templateEl:childEl,
				dataSource:dataSource,
				culture:culture,
				templateLibrary:templateLibrary, 
				stackLevel:stackLevel);
			foreach (var resultChiledEl in resultChiledElList)
				resultBody.AppendChild(resultChiledEl);
		}

		_copyDocumentStylesAndSettings(result.WordDocument, resultItem.WordDocument);
		_copyHeadersAndFooters(result.WordDocument, resultItem.WordDocument,dataSource,culture);
		_copyFootnotesAndEndnotes(result.WordDocument, resultItem.WordDocument);
		_copyImages(result.WordDocument, resultItem.WordDocument);
	}

}

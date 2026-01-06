using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Data;
using System.Globalization;
using Word = DocumentFormat.OpenXml.Wordprocessing;


namespace WebVella.DocumentTemplates.Engines.DocumentFile.Utility;
public partial class WvDocumentFileEngineUtility
{
	private List<OpenXmlElement> _processDocumentElement(
		OpenXmlElement templateEl,
		DataTable dataSource, CultureInfo culture,
		Dictionary<string,WvDocumentFileTemplate> templateLibrary,
		int stackLevel)
	{
		if (templateEl.GetType().FullName == typeof(Paragraph).FullName)
			return _processDocumentParagaraph((Paragraph)templateEl, dataSource, culture,templateLibrary, stackLevel);

		if (templateEl.GetType().FullName == typeof(Word.Run).FullName)
			return _processDocumentRun((Word.Run)templateEl, dataSource, culture,templateLibrary, stackLevel);

		if (templateEl.GetType().FullName == typeof(Word.Table).FullName)
			return _processDocumentTable((Word.Table)templateEl, dataSource, culture,templateLibrary, stackLevel);

		if (templateEl.GetType().FullName == typeof(Word.TableCell).FullName)
			return new List<OpenXmlElement>(); //Table is processed as a whole

		if (templateEl.GetType().FullName == typeof(Word.TableRow).FullName)
			return new List<OpenXmlElement>(); //Table is processed as a whole

		if (templateEl.GetType().FullName == typeof(Word.Text).FullName)
			return _processDocumentText((Word.Text)templateEl, dataSource, culture,templateLibrary, stackLevel);

		if (templateEl.GetType().FullName == typeof(Word.Comment).FullName)
			return new List<OpenXmlElement>();

		if (templateEl.GetType().FullName == typeof(Word.ProofError).FullName)
			return new List<OpenXmlElement>();

		if (templateEl.GetType().FullName == typeof(Word.SectionProperties).FullName)
			return new List<OpenXmlElement>();

		// if (templateEl.GetType().FullName == typeof(Word.ParagraphProperties).FullName)
		// 	return [templateEl.CloneNode(true)];
	
		// if (templateEl.GetType().FullName == typeof(Word.RunProperties).FullName)
		// 	return [templateEl.CloneNode(true)];

		return [templateEl.CloneNode(true)];		
	}


	public List<OpenXmlElement> _preprocessChildElementsForSplittedTags(List<OpenXmlElement> childElements)
	{
		var result = new List<OpenXmlElement>();


		return result;
	}
}

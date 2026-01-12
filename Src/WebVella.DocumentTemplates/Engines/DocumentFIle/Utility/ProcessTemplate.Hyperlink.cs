using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.Data;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using WebVella.DocumentTemplates.Engines.Text;
using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace WebVella.DocumentTemplates.Engines.DocumentFile.Utility;
public partial class WvDocumentFileEngineUtility
{
	private List<OpenXmlElement> _processDocumentHyperlink(Word.Hyperlink template,
		DataTable dataSource, CultureInfo culture)
	{
        Word.Hyperlink resultEl = (Word.Hyperlink)template.CloneNode(true);
        if (String.IsNullOrWhiteSpace(template.InnerText)) return [resultEl];
        resultEl.RemoveAllChildren();
        foreach (var childEl in template.ChildElements)
        {
            var resultChildElList = _processDocumentElement(childEl, dataSource, culture);
            foreach (var resultChildEl in resultChildElList)
                resultEl.AppendChild(resultChildEl);
        }

        return [resultEl];

    }

    private void _copyHyperlinks(WordprocessingDocument original, WordprocessingDocument target)
    {
        var originalMainPart = original.MainDocumentPart;
        var targetMainPart = target.MainDocumentPart;

        // Get the relationship IDs for hyperlinks in the source
        var sourceRelIds = originalMainPart!.HyperlinkRelationships.Select(r => r.Id).ToList();

        // Copy each hyperlink relationship from source to target
        foreach (var relId in sourceRelIds)
        {
            // Get the relationship from source
            HyperlinkRelationship? sourceRel = originalMainPart!.HyperlinkRelationships.FirstOrDefault(x => x.Id == relId);
            if (sourceRel is null) continue;
            var sourceTargetUri = sourceRel!.Uri;

            // Add the same relationship to the target document
            targetMainPart!.AddHyperlinkRelationship(sourceTargetUri, sourceRel.IsExternal, sourceRel.Id);
        }
    }
}

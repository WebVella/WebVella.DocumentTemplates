using DocumentFormat.OpenXml.Packaging;
using Word = DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing;


namespace WebVella.DocumentTemplates.Engines.DocumentFile.Utility;

public partial class WvDocumentFileEngineUtility
{
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
            HyperlinkRelationship? sourceRel = originalMainPart!.HyperlinkRelationships.FirstOrDefault(x=> x.Id == relId);
            if(sourceRel is null) continue;
            var sourceTargetUri = sourceRel!.Uri;

            // Add the same relationship to the target document
            targetMainPart!.AddHyperlinkRelationship(sourceTargetUri, sourceRel.IsExternal, sourceRel.Id);
        }
    }

   
}
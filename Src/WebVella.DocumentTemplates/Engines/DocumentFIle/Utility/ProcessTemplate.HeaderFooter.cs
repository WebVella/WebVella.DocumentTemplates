using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Data;
using System.Globalization;

namespace WebVella.DocumentTemplates.Engines.DocumentFile.Utility;

public partial class WvDocumentFileEngineUtility
{
    private void _copyHeadersAndFooters(WordprocessingDocument original, WordprocessingDocument target,
        DataTable dataSource, CultureInfo culture)
    {
        var originalMainPart = original.MainDocumentPart;
        var targetMainPart = target.MainDocumentPart;
        if (originalMainPart == null || targetMainPart == null)
            return;

        var targetBody = targetMainPart.Document.Body!;
        var originalSectPr = originalMainPart.Document.Body!.Elements<SectionProperties>().LastOrDefault();
        if (originalSectPr == null)
            return;

        targetBody.RemoveAllChildren<SectionProperties>();
        var targetSectPr = originalSectPr.CloneNode(true);
        targetSectPr.RemoveAllChildren<HeaderReference>();
        targetSectPr.RemoveAllChildren<FooterReference>();
        targetBody.Append(targetSectPr);

        #region << Footer >>
        foreach (var originalFooter in originalMainPart.FooterParts.OrderBy(x => x.Uri.ToString()))
        {
            var targetFooter = targetMainPart.AddNewPart<FooterPart>();
            _copyPartContent(originalFooter, targetFooter);

            List<OpenXmlElement> processed = new();
            foreach (var footerEl in originalFooter.Footer.ChildElements)
            {
                var procElList = _processDocumentElement(
                    templateEl: footerEl,
                    dataSource: dataSource,
                    culture: culture
                    );
                if (procElList.Count > 0) processed.AddRange(procElList);
            }
            targetFooter.Footer.RemoveAllChildren();
            foreach (var procEl in processed)
            {
                targetFooter.Footer.AppendChild(procEl);
            }
        }
        var generatedFooterReferences = new List<FooterReference>();
        foreach (var headerRef in originalSectPr.Elements<FooterReference>())
        {
            if (headerRef.Id is null) continue;
            var originalFooter = originalMainPart.GetPartById(headerRef.Id.Value!);
            var targetFooter = targetMainPart.FooterParts
                .FirstOrDefault(x => x.Uri == originalFooter.Uri);
            if (targetFooter is null) continue;
            var newFooterReference = new FooterReference
            {
                Type = headerRef.Type,
                Id = targetMainPart.GetIdOfPart(targetFooter!)
            };
            generatedFooterReferences.Add(newFooterReference);
        }
        if (generatedFooterReferences.Count > 0)
        {
            var currentChildElements = targetSectPr.ChildElements.ToList();
            targetSectPr.RemoveAllChildren();
            foreach (var currentChild in generatedFooterReferences)
            {
                targetSectPr.AppendChild(currentChild);
            }

            foreach (var currentChild in currentChildElements)
            {
                targetSectPr.AppendChild(currentChild);
            }
        }

        #endregion
        #region << Header >>
        foreach (var originalHeader in originalMainPart.HeaderParts.OrderBy(x=> x.Uri.ToString()))
        {
            var targetHeader = targetMainPart.AddNewPart<HeaderPart>();
            _copyPartContent(originalHeader, targetHeader);

            List<OpenXmlElement> processed = new();
            foreach (var headerEl in originalHeader.Header.ChildElements)
            {
                var procElList = _processDocumentElement(
                    templateEl: headerEl,
                    dataSource: dataSource,
                    culture: culture);
                if (procElList.Count > 0) processed.AddRange(procElList);
            }
            targetHeader.Header.RemoveAllChildren();
            foreach (var procEl in processed)
            {
                targetHeader.Header.AppendChild(procEl);
            }
        }
        var generatedHeaderReferences = new List<HeaderReference>();
        foreach (var headerRef in originalSectPr.Elements<HeaderReference>())
        {
            if (headerRef.Id is null) continue;

            var originalHeader = originalMainPart.GetPartById(headerRef.Id.Value!);
            var targetHeader = targetMainPart.HeaderParts
                .FirstOrDefault(x => x.Uri == originalHeader.Uri);
            if (targetHeader is null) continue;
            var newHeaderReference = new HeaderReference
            {
                Type = headerRef.Type,
                Id = targetMainPart.GetIdOfPart(targetHeader!)
            };
            generatedHeaderReferences.Add(newHeaderReference);

        }
        if (generatedHeaderReferences.Count > 0)
        {
            var currentChildElements = targetSectPr.ChildElements.ToList();
            targetSectPr.RemoveAllChildren();
            foreach (var currentChild in generatedHeaderReferences)
            {
                targetSectPr.AppendChild(currentChild);
            }

            foreach (var currentChild in currentChildElements)
            {
                targetSectPr.AppendChild(currentChild);
            }
        }

        #endregion

    }
    private void _copyHeadersAndFootersHyperlinks(WordprocessingDocument original, WordprocessingDocument target)
    {
        var originalMainPart = original.MainDocumentPart;
        var targetMainPart = target.MainDocumentPart;
        #region << Header >>
        {
            var originalParts = originalMainPart.HeaderParts.ToList();
            var targetParts = targetMainPart.HeaderParts.ToList();
            foreach (var originalPart in originalParts)
            {
                var targetPart = targetParts.FirstOrDefault(x => x.Uri == originalPart.Uri);
                if (targetPart is null) continue;
                // Get the relationship IDs for hyperlinks in the source
                var sourceRelIds = originalPart!.HyperlinkRelationships.Select(r => r.Id).ToList();

                // Copy each hyperlink relationship from source to target
                foreach (var relId in sourceRelIds)
                {
                    // Get the relationship from source
                    HyperlinkRelationship? sourceRel = originalPart!.HyperlinkRelationships.FirstOrDefault(x => x.Id == relId);
                    if (sourceRel is null) continue;
                    var sourceTargetUri = sourceRel!.Uri;

                    // Add the same relationship to the target document
                    targetPart!.AddHyperlinkRelationship(sourceTargetUri, sourceRel.IsExternal, sourceRel.Id);
                }
            }
        }
        #endregion
        #region << Footer >>
        {
            var originalParts = originalMainPart.FooterParts.ToList();
            var targetParts = targetMainPart.FooterParts.ToList();
            foreach (var originalPart in originalParts)
            {
                var targetPart = targetParts.FirstOrDefault(x => x.Uri == originalPart.Uri);
                if (targetPart is null) continue;
                // Get the relationship IDs for hyperlinks in the source
                var sourceRelIds = originalPart!.HyperlinkRelationships.Select(r => r.Id).ToList();

                // Copy each hyperlink relationship from source to target
                foreach (var relId in sourceRelIds)
                {
                    // Get the relationship from source
                    HyperlinkRelationship? sourceRel = originalPart!.HyperlinkRelationships.FirstOrDefault(x => x.Id == relId);
                    if (sourceRel is null) continue;
                    var sourceTargetUri = sourceRel!.Uri;

                    // Add the same relationship to the target document
                    targetPart!.AddHyperlinkRelationship(sourceTargetUri, sourceRel.IsExternal, sourceRel.Id);
                }

            }
        }
        #endregion
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

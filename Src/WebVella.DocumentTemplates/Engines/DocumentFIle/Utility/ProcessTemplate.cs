using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Globalization;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
using Word = DocumentFormat.OpenXml.Wordprocessing;


namespace WebVella.DocumentTemplates.Engines.DocumentFile.Utility;

public partial class WvDocumentFileEngineUtility
{
    public void ProcessDocumentTemplate(WvDocumentFileTemplateProcessResult result,
        WvDocumentFileTemplateProcessResultItem resultItem, DataTable dataSource,
        CultureInfo culture)
    {
        //Validate
        if (resultItem is null) throw new Exception("No result provided!");
        if (dataSource is null) throw new Exception("No datasource provided!");
        if (result.Template is null) throw new Exception("No Template provided!");
        if (result.WordDocument is null) throw new Exception("No result.WordDocument provided!");
        if (result.WordDocument.MainDocumentPart is null)
            throw new Exception("No  result.WordDocument.MainDocumentPart provided!");
        if (result.WordDocument.MainDocumentPart.Document is null)
            throw new Exception("No  result.WordDocument.MainDocumentPart.Document provided!");
        if (result.WordDocument.MainDocumentPart.Document!.Body is null)
            throw new Exception("No  result.WordDocument.MainDocumentPart.Document.Body provided!");
        if (resultItem.WordDocument is null) throw new Exception("No resultItem.WordDocument provided!");
        if (resultItem.WordDocument.MainDocumentPart is null)
        {
            var mainPart = resultItem.WordDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Word.Body());
        }

        //Process Body
        var templateBody = result.WordDocument.MainDocumentPart.Document!.Body!;
        var resultBody = resultItem.WordDocument.MainDocumentPart!.Document!.Body!;
        var processedTemplateBodyElements = new List<OpenXmlElement>();
        WvTemplateTag? multiParagraphStartTag = null;
        List<OpenXmlElement> multiParagraphTemplateQueue = new();

        #region << Inline Template >>
        //Process inline templates first - inline template in another template is not allowed
        //Option1: inline template single paragraph that is in a paragraph -> should start with the template tag and end with the tag
        //the paragraph will be repeater
        //Option2: inline template multiparagraph that starts with one paragraph, ends with another (the content inside is repeated)

        foreach (var childEl in templateBody.ChildElements)
        {
            var isParagraph = childEl.GetType().FullName == typeof(Word.Paragraph).FullName;
            var paragraphTemplateTags = new List<WvTemplateTag>();
            if (isParagraph)
            {
                var innerText = ((Paragraph)childEl).InnerText.Trim();
                paragraphTemplateTags = new WvTemplateUtility().GetTagsFromTemplate(innerText);
            }

            #region << Outside a MultiParagraph template >>

            //Check for multiparagraph template start

            if (multiParagraphStartTag is null)
            {
                if (isParagraph)
                {
                    //check for multiparagraph inline template start
                    // ReSharper disable once MergeIntoPattern
                    if (paragraphTemplateTags.Count == 1 &&
                        paragraphTemplateTags[0].Type == WvTemplateTagType.InlineStart)
                    {
                        multiParagraphStartTag = paragraphTemplateTags[0];
                        continue;
                    }

                    //check for single paragraph inline template start
                    if (paragraphTemplateTags.Count >= 2
                        // ReSharper disable once MergeIntoPattern
                        && paragraphTemplateTags[0].Type == WvTemplateTagType.InlineStart
                        && paragraphTemplateTags.Last().Type == WvTemplateTagType.InlineEnd)
                    {
                        processedTemplateBodyElements.AddRange(
                            _processSingleParagraphInlineTemplate(paragraphTemplateTags[0], (Paragraph)childEl,
                                dataSource, culture));
                        continue;
                    }
                }

                processedTemplateBodyElements.Add(childEl);
                continue;
            }

            #endregion

            #region << Inside the template >>

            //check for multiparagraph inline template end
            // ReSharper disable once MergeIntoPattern
            if (paragraphTemplateTags.Count == 1 && paragraphTemplateTags[0].Type == WvTemplateTagType.InlineEnd)
            {
                processedTemplateBodyElements.AddRange(_processMultiParagraphInlineTemplate(multiParagraphStartTag!,
                    multiParagraphTemplateQueue, dataSource, culture));
                multiParagraphStartTag = null;
                continue;
            }

            multiParagraphTemplateQueue.Add(childEl);

            #endregion
        }

        //if something is not closed add it for processing
        if (multiParagraphTemplateQueue.Count > 0)
        {
            foreach (var element in multiParagraphTemplateQueue)
                processedTemplateBodyElements.Add(element);
        }

        #endregion

        foreach (var childEl in processedTemplateBodyElements)
        {
            var resultChildElList = _processDocumentElement(
                templateEl: childEl,
                dataSource: dataSource,
                culture: culture);
            foreach (var resultChildEl in resultChildElList)
                resultBody.AppendChild(resultChildEl);
        }

        _copyDocumentStylesAndSettings(result.WordDocument, resultItem.WordDocument);
        _copyHeadersAndFooters(result.WordDocument, resultItem.WordDocument, dataSource, culture);
        _copyHeadersAndFootersHyperlinks(result.WordDocument, resultItem.WordDocument);
        _copyFootnotesAndEndnotes(result.WordDocument, resultItem.WordDocument);
        _copyImages(result.WordDocument, resultItem.WordDocument);
        _copyHyperlinks(result.WordDocument, resultItem.WordDocument);

    }


    private List<OpenXmlElement> _processSingleParagraphInlineTemplate(WvTemplateTag tag, Paragraph template,
        DataTable dataSource,
        CultureInfo culture)
    {
        var result = new List<OpenXmlElement>();
        foreach (DataRow row in dataSource.Rows)
        {
            if (tag.IndexList is not null
                && tag.IndexList.Count > 0
                && !tag.IndexList.Contains(dataSource.Rows.IndexOf(row))) continue;

            DataTable newTable = dataSource.Clone();
            newTable.ImportRow(row);
            result.AddRange(_processDocumentElement(template, newTable, culture));
        }

        return result;
    }

    private List<OpenXmlElement> _processMultiParagraphInlineTemplate(WvTemplateTag tag, List<OpenXmlElement> queue,
        DataTable dataSource,
        CultureInfo culture)
    {
        var result = new List<OpenXmlElement>();
        foreach (DataRow row in dataSource.Rows)
        {
            if (tag.IndexList is not null
                && tag.IndexList.Count > 0
                && !tag.IndexList.Contains(dataSource.Rows.IndexOf(row))) continue;

            DataTable newTable = dataSource.Clone();
            newTable.ImportRow(row);
            foreach (var element in queue)
            {
                result.AddRange(_processDocumentElement(element, newTable, culture));
            }
        }

        queue.Clear();
        return result;
    }
}
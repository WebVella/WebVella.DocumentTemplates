using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Globalization;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Extensions;
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
        IsolateTemplateTags(result.WordDocument); //Ensures tags are not split across multiple texts
        result.WordDocument.Save();        
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


    private List<OpenXmlElement> _processSingleParagraphInlineTemplate(WvTemplateTag firstStartTag, Paragraph template,
        DataTable dataSource,
        CultureInfo culture)
    {
        var result = new List<OpenXmlElement>();
        //Alter DataSource for spec
        DataTable templateDt = firstStartTag.IndexGroups.Count > 0
            ? dataSource.CreateAsNew(firstStartTag.IndexGroups[0].Indexes)
            : dataSource;

        Paragraph cleanTemplate = (Paragraph)_cleanInlineTemplateTags((OpenXmlElement)template);
       
        
        //the general case when we want grouping in the general iteration
        if (String.IsNullOrWhiteSpace(firstStartTag.ItemName))
        {
            foreach (DataRow row in templateDt.Rows)
            {
                DataTable newTable = templateDt.Clone();
                newTable.ImportRow(row);
                result.AddRange(_processDocumentElement(template, newTable, culture));
            }
        }
        //if there is an item name defined so new datasource should be created
        else
        {
            var columnIndices =
                new WvTemplateUtility().GetColumnsIndexFromTagItemName(firstStartTag.ItemName, templateDt);
            //matches no columns:
            if (columnIndices.Count == 0)
            {
                //1. Empty DS returns 
                return result;
            }

            //each row should create its own datasource according to the template rules, so the template can be 
            // rendered with it
            List<DataColumn> columns = new();
            foreach (var index in columnIndices)
            {
                columns.Add(templateDt.Columns[index]);
            }

            //Init the new DT
            var rowDataTable = new DataTable();
            var originalNameNewNameDict = new Dictionary<string, string>();
            foreach (var column in columns)
            {
                var newName = column.ColumnName.Substring(firstStartTag.ItemName.Length);
                if (String.IsNullOrWhiteSpace(newName))
                    newName = firstStartTag.ItemName;
                originalNameNewNameDict[column.ColumnName] = newName;
                var newType = column.DataType;
                var (isEnumarable, dataType) = new WvTemplateUtility().CheckEnumerable(column);
                if (isEnumarable)
                    newType = dataType!; //should have type if enumarable
                rowDataTable.Columns.Add(newName, newType);
            }

            //Process row by row
            foreach (DataRow templateDtRow in templateDt.Rows)
            {
                int rowDtRows = 0; //calculated based on the maximum values found in the columns
                foreach (var templateDtColumn in columns)
                {
                    var (isEnumarable, dataType) = new WvTemplateUtility().CheckEnumerable(templateDtColumn);
                    if (!isEnumarable)
                        rowDtRows = Math.Max(rowDtRows, 1);
                    else
                    {
                        var valuesCount = ((IEnumerable<object>)templateDtRow[templateDtColumn.ColumnName]).Count();
                        rowDtRows = Math.Max(rowDtRows, valuesCount);
                    }
                }

                if (rowDtRows == 0) continue;

                rowDataTable.Clear();
                for (int rowDtRowIndex = 0; rowDtRowIndex < rowDtRows; rowDtRowIndex++)
                {
                    var dsrow = rowDataTable.NewRow();
                    foreach (DataColumn templateDtColumn in columns)
                    {
                        var templateDtColumnValue = templateDtRow[templateDtColumn.ColumnName];
                        var (isEnumarable, dataType) = new WvTemplateUtility().CheckEnumerable(templateDtColumn);
                        if (!isEnumarable)
                        {
                            dsrow[originalNameNewNameDict[templateDtColumn.ColumnName]] = templateDtColumnValue;
                        }
                        else
                        {
                            var indexValue =
                                new WvTemplateUtility().GetItemAt((IEnumerable<object?>)templateDtColumnValue,
                                    rowDtRowIndex);
                            if (indexValue is null)
                                indexValue =
                                    new WvTemplateUtility().GetItemAt((IEnumerable<object?>)templateDtColumnValue, 0);
                            dsrow[originalNameNewNameDict[templateDtColumn.ColumnName]] = indexValue;
                        }
                    }

                    rowDataTable.Rows.Add(dsrow);
                }

                result.AddRange(_processDocumentElement(cleanTemplate, rowDataTable, culture));
                // foreach (DataRow row in rowDataTable.Rows)
                // {
                //     DataTable newTable = rowDataTable.Clone();
                //     newTable.ImportRow(row);
                //     result.AddRange(_processDocumentElement(cleanTemplate, newTable, culture));
                // }
            }
        }


        return result;
    }

    private List<OpenXmlElement> _processMultiParagraphInlineTemplate(WvTemplateTag firstStartTag,
        List<OpenXmlElement> queue,
        DataTable dataSource,
        CultureInfo culture)
    {
        var result = new List<OpenXmlElement>();
        //Alter DataSource for spec
        DataTable templateDt = firstStartTag.IndexGroups.Count > 0
            ? dataSource.CreateAsNew(firstStartTag.IndexGroups[0].Indexes)
            : dataSource;

        //the general case when we want grouping in the general iteration
        if (String.IsNullOrWhiteSpace(firstStartTag.ItemName))
        {
            foreach (DataRow row in templateDt.Rows)
            {
                DataTable newTable = templateDt.Clone();
                newTable.ImportRow(row);
                foreach (var element in queue)
                {
                    result.AddRange(_processDocumentElement(element, newTable, culture));
                }
            }
        }
        //if there is an item name defined so new datasource should be created
        else
        {
            var columnIndices =
                new WvTemplateUtility().GetColumnsIndexFromTagItemName(firstStartTag.ItemName, templateDt);
            //matches no columns:
            if (columnIndices.Count == 0)
            {
                //1. Empty DS returns 
                return result;
            }

            //each row should create its own datasource according to the template rules, so the template can be 
            // rendered with it
            List<DataColumn> columns = new();
            foreach (var index in columnIndices)
            {
                columns.Add(templateDt.Columns[index]);
            }

            //Init the new DT
            var rowDataTable = new DataTable();
            var originalNameNewNameDict = new Dictionary<string, string>();
            foreach (var column in columns)
            {
                var newName = column.ColumnName.Substring(firstStartTag.ItemName.Length);
                if (String.IsNullOrWhiteSpace(newName))
                    newName = firstStartTag.ItemName;
                originalNameNewNameDict[column.ColumnName] = newName;
                var newType = column.DataType;
                var (isEnumarable, dataType) = new WvTemplateUtility().CheckEnumerable(column);
                if (isEnumarable)
                    newType = dataType!; //should have type if enumarable
                rowDataTable.Columns.Add(newName, newType);
            }

            //Process row by row
            foreach (DataRow templateDtRow in templateDt.Rows)
            {
                int rowDtRows = 0; //calculated based on the maximum values found in the columns
                foreach (var templateDtColumn in columns)
                {
                    var (isEnumarable, dataType) = new WvTemplateUtility().CheckEnumerable(templateDtColumn);
                    if (!isEnumarable)
                        rowDtRows = Math.Max(rowDtRows, 1);
                    else
                    {
                        var valuesCount = ((IEnumerable<object>)templateDtRow[templateDtColumn.ColumnName]).Count();
                        rowDtRows = Math.Max(rowDtRows, valuesCount);
                    }
                }

                if (rowDtRows == 0) continue;

                rowDataTable.Clear();
                for (int rowDtRowIndex = 0; rowDtRowIndex < rowDtRows; rowDtRowIndex++)
                {
                    var dsrow = rowDataTable.NewRow();
                    foreach (DataColumn templateDtColumn in columns)
                    {
                        var templateDtColumnValue = templateDtRow[templateDtColumn.ColumnName];
                        var (isEnumarable, dataType) = new WvTemplateUtility().CheckEnumerable(templateDtColumn);
                        if (!isEnumarable)
                        {
                            dsrow[originalNameNewNameDict[templateDtColumn.ColumnName]] = templateDtColumnValue;
                        }
                        else
                        {
                            var indexValue =
                                new WvTemplateUtility().GetItemAt((IEnumerable<object?>)templateDtColumnValue,
                                    rowDtRowIndex);
                            if (indexValue is null)
                                indexValue =
                                    new WvTemplateUtility().GetItemAt((IEnumerable<object?>)templateDtColumnValue, 0);
                            dsrow[originalNameNewNameDict[templateDtColumn.ColumnName]] = indexValue;
                        }
                    }

                    rowDataTable.Rows.Add(dsrow);
                }
                foreach (var element in queue)
                {
                    var cleanElement = _cleanInlineTemplateTags(element);
                    result.AddRange(_processDocumentElement(cleanElement, rowDataTable, culture));
                }
            }
        }

        queue.Clear();
        return result;
    }

    private OpenXmlElement _cleanInlineTemplateTags(OpenXmlElement element)
    {
        var resultEl = element.CloneNode(true);
        if (element.GetType().FullName == typeof(Word.Text).FullName)
        {
            var textEl = (Word.Text)resultEl;
            var innerText = textEl.InnerText;
            if (!String.IsNullOrWhiteSpace(innerText))
            {
                var tagsInTemplate = new WvTemplateUtility().GetTagsFromTemplate(innerText);
                foreach (var tag in tagsInTemplate)
                {
                    if (String.IsNullOrWhiteSpace(tag.FullString)) continue;
                    if (tag.Type != WvTemplateTagType.InlineStart
                        && tag.Type != WvTemplateTagType.InlineEnd)
                        continue;
                    innerText = innerText.Replace(tag.FullString, string.Empty);
                }

                if (textEl.InnerText != innerText)
                {
                    textEl.Space = SpaceProcessingModeValues.Preserve;
                    textEl.Text = innerText;
                }
            }

        }
        else
        {
            resultEl.RemoveAllChildren();
            foreach (var child in element.ChildElements)
            {
                var cleaned = _cleanInlineTemplateTags(child);
                if (cleaned.GetType().FullName == typeof(Word.Text).FullName
                    && String.IsNullOrEmpty(cleaned.InnerText))
                    continue;
                if (cleaned.GetType().FullName == typeof(Word.Run).FullName 
                    && cleaned.ChildElements.Count == 0)
                    continue;
                
                resultEl.AppendChild(cleaned);
            }
        }
        return resultEl;
    }
}
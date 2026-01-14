using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Engines.SpreadsheetFile;
using WebVella.DocumentTemplates.Engines.SpreadsheetFile.Utility;
using WebVella.DocumentTemplates.Engines.Text;
using WebVella.DocumentTemplates.Extensions;
using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace WebVella.DocumentTemplates.Engines.DocumentFile.Utility;

public partial class WvDocumentFileEngineUtility
{
    private List<OpenXmlElement> _processDocumentParagaraph(Word.Paragraph template,
        DataTable dataSource, CultureInfo culture)
    {
        Word.Paragraph resultEl = (Word.Paragraph)template.CloneNode(true);
        if (String.IsNullOrWhiteSpace(template.InnerText)) return [resultEl];
        resultEl.RemoveAllChildren();

        #region << Process template >>
        //Runs Need to be preprocessed as there is often a problem where the tag is 
        //delivered in different runs
        var queue = new Queue<Word.Run>();
        int queueStartTags = 0;
        int queueEndTags = 0;
        var startTagRegex = new Regex(@"{");
        var endTagRegex = new Regex(@"}");
        foreach (var childEl in template.ChildElements)
        {
            var innerText = childEl.InnerText;
            var resultChildElList = _processDocumentElement(childEl, dataSource, culture);
            foreach (var resultChildEl in resultChildElList)
            {
                var isRun = resultChildEl.GetType().FullName == typeof(Word.Run).FullName;
                var startTagsCount = startTagRegex.Matches(resultChildEl.InnerText).Count;
                var endTagsCount = endTagRegex.Matches(resultChildEl.InnerText).Count;
                var hasUnclosedStartTags = startTagsCount > endTagsCount;
                var hasUnclosedEndTags = startTagsCount < endTagsCount;

                #region << Outside queue >>

                if (template.ChildElements.Count == 1)
                {
                    if(resultChildEl.InnerText != String.Empty)
                        resultEl.AppendChild(resultChildEl);
                }
                else if (queue.Count == 0)
                {
                    //Start a queue
                    if (isRun && hasUnclosedStartTags)
                    {
                        queue.Enqueue((Word.Run)childEl);
                        queueStartTags = startTagsCount;
                        queueEndTags = endTagsCount;
                    }
                    //No queue needed
                    else
                    {
                        if (isRun)
                        {
                            var run = (Word.Run)resultChildEl;
                            if(run.ChildElements.Any(x=> x.GetType().FullName != typeof(Word.Text).FullName)
                               || run.InnerText != String.Empty)
                                resultEl.AppendChild(resultChildEl);    
                        }
                        else
                        {
                            resultEl.AppendChild(resultChildEl);
                        }
                    }
                }

                #endregion

                #region << Inside queue >>

                else if (queue.Count > 0)
                {
                    //Case 2 this is not a Run element (need to process the queue and than add the element
                    if (!isRun)
                    {
                        _processQueue(queue, resultEl, dataSource, culture);
                        if(resultChildEl.InnerText != String.Empty)
                            resultEl.AppendChild(resultChildEl);
                    }
                    //Case 1 this is the last element
                    else if (childEl == template.ChildElements.Last())
                    {
                        queue.Enqueue((Word.Run)childEl);
                        _processQueue(queue, resultEl, dataSource, culture);
                    }
                    //Case 3 queue closed and unclose tags are equalised with this run
                    else if ((queueStartTags + startTagsCount) == (queueEndTags + endTagsCount))
                    {
                        queue.Enqueue((Word.Run)childEl);
                        _processQueue(queue, resultEl, dataSource, culture);
                    }
                    //Enqueue
                    else
                    {
                        queue.Enqueue((Word.Run)childEl);
                        queueStartTags += startTagsCount;
                        queueEndTags += endTagsCount;                        
                    }
                }

                #endregion
            }
        }
        #endregion
        return [resultEl];
    }

    private void _processQueue(Queue<Word.Run> queue, Word.Paragraph resultEl,
        DataTable dataSource, CultureInfo culture)
    {
        Word.Run? mergedElement = null;
        StringBuilder sb = new();
        while (queue.Count > 0)
        {
            var element = queue.Dequeue();
            if (mergedElement is null)
                mergedElement = (Word.Run)element.CloneNode(true);
            sb.Append(element.InnerText);
        }

        if (mergedElement is null) return;

        mergedElement.RemoveAllChildren<Word.Text>();
        var textEl = new Word.Text();
        textEl.Text = sb.ToString();
        mergedElement.AppendChild(textEl);

        var resultChildElList =
            _processDocumentElement(mergedElement, dataSource, culture);
        foreach (var resultChildEl in resultChildElList)
            resultEl.AppendChild(resultChildEl);
    }
}
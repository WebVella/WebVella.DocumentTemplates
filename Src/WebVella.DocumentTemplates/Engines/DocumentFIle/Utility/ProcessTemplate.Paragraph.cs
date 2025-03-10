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
using System.Text.RegularExpressions;

namespace WebVella.DocumentTemplates.Engines.DocumentFile.Utility;
public partial class WvDocumentFileEngineUtility
{
	private Word.Paragraph _processDocumentParagaraph(Word.Paragraph template,
		DataTable dataSource, CultureInfo culture)
	{
		Word.Paragraph resultEl = (Word.Paragraph)template.CloneNode(true);
		if (String.IsNullOrWhiteSpace(template.InnerText)) return resultEl;
		resultEl.RemoveAllChildren();

		//Runs Need to be preprocessed as there is often a problem where the tag is 
		//delivered in different runs
		var queue = new Queue<Word.Run>();
		int queueStartTags = 0;
		int queueEndTags = 0;
		foreach (var childEl in template.ChildElements)
		{
			var resultChildEl = _processDocumentElement(childEl, dataSource, culture);
			if (resultChildEl is null) continue;

			var isRun = resultChildEl.GetType().FullName == typeof(Word.Run).FullName;
			var startTagRegex = new Regex(@"{{");
			var endTagRegex = new Regex(@"}}");
			var startTagsCount = startTagRegex.Matches(resultChildEl.InnerText).Count;
			var endTagsCount = endTagRegex.Matches(resultChildEl.InnerText).Count;
			var hasUnclosedStartTags = startTagsCount > endTagsCount;
			var hasUnclosedEndTags = startTagsCount < endTagsCount;
			#region << Outside queue >>
			if(template.ChildElements.Count == 1){ 
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
					resultEl.AppendChild(resultChildEl);
				}
			}
			#endregion

			#region << Inside queue >>
			else if (queue.Count > 0)
			{
				//Case 2 this is not a Run element (need to process the queue and than add the element
				if (!isRun)
				{
					_processQueue(queue, resultEl,dataSource,culture);
					resultEl.AppendChild(resultChildEl);
				}
				//Case 1 this is the last element
				else if (childEl == template.ChildElements.Last())
				{
					queue.Enqueue((Word.Run)childEl);
					_processQueue(queue, resultEl,dataSource,culture);
				}
				//Case 3 queue closed and unclose tags are equalised with this run
				else if ((queueStartTags + startTagsCount) == (queueEndTags + endTagsCount))
				{
					queue.Enqueue((Word.Run)childEl);
					_processQueue(queue, resultEl,dataSource,culture);
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
		return resultEl;
	}

	private void _processQueue(Queue<Word.Run> queue, Word.Paragraph resultEl,
	DataTable dataSource, CultureInfo culture)
	{
		Word.Run? mergedElement = null;
		StringBuilder sb = new();
		while(queue.Count > 0){ 
			var element = queue.Dequeue();
			if(mergedElement is null)
				mergedElement = (Word.Run)element.CloneNode(true);
			sb.Append(element.InnerText);
		}
		if(mergedElement is null) return;

		mergedElement.RemoveAllChildren();
		var textEl = new Word.Text();
		textEl.Text = sb.ToString();
		mergedElement.AppendChild(textEl);

		var resultChildEl = _processDocumentElement(mergedElement, dataSource, culture);
		resultEl.AppendChild(resultChildEl);
	}
}

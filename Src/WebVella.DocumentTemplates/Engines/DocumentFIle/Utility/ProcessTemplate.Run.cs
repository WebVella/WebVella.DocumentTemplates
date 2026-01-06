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

namespace WebVella.DocumentTemplates.Engines.DocumentFile.Utility;

public partial class WvDocumentFileEngineUtility
{
    private List<OpenXmlElement> _processDocumentRun(Word.Run template,
        DataTable dataSource, CultureInfo culture,
        Dictionary<string, WvDocumentFileTemplate> templateLibrary,
        int stackLevel)
    {
        Word.Run resultEl = (Word.Run)template.CloneNode(true);
        if (String.IsNullOrWhiteSpace(template.InnerText)) return [resultEl];
        resultEl.RemoveAllChildren();
        //Process Runs
        foreach (var childEl in template.ChildElements)
        {
            var resultChildElList = _processDocumentElement(childEl, dataSource, culture, templateLibrary, stackLevel);
            foreach (var resultChildEl in resultChildElList)
                resultEl.AppendChild(resultChildEl);
        }

        return [resultEl];
    }
}
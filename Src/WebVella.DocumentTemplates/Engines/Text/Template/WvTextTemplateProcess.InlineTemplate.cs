using DocumentFormat.OpenXml.Vml;
using System.Data;
using System.Globalization;
using System.Text;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;

namespace WebVella.DocumentTemplates.Engines.Text;

public partial class WvTextTemplateProcess
{
    public string? ProcessInlineTemplates(string? processedTemplate, DataTable dataSource, CultureInfo culture)
    {
        if (string.IsNullOrWhiteSpace(processedTemplate)) return processedTemplate;

        //RULE: inline template in another template is not allowed
        //Option1: inline template single line that is in a line -> should start with the template tag and end with the tag
        //the paragraph will be repeater
        //Option2: inline template multiline that starts with one line, ends with another (the content inside is repeated)

        WvTemplateTag? multiLineStartTag = null;
        List<string> multiLineTemplateQueue = new();
        var sb = new StringBuilder();
        var lines = processedTemplate.Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries);
        var endWithNewLine = processedTemplate.EndsWith(Environment.NewLine);
        if (lines.Count() == 1)
        {
            var lineTemplateTags = new WvTemplateUtility().GetTagsFromTemplate(lines[0]);
            if (lineTemplateTags.Count >= 2
                // ReSharper disable once MergeIntoPattern
                && lineTemplateTags[0].Type == WvTemplateTagType.InlineStart
                && lineTemplateTags.Last().Type == WvTemplateTagType.InlineEnd)
            {
                var resultLines = _processSingleLineInlineTemplate(lineTemplateTags[0], lines[0],
                    dataSource, culture);
                foreach (var resultLine in resultLines)
                {
                    if (endWithNewLine)
                        sb.AppendLine(resultLine);
                    else
                        sb.Append(resultLine);
                }
            }
            else
            {
                if (endWithNewLine)
                    sb.AppendLine(lines[0]);
                else
                    sb.Append(lines[0]);
            }
        }
        else
        {
            foreach (var line in lines)
            {
                var lineTemplateTags = new WvTemplateUtility().GetTagsFromTemplate(line);

                #region << Outside a MultiLine template >>

                if (multiLineStartTag is null)
                {
                    //check for multiparagraph inline template start
                    // ReSharper disable once MergeIntoPattern
                    if (lineTemplateTags.Count == 1 &&
                        lineTemplateTags[0].Type == WvTemplateTagType.InlineStart)
                    {
                        multiLineStartTag = lineTemplateTags[0];
                        multiLineTemplateQueue.Add(line);
                        continue;
                    }

                    //check for single paragraph inline template start
                    if (lineTemplateTags.Count >= 2
                        // ReSharper disable once MergeIntoPattern
                        && lineTemplateTags[0].Type == WvTemplateTagType.InlineStart
                        && lineTemplateTags.Last().Type == WvTemplateTagType.InlineEnd)
                    {
                        var resultLines = _processSingleLineInlineTemplate(lineTemplateTags[0], line,
                            dataSource, culture);
                        foreach (var resultLine in resultLines)
                        {
                            sb.AppendLine(resultLine);
                        }

                        continue;
                    }

                    sb.AppendLine(line);
                    continue;
                }

                #endregion

                #region << Inside the template >>

                //check for multiparagraph inline template end
                // ReSharper disable once MergeIntoPattern
                if (lineTemplateTags.Count == 1 &&
                    lineTemplateTags[0].Type == WvTemplateTagType.InlineEnd)
                {
                    multiLineTemplateQueue.Add(line);
                    var resultLines = _processMultiLineInlineTemplate(multiLineStartTag,
                        multiLineTemplateQueue, dataSource, culture);
                    multiLineStartTag = null;
                    foreach (var resultLine in resultLines)
                    {
                        sb.AppendLine(resultLine);
                    }

                    continue;
                }

                multiLineTemplateQueue.Add(line);

                #endregion
            }
        }

        return sb.ToString();
    }

    private List<string> _processSingleLineInlineTemplate(WvTemplateTag tag, string template,
        DataTable dataSource,
        CultureInfo culture)
    {
        var processedItems = new List<string>();
        //foreach (DataRow row in dataSource.Rows)
        //{
        //    if (tag.IndexGroups.Count > 0
        //        && !tag.IndexGroups[0].Indexes.Contains(dataSource.Rows.IndexOf(row))) continue;

        //    DataTable newTable = dataSource.Clone();
        //    newTable.ImportRow(row);
        //    var tagProcessResult =
        //        new WvTemplateUtility().ProcessTemplateTag(template, newTable, culture);
        //    foreach (var value in tagProcessResult.Values)
        //    {
        //        var stringValue = value.ToString();
        //        if (string.IsNullOrWhiteSpace(stringValue)) continue;
        //        processedItems.Add(stringValue);
        //    }
        //}

        var tagProcessResult =
            new WvTemplateUtility().ProcessTemplateTag(template, dataSource, culture);
        foreach (var value in tagProcessResult.Values)
        {
            var stringValue = value.ToString();
            if (string.IsNullOrWhiteSpace(stringValue)) continue;
            processedItems.Add(stringValue);
        }

        return [string.Join(tag.FlowSeparator, processedItems)];
    }

    private List<string> _processMultiLineInlineTemplate(WvTemplateTag tag, List<string> queue,
        DataTable dataSource,
        CultureInfo culture)
    {
        var result = new List<string>();
        //foreach (DataRow row in dataSource.Rows)
        //{
        //    if (tag.IndexGroups.Count > 0
        //        && !tag.IndexGroups[0].Indexes.Contains(dataSource.Rows.IndexOf(row))) continue;

        //    DataTable newTable = dataSource.Clone();
        //    newTable.ImportRow(row);
        //    foreach (var line in queue)
        //    {
        //        var tagProcessResult =
        //            new WvTemplateUtility().ProcessTemplateTag(line, newTable, culture);
        //        foreach (var value in tagProcessResult.Values)
        //        {
        //            var stringValue = value.ToString();
        //            if (string.IsNullOrWhiteSpace(stringValue)) continue;
        //            result.Add(stringValue);
        //        }
        //    }
        //}
        var sb = new StringBuilder();
        foreach (var line in queue)
        {
            sb.AppendLine(line);
        }
        var tagProcessResult =
       new WvTemplateUtility().ProcessTemplateTag(sb.ToString(), dataSource, culture);
        foreach (var value in tagProcessResult.Values)
        {
            var stringValue = value.ToString();
            if (string.IsNullOrWhiteSpace(stringValue)) continue;
            result.Add(stringValue);
        }

        queue.Clear();
        return result;
    }
}
using System.Data;
using System.Globalization;
using System.Text;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Extensions;

namespace WebVella.DocumentTemplates.Engines.Text;

public class WvTextTemplate : WvTemplateBase
{
    public string? Template { get; set; }

    public WvTextTemplateProcessResult Process(DataTable? dataSource, CultureInfo? culture = null)
    {
        if (culture == null) culture = new CultureInfo("en-US");
        if (dataSource is null) throw new ArgumentException("No datasource provided!", nameof(dataSource));

        var result = new WvTextTemplateProcessResult()
        {
            Template = Template,
            GroupByDataColumns = GroupDataByColumns,
            ResultItems = new()
        };
        if (String.IsNullOrWhiteSpace(Template))
        {
            result.ResultItems.Add(new WvTextTemplateProcessResultItem
            {
                Result = Template
            });
            return result;
        }

        ;

        var datasourceGroups = dataSource.GroupBy(GroupDataByColumns);

        foreach (var grouptedDs in datasourceGroups)
        {
            var resultItem = new WvTextTemplateProcessResultItem()
            {
                DataTable = grouptedDs
            };
            var resultContext = new WvTextTemplateProcessContext();
            try
            {
                var processedTemplate = Template;

                #region << Process Inline template >>

                {
                    //Process inline templates first - inline template in another template is not allowed
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
                        var paragraphTemplateTags = new WvTemplateUtility().GetTagsFromTemplate(lines[0]);
                        if (paragraphTemplateTags.Count >= 2
                            // ReSharper disable once MergeIntoPattern
                            && paragraphTemplateTags[0].Type == WvTemplateTagType.InlineStart
                            && paragraphTemplateTags.Last().Type == WvTemplateTagType.InlineEnd)
                        {
                            var resultLines = _processSingleLineInlineTemplate(paragraphTemplateTags[0], lines[0],
                                dataSource, culture);
                            foreach (var resultLine in resultLines)
                            {
                                if(endWithNewLine)
                                    sb.AppendLine(resultLine);
                                else
                                    sb.Append(resultLine);
                            }
                        }
                        else
                        {
                            if(endWithNewLine)
                                sb.AppendLine(lines[0]);
                            else
                                sb.Append(lines[0]);                            
                        }


                    }
                    else
                    {
                        foreach (var line in lines)
                        {
                            var paragraphTemplateTags = new WvTemplateUtility().GetTagsFromTemplate(line);

                            #region << Outside a MultiParagraph template >>

                            if (multiLineStartTag is null)
                            {
                                //check for multiparagraph inline template start
                                if (paragraphTemplateTags.Count == 1 &&
                                    paragraphTemplateTags[0].Type == WvTemplateTagType.InlineStart)
                                {
                                    multiLineStartTag = paragraphTemplateTags[0];
                                    continue;
                                }

                                //check for single paragraph inline template start
                                if (paragraphTemplateTags.Count >= 2
                                    // ReSharper disable once MergeIntoPattern
                                    && paragraphTemplateTags[0].Type == WvTemplateTagType.InlineStart
                                    && paragraphTemplateTags.Last().Type == WvTemplateTagType.InlineEnd)
                                {
                                    var resultLines = _processSingleLineInlineTemplate(paragraphTemplateTags[0], line,
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
                            if (paragraphTemplateTags.Count == 1 &&
                                paragraphTemplateTags[0].Type == WvTemplateTagType.InlineEnd)
                            {
                                var resultLines = _processMultiLineInlineTemplate(multiLineStartTag!,
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

                    processedTemplate = sb.ToString();
                }

                #endregion

                #region << Processtemplate >>

                {
                    var sb = new StringBuilder();
                    var lines = processedTemplate.Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries);
                    var endWithNewLine = processedTemplate.EndsWith(Environment.NewLine);
                    if (lines.Count() == 1)
                    {
                        var tagProcessResult =
                            new WvTemplateUtility().ProcessTemplateTag(lines.First(), grouptedDs, culture);
                        if (tagProcessResult.Values.Count == 1)
                        {
                            sb.Append(tagProcessResult.Values[0]?.ToString());
                        }
                        else if (tagProcessResult.Values.Count > 1)
                        {
                            foreach (var value in tagProcessResult.Values)
                            {
                                if (endWithNewLine)
                                    sb.AppendLine(value?.ToString());
                                else
                                    sb.Append(value?.ToString());
                            }
                        }
                    }
                    else if (lines.Count() > 1)
                    {
                        foreach (string line in lines)
                        {
                            var tagProcessResult =
                                new WvTemplateUtility().ProcessTemplateTag(line, grouptedDs, culture);
                            foreach (var value in tagProcessResult.Values)
                            {
                                sb.AppendLine(value.ToString());
                            }
                        }
                    }

                    processedTemplate = sb.ToString();
                }
                resultItem.Result = processedTemplate;

                #endregion
            }
            catch (Exception ex)
            {
                resultContext.Errors.Add(ex.Message);
            }

            resultItem.Contexts.Add(resultContext);
            result.ResultItems.Add(resultItem);
        }

        return result;
    }

    private List<string> _processSingleLineInlineTemplate(WvTemplateTag tag, string template,
        DataTable dataSource,
        CultureInfo culture)
    {
        var processedItems = new List<string>();
        foreach (DataRow row in dataSource.Rows)
        {
            if (tag.IndexList is not null
                && tag.IndexList.Count > 0
                && !tag.IndexList.Contains(dataSource.Rows.IndexOf(row))) continue;

            DataTable newTable = dataSource.Clone();
            newTable.ImportRow(row);
            var tagProcessResult =
                new WvTemplateUtility().ProcessTemplateTag(template, newTable, culture);
            foreach (var value in tagProcessResult.Values)
            {
                var stringValue = value?.ToString();
                if(string.IsNullOrWhiteSpace(stringValue)) continue;
                processedItems.Add(stringValue);
            }
        }
        return [string.Join(tag.FlowSeparator, processedItems)];
    }

    private List<string> _processMultiLineInlineTemplate(WvTemplateTag tag, List<string> queue,
        DataTable dataSource,
        CultureInfo culture)
    {
        var result = new List<string>();
        foreach (DataRow row in dataSource.Rows)
        {
            if (tag.IndexList is not null
                && tag.IndexList.Count > 0
                && !tag.IndexList.Contains(dataSource.Rows.IndexOf(row))) continue;

            DataTable newTable = dataSource.Clone();
            newTable.ImportRow(row);
            foreach (var line in queue)
            {
                var tagProcessResult =
                    new WvTemplateUtility().ProcessTemplateTag(line, newTable, culture);
                foreach (var value in tagProcessResult.Values)
                {
                    var stringValue = value?.ToString();
                    if(string.IsNullOrWhiteSpace(stringValue)) continue;
                    result.Add(stringValue);
                }                
            }
        }

        queue.Clear();
        return result;
    }
}
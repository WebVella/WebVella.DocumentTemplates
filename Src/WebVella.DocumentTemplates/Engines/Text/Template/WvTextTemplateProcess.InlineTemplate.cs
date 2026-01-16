using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Vml;
using System.Data;
using System.Globalization;
using System.Text;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Extensions;

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

    private List<string> _processSingleLineInlineTemplate(WvTemplateTag firstStartTag, string template,
        DataTable dataSource,
        CultureInfo culture)
    {
        var processedItems = new List<string>();
        DataTable templateDt = firstStartTag.IndexGroups.Count > 0
                    ? dataSource.CreateAsNew(firstStartTag.IndexGroups[0].Indexes)
                    : dataSource;
        var cleanTemplate = _cleanInlineTemplateTags(template);
        //the general case when we want grouping in the general iteration
        if (String.IsNullOrWhiteSpace(firstStartTag.ItemName))
        {
            foreach (DataRow row in templateDt.Rows)
            {
                DataTable newTable = templateDt.Clone();
                newTable.ImportRow(row);
                var tagProcessResult =
                            new WvTemplateUtility().ProcessTemplateTag(cleanTemplate, newTable, culture);
                foreach (var value in tagProcessResult.Values)
                {
                    var stringValue = value.ToString();
                    if (string.IsNullOrWhiteSpace(stringValue)) continue;
                    processedItems.Add(stringValue);
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
                return processedItems;
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
                var rowDataTableIndex = 0;
                for (int rowDtRowIndex = 0; rowDtRowIndex < rowDtRows; rowDtRowIndex++)
                {
                    if (firstStartTag.IndexGroups.Count > 1
                        && !firstStartTag.IndexGroups[1].Indexes.Contains(rowDataTableIndex++))
                        continue;
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

                var tagProcessResult =
                            new WvTemplateUtility().ProcessTemplateTag(cleanTemplate, rowDataTable, culture);
                foreach (var value in tagProcessResult.Values)
                {
                    var stringValue = value.ToString();
                    if (string.IsNullOrWhiteSpace(stringValue)) continue;
                    processedItems.Add(stringValue);
                }
            }
        }

        return [string.Join(firstStartTag.FlowSeparator, processedItems)];
    }

    private List<string> _processMultiLineInlineTemplate(WvTemplateTag firstStartTag, List<string> queue,
        DataTable dataSource,
        CultureInfo culture)
    {
        var result = new List<string>();
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
                    var cleanElement = _cleanInlineTemplateTags(element);

                    //skip empty paragraphs that may had inline tags
                    if (String.IsNullOrEmpty(cleanElement))
                    {
                        var tags = new WvTemplateUtility().GetTagsFromTemplate(element);
                        if (tags.Any(x => x.Type == WvTemplateTagType.InlineStart)
                            || tags.Any(x => x.Type == WvTemplateTagType.InlineEnd))
                            continue;
                    }
                    var tagProcessResult =
                        new WvTemplateUtility().ProcessTemplateTag(cleanElement, newTable, culture);
                    foreach (var value in tagProcessResult.Values)
                    {
                        var stringValue = value.ToString();
                        if (string.IsNullOrWhiteSpace(stringValue)) continue;
                        result.Add(stringValue);
                    }
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
                var rowDataTableIndex = 0;
                foreach (DataRow row in rowDataTable.Rows)
                {
                    if (firstStartTag.IndexGroups.Count > 1
                        && !firstStartTag.IndexGroups[1].Indexes.Contains(rowDataTableIndex++))
                        continue;

                    DataTable newTable = rowDataTable.Clone();
                    newTable.ImportRow(row);
                    foreach (var element in queue)
                    {
                        var cleanElement = _cleanInlineTemplateTags(element);
                        //skip empty paragraphs that may had inline tags
                        if (String.IsNullOrEmpty(cleanElement))
                        {
                            var tags = new WvTemplateUtility().GetTagsFromTemplate(element);
                            if (tags.Any(x => x.Type == WvTemplateTagType.InlineStart)
                                || tags.Any(x => x.Type == WvTemplateTagType.InlineEnd))
                                continue;
                        }
                        var tagProcessResult =
                       new WvTemplateUtility().ProcessTemplateTag(cleanElement, newTable, culture);
                        foreach (var value in tagProcessResult.Values)
                        {
                            var stringValue = value.ToString();
                            if (string.IsNullOrWhiteSpace(stringValue)) continue;
                            result.Add(stringValue);
                        }
                    }
                }
            }
        }
        // var sb = new StringBuilder();
        // foreach (var line in queue)
        // {
        //     sb.AppendLine(line);
        // }
        // var tagProcessResult =
        //new WvTemplateUtility().ProcessTemplateTag(sb.ToString(), dataSource, culture);
        // foreach (var value in tagProcessResult.Values)
        // {
        //     var stringValue = value.ToString();
        //     if (string.IsNullOrWhiteSpace(stringValue)) continue;
        //     result.Add(stringValue);
        // }

        queue.Clear();
        return result;
    }

    private string _cleanInlineTemplateTags(string template)
    {
        var tagsInTemplate = new WvTemplateUtility().GetTagsFromTemplate(template);
        foreach (var tag in tagsInTemplate)
        {
            if (String.IsNullOrWhiteSpace(tag.FullString)) continue;
            if (tag.Type != WvTemplateTagType.InlineStart
                && tag.Type != WvTemplateTagType.InlineEnd)
                continue;
            template = template.Replace(tag.FullString, string.Empty);
        }
        return template;
    }
}
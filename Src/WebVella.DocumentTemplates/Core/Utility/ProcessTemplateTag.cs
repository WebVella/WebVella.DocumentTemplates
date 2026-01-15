using System.Data;
using System.Globalization;
using WebVella.DocumentTemplates.Extensions;

namespace WebVella.DocumentTemplates.Core.Utility;

public partial class WvTemplateUtility
{
    public WvTemplateTagResultList ProcessTemplateTag(
        string? template,
        DataTable dataSource,
        CultureInfo culture)
    {
        var result = new WvTemplateTagResultList
        {
            Tags = GetTagsFromTemplate(template)
        };
        //if there are no tags - return one with the template
        if (result.Tags.Count == 0 || dataSource.Rows.Count == 0)
        {
            result.Values.Add(template ?? String.Empty);
            return result;
        }

        #region << Conditional Tags >>

        if (result.Tags.Any(x => x.Type == WvTemplateTagType.ConditionStart) &&
            result.Tags.Any(x => x.Type == WvTemplateTagType.ConditionEnd))
        {
            var firstStartTag = result.Tags.First(x => x.Type == WvTemplateTagType.ConditionStart);
            var lastEndTag = result.Tags.Last(x => x.Type == WvTemplateTagType.ConditionEnd);
            var firstTagIndex = template!.IndexOf(firstStartTag.FullString!, StringComparison.Ordinal);
            var lastTagIndex = template!.LastIndexOf(lastEndTag.FullString!, StringComparison.Ordinal);
            string templateWithWrappers = template!;
            templateWithWrappers = templateWithWrappers.Substring(firstTagIndex);
            templateWithWrappers = templateWithWrappers.Substring(0, (lastTagIndex + lastEndTag.FullString!.Length));

            string templateWithoutWrappers = templateWithWrappers;
            templateWithoutWrappers = templateWithoutWrappers.Substring(firstStartTag.FullString!.Length);
            templateWithoutWrappers =
                templateWithoutWrappers.Substring(0, templateWithoutWrappers.Length - lastEndTag.FullString!.Length);

            string cleanedTemplate = templateWithoutWrappers;
            //clean all other inline tags
            foreach (var tag in result.Tags)
            {
                if (tag.Type != WvTemplateTagType.InlineStart && tag.Type != WvTemplateTagType.InlineEnd)
                    continue;
                cleanedTemplate = cleanedTemplate.Replace(tag.FullString!, "");
            }

            if (firstStartTag.ParamGroups.Count > 0)
            {
                template = template.Replace(templateWithWrappers, String.Empty);
            }
            else
            {
                template = template.Replace(templateWithWrappers, cleanedTemplate);
            }

            result.Tags = GetTagsFromTemplate(template);
        }

        #endregion

        #region <<Inline Template >>

        //Process inline templates 
        if (result.Tags.Any(x => x.Type == WvTemplateTagType.InlineStart) &&
            result.Tags.Any(x => x.Type == WvTemplateTagType.InlineEnd))
        {
            var firstStartTag = result.Tags.First(x => x.Type == WvTemplateTagType.InlineStart);
            var lastEndTag = result.Tags.Last(x => x.Type == WvTemplateTagType.InlineEnd);
            var firstTagIndex = template!.IndexOf(firstStartTag.FullString!, StringComparison.Ordinal);
            var lastTagIndex = template!.LastIndexOf(lastEndTag.FullString!, StringComparison.Ordinal);

            string templateWithWrappers = template!;
            templateWithWrappers = templateWithWrappers.Substring(firstTagIndex);
            templateWithWrappers = templateWithWrappers.Substring(0, (lastTagIndex + lastEndTag.FullString!.Length));

            string templateWithoutWrappers = templateWithWrappers;
            templateWithoutWrappers = templateWithoutWrappers.Substring(firstStartTag.FullString!.Length);
            templateWithoutWrappers =
                templateWithoutWrappers.Substring(0, templateWithoutWrappers.Length - lastEndTag.FullString!.Length);

            string cleanedTemplate = templateWithoutWrappers;
            //clean all other inline tags
            foreach (var tag in result.Tags)
            {
                if (tag.Type != WvTemplateTagType.InlineStart && tag.Type != WvTemplateTagType.InlineEnd)
                    continue;
                cleanedTemplate = cleanedTemplate.Replace(tag.FullString!, "");
            }

            DataTable templateDt = firstStartTag.IndexGroups.Count > 0
                ? dataSource.CreateAsNew(firstStartTag.IndexGroups[0].Indexes)
                : dataSource;
            //the general case when we want grouping in the general iteration
            if (String.IsNullOrWhiteSpace(firstStartTag.ItemName))
            {
                var inlineTemplateResult = ProcessTemplateTag(cleanedTemplate, templateDt, culture);
                if (inlineTemplateResult.Values.Count == 1)
                {
                    if (inlineTemplateResult.Values[0] is string)
                        template = template.Replace(templateWithWrappers, (string)inlineTemplateResult.Values[0]);
                }
                else if (inlineTemplateResult.Values.All(x => x is string))
                {
                    var separator = firstStartTag.FlowSeparator;
                    template = template.Replace(templateWithWrappers,
                        String.Join(separator, inlineTemplateResult.Values.Select(x => (string)x)));
                }
            }
            //if there is an item name defined so new datasource should be created
            else
            {
                var columnIndices = GetColumnsIndexFromTagItemName(firstStartTag.ItemName, templateDt);
                //matches no columns:
                if (columnIndices.Count == 0)
                {
                    //1. Empty DS returns 
                    result.Values.Add(template);
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
                    var (isEnumarable, dataType) = CheckEnumerable(column);
                    if (isEnumarable)
                        newType = dataType!; //should have type if enumarable
                    rowDataTable.Columns.Add(newName, newType);
                }

                //Process row by row
                var templateDtRowResults = new List<string>();
                foreach (DataRow templateDtRow in templateDt.Rows)
                {
                    int rowDtRows = 0; //calculated based on the maximum values found in the columns
                    foreach (var templateDtColumn in columns)
                    {
                        var (isEnumarable, dataType) = CheckEnumerable(templateDtColumn);
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
                            var (isEnumarable, dataType) = CheckEnumerable(templateDtColumn);
                            if (!isEnumarable)
                            {
                                dsrow[originalNameNewNameDict[templateDtColumn.ColumnName]] = templateDtColumnValue;
                            }
                            else
                            {
                                var indexValue = GetItemAt((IEnumerable<object?>)templateDtColumnValue, rowDtRowIndex);
                                if (indexValue is null)
                                    indexValue = GetItemAt((IEnumerable<object?>)templateDtColumnValue, 0);
                                dsrow[originalNameNewNameDict[templateDtColumn.ColumnName]] = indexValue;
                            }
                        }

                        rowDataTable.Rows.Add(dsrow);
                    }

                    var inlineTemplateResult = ProcessTemplateTag(cleanedTemplate, rowDataTable, culture);
                    if (inlineTemplateResult.Values.Count == 1)
                    {
                        if (inlineTemplateResult.Values[0] is string)
                            templateDtRowResults.Add(template.Replace(templateWithWrappers, (string)inlineTemplateResult.Values[0]));
                    }
                    else
                    {
                        var subSeparator = firstStartTag.FlowSeparator;
                        if (firstStartTag.FlowSeparatorList.Count > 1)
                        {
                            subSeparator = firstStartTag.FlowSeparatorList[1];
                        }

                        templateDtRowResults.Add(template.Replace(templateWithWrappers,
                            String.Join(subSeparator, inlineTemplateResult.Values.Select(x => x?.ToString()))));
                    }
                }
                var separator = firstStartTag.FlowSeparator;
                template = template.Replace(templateWithWrappers,
                            String.Join(separator, templateDtRowResults));
            }
        }

        #endregion

        //if all tags are index - return one with processed template
        if (result.ShouldGenerateOneResult(dataSource))
        {
            var resultValue = GenerateTemplateTagResult(template, result.Tags, dataSource, null, culture);
            if (resultValue is not null && resultValue.Value is not null)
            {
                result.Values.Add(resultValue.Value);
            }

            return result;
        }

        var resultValues = new List<object>();
        for (int i = 0; i < dataSource.Rows.Count; i++)
        {
            var resultValue = GenerateTemplateTagResult(template, result.Tags, dataSource, i, culture);
            if (resultValue is not null && resultValue.Value is not null)
            {
                resultValues.Add(resultValue.Value);
            }
            else
            {
                resultValues.Add(String.Empty);
            }
        }

        if (resultValues.Count == 0)
        {
            result.Values = new();
            result.ExpandCount = 0;
            return result;
        }

        //If all the results are the same as the template return only one
        bool allValuesMatchTemplate = true;
        foreach (var rstValue in resultValues)
        {
            if (rstValue is not string || rstValue.ToString() != template)
            {
                allValuesMatchTemplate = false;
                break;
            }
        }

        if (allValuesMatchTemplate) result.Values.Add(resultValues[0]);
        else result.Values.AddRange(resultValues);
        result.ExpandCount = result.Values.Count;
        return result;
    }

    public object? GetTemplateValue(
        string? template,
        int dataRowPosition,
        DataTable dataSource,
        CultureInfo culture)
    {
        if (String.IsNullOrWhiteSpace(template)) return null;
        var tags = GetTagsFromTemplate(template);
        //if there are no tags - return one with the template
        if (tags.Count == 0) return template;

        string? valueString = template;
        object? value = null;
        if (dataRowPosition < 1) dataRowPosition = 1;
        if (tags.Count == 1 && tags[0].FullString == template)
        {
            (valueString, value) =
                ProcessTagInTemplate(valueString, value, tags[0], dataSource, dataRowPosition - 1, culture);
            return value;
        }

        foreach (var tag in tags)
        {
            (valueString, value) =
                ProcessTagInTemplate(valueString, value, tag, dataSource, dataRowPosition - 1, culture);
        }

        if (value is string) return valueString;

        return value;
    }

    public (bool IsEnumerable, Type? EnumerateElementType) CheckEnumerable(DataColumn column)
    {
        if (column == null)
            throw new ArgumentNullException(nameof(column));

        var columnType = column.DataType;
        if (columnType == typeof(string))
            return (false, null);

        // Look for the generic IEnumerable<T> interface in the type's implemented interfaces
        foreach (var iface in columnType.GetInterfaces())
        {
            if (iface.IsGenericType && iface.GetGenericTypeDefinition() == typeof(IEnumerable<>))
            {
                var genericArguments = iface.GetGenericArguments().ToList();
                if (genericArguments.Count > 0)
                    return (true, genericArguments[0]);
                else
                    return (true, iface.GetGenericTypeDefinition());
            }
        }

        // If no enumerable interface is found, return (false, null)
        return (false, null);
    }

    public object? GetItemAt(IEnumerable<object?>? enumerable, int index)
    {
        if (enumerable is null)
            return null;

        // Optional: validate index range early for efficiency
        using var enumerator = enumerable.GetEnumerator();
        int i = 0;
        while (enumerator.MoveNext())
        {
            if (i == index)
                return enumerator.Current;
            i++;
        }

        return null;
    }
}
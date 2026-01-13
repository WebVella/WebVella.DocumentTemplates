using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;

namespace WebVella.DocumentTemplates.Core.Utility;

public partial class WvTemplateUtility
{
    public (string?, object?) ReplaceSingleValueDataTagInTemplate(string? templateResultString,
        object? templateResultObject, WvTemplateTag tag, DataTable dataSource, int? contextRowIndex)
    {
        var columnIndices = GetColumnsIndexFromTagItemName(tag.ItemName, dataSource);

        if (columnIndices.Count == 0 || dataSource.Rows.Count == 0)
        {
            if (templateResultObject is null)
                templateResultObject = templateResultString;
            return (templateResultString, templateResultObject);
        }

        int rowIndex = 0;
        if (tag.IndexList is not null && tag.IndexList.Count > 0)
        {
            rowIndex = tag.IndexList[0];
        }
        else if (contextRowIndex is not null && dataSource.Rows.Count - 1 >= contextRowIndex)
        {
            rowIndex = contextRowIndex.Value;
        }


        var valueList = new List<string>();
        object? columnObject = null;
        var oneTagOnlyTemplate = templateResultString?.ToLowerInvariant() == tag.FullString?.ToLowerInvariant();
        foreach (var columnIndex in columnIndices)
        {
            columnObject = dataSource.Rows[rowIndex][columnIndex];
            var columnValue = columnObject?.ToString();
            if (columnObject is IEnumerable<object>)
            {
                //Second separator should be used, if not present the first
                var flowSeparator = tag.FlowSeparatorList.Count >= 2 ? tag.FlowSeparatorList[1] : tag.FlowSeparator;
                columnValue = String.Join(flowSeparator,
                    ((IEnumerable<object>)columnObject).Select(x => x?.ToString()));
            }

            if (!String.IsNullOrWhiteSpace(columnValue))
                valueList.Add(columnValue);
        }

        var tagValue = String.Join(tag.FlowSeparator, valueList);
        
        if (!String.IsNullOrWhiteSpace(templateResultString) && !String.IsNullOrWhiteSpace(tag.FullString))
            templateResultString =
                templateResultString.Replace(tag.FullString, tagValue);
        object? newResultObject = null;
        if (oneTagOnlyTemplate)
        {
            if (
                (templateResultObject is string && !String.IsNullOrWhiteSpace(templateResultObject as string))
                || (templateResultObject is not string && templateResultObject is not null)
            )
            {
                newResultObject = templateResultString;
            }
            else if (columnObject is not null && valueList.Count == 1)
            {
                newResultObject = columnObject;
                //newResultObject = TryExractValue(templateResultString, dataSource.Columns[columnIndex]);
            }
            else
            {
                newResultObject = templateResultString;
            }
        }
        else
        {
            newResultObject = templateResultString;
        }


        return (templateResultString, newResultObject);
    }
}
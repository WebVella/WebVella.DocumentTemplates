using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;

namespace WebVella.DocumentTemplates.Core.Utility;

public partial class WvTemplateUtility
{
    public (string?, object?) ReplaceAllValuesDataTagInTemplate(string? templateResultString,
        object? templateResultObject, WvTemplateTag tag, DataTable dataSource)
    {
        if (String.IsNullOrWhiteSpace(templateResultString)) return (templateResultString, templateResultObject);
        var columnIndices = GetColumnsIndexFromTagItemName(tag.ItemName, dataSource);

        if (columnIndices.Count == 0 || dataSource.Rows.Count == 0)
        {
            if (templateResultObject is null)
                templateResultObject = templateResultString;
            return (templateResultString, templateResultObject);
        }

        List<int> rowIndices = new();
        if (tag.IndexList.Count > 0)
        {
            foreach (var tagIndex in tag.IndexList)
            {
                if (dataSource.Rows.Count - 1 >= tagIndex)
                    rowIndices.Add(tagIndex);
            }
            //if all requested indexes are not present add the first row
            if(rowIndices.Count == 0)
                rowIndices.Add(0);
        }
        else
        {
            for (int i = 0; i < dataSource.Rows.Count; i++)
            {
                rowIndices.Add(i);
            }            
        }

        var tagContentList = new List<string>();
        object? columnObject = null;
        foreach (var rowIndex in rowIndices)
        {
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
                    tagContentList.Add(columnValue);
            }
        }

        if (!String.IsNullOrWhiteSpace(templateResultString)
            && !String.IsNullOrWhiteSpace(tag.FullString)
            && tagContentList.Count > 0)
        {
            var separator = tag.FlowSeparator;
            templateResultString = templateResultString.Replace(tag.FullString, String.Join(separator, tagContentList));
        }

        object? newResultObject = templateResultString;
        return (templateResultString, newResultObject);
    }
}
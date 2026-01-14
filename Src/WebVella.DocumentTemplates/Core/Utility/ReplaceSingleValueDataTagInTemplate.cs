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

        List<int> rowIndices = new();
        //If no rowIndex is requested
        if (contextRowIndex is null)
        {
            //Add only the requested indexes in the tag
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
            //add all rows to the processed index
            else
            {
                for (int i = 0; i < dataSource.Rows.Count; i++)
                {
                    rowIndices.Add(i);
                }
            }
        }
        //If rowIndex is requested
        else
        {
            //If tag has requested indexes force them
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
            //if tag does not request indexes add the processed one
            else if (dataSource.Rows.Count - 1 >= contextRowIndex.Value)
            {
                rowIndices.Add(contextRowIndex.Value);
            }
            else
                rowIndices.Add(0);
        }

        var valueList = new List<string>();
        object? columnObject = null;
        var oneTagOnlyTemplate = templateResultString?.ToLowerInvariant() == tag.FullString?.ToLowerInvariant();

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
                    valueList.Add(columnValue);
            }
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
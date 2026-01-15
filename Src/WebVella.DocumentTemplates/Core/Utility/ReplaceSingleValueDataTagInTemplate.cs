using System.Data;
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
            if (tag.IndexGroups.Count > 0)
            {
                foreach (var tagIndex in tag.IndexGroups[0].Indexes)
                {
                    if (dataSource.Rows.Count - 1 >= tagIndex)
                        rowIndices.Add(tagIndex);
                }
                //if all requested indexes are not present add the first row
                if (rowIndices.Count == 0)
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
            if (tag.IndexGroups.Count > 0)
            {
                foreach (var tagIndex in tag.IndexGroups[0].Indexes)
                {
                    if (dataSource.Rows.Count - 1 >= tagIndex)
                        rowIndices.Add(tagIndex);
                }
                //if all requested indexes are not present add the first row
                if (rowIndices.Count == 0)
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
        DataColumn? column = null;
        object? columnObject = null;
        var oneTagOnlyTemplate = templateResultString?.ToLowerInvariant() == tag.FullString?.ToLowerInvariant();
        foreach (var rowIndex in rowIndices)
        {
            foreach (var columnIndex in columnIndices)
            {
                columnObject = dataSource.Rows[rowIndex][columnIndex];
                column = dataSource.Columns[columnIndex];
                var columnValue = columnObject?.ToString();
                var (isEnumarable, type) = CheckEnumerable(column);
                if (isEnumarable)
                {
                    var valueListInString = ((IEnumerable<object>)columnObject).Select(x => x?.ToString()).ToList();
                    //Calculate requested value indexes
                    //The second index group is applicable in this case
                    if (tag.IndexGroups.Count > 1)
                    {
                        var valueListByIndex = new List<string?>();
                        foreach (var valIndex in tag.IndexGroups[1].Indexes)
                        {
                            if (valueListInString.Count <= valIndex + 1)
                            {
                                valueListByIndex.Add(valueListInString[valIndex]);
                            }
                        }
                        valueListInString = valueListByIndex;
                    }
                    //all values are eligible
                    //Second separator should be used, if not present the first
                    var flowSeparator = tag.FlowSeparatorList.Count >= 2 ? tag.FlowSeparatorList[1] : tag.FlowSeparator;
                    columnValue = String.Join(flowSeparator, valueListInString);
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
        if (oneTagOnlyTemplate && columnIndices.Count == 1)
        {
            if (
                (templateResultObject is string && !String.IsNullOrWhiteSpace(templateResultObject as string))
                || (templateResultObject is not string && templateResultObject is not null)
            )
            {
                newResultObject = templateResultString;
            }
            else if (column is not null && valueList.Count == 1)
            {
                var (isEnumarable, type) = CheckEnumerable(column);
                if (isEnumarable)
                    newResultObject = templateResultString;
                else
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
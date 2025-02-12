using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;

namespace WebVella.DocumentTemplates.Core.Utility;
public static partial class WvTemplateUtility
{
	public static (string?, object?) ReplaceSingleValueDataTagInTemplate(string? templateResultString, object? templateResultObject, WvTemplateTag tag, DataTable dataSource, int? contextRowIndex)
	{

		int columnIndex = -1;
		foreach (DataColumn col in dataSource.Columns)
		{
			if (col.ColumnName.ToLowerInvariant() == tag.Name)
			{
				columnIndex = dataSource.Columns.IndexOf(col);
				break;
			}
		}
		if (columnIndex == -1) return (templateResultString, templateResultObject);
		if (dataSource.Rows.Count == 0) return (templateResultString, templateResultObject);
		int rowIndex = 0;
		if (tag.IndexList is not null && tag.IndexList.Count > 0)
		{
			rowIndex = tag.IndexList[0];
		}
		else if (contextRowIndex is not null && dataSource.Rows.Count - 1 >= contextRowIndex)
		{
			rowIndex = contextRowIndex.Value;
		}
		var indexCanBeApplied = (dataSource.Rows.Count >= rowIndex + 1) && (dataSource.Columns.Count >= columnIndex + 1);
		var oneTagOnlyTemplate = templateResultString?.ToLowerInvariant() == tag.FullString?.ToLowerInvariant();

		if (!String.IsNullOrWhiteSpace(templateResultString)
					&& !String.IsNullOrWhiteSpace(tag.FullString)
					&& indexCanBeApplied)
			templateResultString = templateResultString.Replace(tag.FullString, dataSource.Rows[rowIndex][columnIndex]?.ToString());

		object? newResultObject = null;
		if (oneTagOnlyTemplate)
		{
			if (templateResultObject is not null)
			{
				newResultObject = templateResultString;
			}
			else if (indexCanBeApplied)
			{
				newResultObject = dataSource.Rows[rowIndex][columnIndex];
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

using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;

namespace WebVella.DocumentTemplates.Core.Utility;
public partial class WvTemplateUtility
{
	public static (string?, object?) ReplaceAllValuesDataTagInTemplate(string? templateResultString, object? templateResultObject, WvTemplateTag tag, DataTable dataSource)
	{
		if(String.IsNullOrWhiteSpace(templateResultString)) return (templateResultString, templateResultObject);

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

		var oneTagOnlyTemplate = templateResultString?.ToLowerInvariant() == tag.FullString?.ToLowerInvariant();

		var tagContentList = new List<string>();
		for (int rowIndex = 0; rowIndex < dataSource.Rows.Count; rowIndex++)
		{
			var value = dataSource.Rows[rowIndex][columnIndex]?.ToString();
			if (!String.IsNullOrWhiteSpace(value))
			{
				tagContentList.Add(value);
			}
		}

		if (!String.IsNullOrWhiteSpace(templateResultString) 
			&&	!String.IsNullOrWhiteSpace(tag.FullString) 
			&& tagContentList.Count > 0)
		{
			var separator = tag.FlowSeparator ?? "";
			templateResultString = templateResultString.Replace(tag.FullString, String.Join(separator, tagContentList));
		}
		object? newResultObject = templateResultString;
		return (templateResultString, newResultObject);
	}



}

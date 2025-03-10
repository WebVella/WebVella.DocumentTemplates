using System.Data;
using System.Text.RegularExpressions;

namespace WebVella.DocumentTemplates.Core.Utility;
public partial class WvTemplateUtility
{
	public int GetColumnIndexFromTagName(string? tagName, DataTable dataSource)
	{
		int result = -1;
		if (String.IsNullOrWhiteSpace(tagName)) return result;

		tagName = tagName.Trim().ToLowerInvariant();

		foreach (DataColumn col in dataSource.Columns)
		{
			if (col.ColumnName.ToLowerInvariant() == tagName)
			{
				result = dataSource.Columns.IndexOf(col);
				break;
			}
		}
		return result;
	}
}

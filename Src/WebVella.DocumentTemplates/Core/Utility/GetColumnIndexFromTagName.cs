using System.Data;
using System.Text.RegularExpressions;

namespace WebVella.DocumentTemplates.Core.Utility;
public partial class WvTemplateUtility
{
	public List<int> GetColumnsIndexFromTagItemName(string? itemName, DataTable dataSource)
	{
		List<int> result = new();
		if (String.IsNullOrWhiteSpace(itemName)) return result;

		itemName = itemName.Trim().ToLowerInvariant();

		//Check for exact column 
		foreach (DataColumn col in dataSource.Columns)
		{
			if (col.ColumnName.ToLowerInvariant() == itemName)
			{
				result.Add(dataSource.Columns.IndexOf(col)); 
				break;
			}
		}
		if (result.Count > 0) return result;
		
		//Check if it is not a startswith string
		foreach (DataColumn col in dataSource.Columns)
		{
			if (col.ColumnName.ToLowerInvariant().StartsWith(itemName))
			{
				result.Add(dataSource.Columns.IndexOf(col)); 
			}
		}		
		
		return result;
	}
}

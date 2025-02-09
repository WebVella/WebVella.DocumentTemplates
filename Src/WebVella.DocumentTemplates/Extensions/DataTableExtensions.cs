using System.Data;
using System.Text;

namespace WebVella.DocumentTemplates.Extensions;
public static class DataTableExtensions
{
	public static DataTable CreateAsEmpty(this DataTable originalTable)
	{
		if (originalTable is null) throw new ArgumentNullException(nameof(originalTable));
		// Create a new DataTable with the same structure as the original table
		DataTable newTable = new DataTable();

		foreach (DataColumn column in originalTable.Columns)
		{
			newTable.Columns.Add(column.ColumnName, column.DataType);
		}
		return newTable;
	}
	public static DataTable CreateNew(this DataTable originalTable, List<int>? rowIndices = null)
	{
		if (originalTable is null) throw new ArgumentNullException(nameof(originalTable));
		// Create a new DataTable with the same structure as the original table
		DataTable newTable = new DataTable();

		foreach (DataColumn column in originalTable.Columns)
		{
			newTable.Columns.Add(column.ColumnName, column.DataType);
		}

		// Check if there is at least one row in the original table
		if (originalTable.Rows.Count > 0)
		{
			var currentIndex = -1;
			foreach (DataRow row in originalTable.Rows)
			{
				currentIndex++;
				if (rowIndices is not null && !rowIndices.Contains(currentIndex)) continue;
				DataRow newRow = newTable.NewRow();

				foreach (DataColumn column in originalTable.Columns)
				{
					newRow[column.ColumnName] = row[column.ColumnName];
				}
				newTable.Rows.Add(newRow);
			}
		}
		return newTable;
	}
	public static List<DataTable> GroupBy(this DataTable originalTable, List<string> groupColumns)
	{
		var result = new List<DataTable>();
		if (groupColumns is null || groupColumns.Count == 0
			|| originalTable.Columns.Count == 0
			|| originalTable.Rows.Count == 0)
		{
			result.Add(originalTable);
			return result;
		}

		var groupDict = new Dictionary<string, DataTable>();
		var dtColumnsHash = new HashSet<string>();
		foreach (DataColumn column in originalTable.Columns)
		{
			if (String.IsNullOrEmpty(column.ColumnName)) continue;

			if (originalTable.CaseSensitive)
				dtColumnsHash.Add(column.ColumnName);
			else
				dtColumnsHash.Add(column.ColumnName.ToLowerInvariant());
		}
		foreach (DataRow row in originalTable.Rows)
		{
			var sbKey = new StringBuilder();
			foreach (var column in groupColumns)
			{
				if (String.IsNullOrWhiteSpace(column)) continue;
				var groupColumName = originalTable.CaseSensitive ? column : column.ToLowerInvariant();
				if (dtColumnsHash.Contains(groupColumName))
					sbKey.Append($"{row[groupColumName]}$$$|||$$$");
			}
			var key = sbKey.ToString();
			if (!groupDict.ContainsKey(key))
			{
				groupDict.Add(key, originalTable.CreateAsEmpty());
			}
			var newRow = groupDict[key].NewRow();
			foreach (DataColumn column in originalTable.Columns)
			{
				newRow[column.ColumnName] = row[column.ColumnName];
			}

			groupDict[key].Rows.Add(newRow);
		}

		foreach (string groupHash in groupDict.Keys)
		{
			result.Add(groupDict[groupHash]);
		}

		return result;
	}

}

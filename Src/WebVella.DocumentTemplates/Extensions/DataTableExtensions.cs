using System.Data;

namespace WebVella.DocumentTemplates.Extensions;
public static class DataTableExtensions
{
	public static DataTable CreateNew(this DataTable originalTable, List<int>? rowIndices = null)
	{
		if(originalTable is null) throw new ArgumentNullException(nameof(originalTable));
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

}

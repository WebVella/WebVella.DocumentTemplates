using System.Data;
using System.Text;
using WebVella.DocumentTemplates.Extensions;
using WebVella.DocumentTemplates.Tests.Models;
namespace WebVella.DocumentTemplates.Tests.Extensions;
public class DataTableExtensionsTests
{
	public DataTable SampleData;

	public DataTableExtensionsTests()
	{
		//Data
		{
			var ds = new DataTable();
			ds.Columns.Add("position", typeof(int));
			ds.Columns.Add("sku", typeof(string));
			ds.Columns.Add("name", typeof(string));
			ds.Columns.Add("price", typeof(decimal));

			for (int i = 0; i < 5; i++)
			{
				var position = i + 1;
				var dsrow = ds.NewRow();
				dsrow["position"] = position;
				dsrow["sku"] = $"sku{position}";
				dsrow["name"] = $"item{position}";
				dsrow["price"] = (decimal)(position * 10);
				ds.Rows.Add(dsrow);
			}
			SampleData = ds;
		}
	}


	[Fact]
	public void Type_ShouldReturnEmptyIfEmpty()
	{
		DataTable? originalDt = new DataTable();
		DataTable? resultDt = originalDt.CreateNew();
		Assert.NotNull(resultDt);
		Assert.Empty(resultDt.Rows);
		Assert.Empty(resultDt.Columns);
	}

	[Fact]
	public void Type_ShouldReturnClone()
	{
		DataTable? originalDt = SampleData;
		DataTable? resultDt = originalDt.CreateNew();
		Assert.NotNull(resultDt);
		Assert.NotEmpty(resultDt.Rows);
		Assert.NotEmpty(resultDt.Columns);

		for (int i = 0; i < originalDt.Columns.Count; i++)
		{
			Assert.Equal(originalDt.Columns[i].ColumnName, resultDt.Columns[i].ColumnName);
			Assert.Equal(originalDt.Columns[i].DataType.FullName, resultDt.Columns[i].DataType.FullName);
		}

		for (int i = 0; i < originalDt.Rows.Count; i++)
		{
			foreach (DataColumn col in originalDt.Columns)
			{
				Assert.Equal(originalDt.Rows[i][col.ColumnName], resultDt.Rows[i][col.ColumnName]);
			}
		}
	}

	[Fact]
	public void Type_ShouldReturnOnlySelectedRows()
	{
		var rowIndices = new List<int>{ 1, 2};
		DataTable? originalDt = SampleData;
		DataTable? resultDt = originalDt.CreateNew(rowIndices);
		Assert.NotNull(resultDt);
		Assert.NotEmpty(resultDt.Rows);
		Assert.NotEmpty(resultDt.Columns);

		for (int i = 0; i < originalDt.Columns.Count; i++)
		{
			Assert.Equal(originalDt.Columns[i].ColumnName, resultDt.Columns[i].ColumnName);
			Assert.Equal(originalDt.Columns[i].DataType.FullName, resultDt.Columns[i].DataType.FullName);
		}
		List<string> resultRowHash = new List<string>();
		for (int i = 0; i < resultDt.Rows.Count; i++)
		{
			var rowHashSB = new StringBuilder();
			foreach (DataColumn col in resultDt.Columns)
			{
				string colStringValue = resultDt.Rows[i][col.ColumnName]?.ToString() ?? String.Empty;
				rowHashSB.Append(colStringValue);
			}
			resultRowHash.Add(rowHashSB.ToString());
		}

		var currentResultRowIndex = 0;
		for (int i = 0; i < originalDt.Rows.Count; i++)
		{
			if(!rowIndices.Contains(i)) continue;

			var rowHashSB = new StringBuilder();
			foreach (DataColumn col in originalDt.Columns)
			{
				string colStringValue = originalDt.Rows[i][col.ColumnName]?.ToString() ?? String.Empty;
				rowHashSB.Append(colStringValue);
			}

			Assert.Equal(rowHashSB.ToString(),resultRowHash[currentResultRowIndex]);
			currentResultRowIndex++;
		}
	}

}
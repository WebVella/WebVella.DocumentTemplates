using System.Data;
using System.Text;
using WebVella.DocumentTemplates.Extensions;
using WebVella.DocumentTemplates.Tests.Models;
namespace WebVella.DocumentTemplates.Tests.Extensions;
public class DataTableExtensionsTests
{
	public DataTable SampleData;
	public DataTable GroupData;

	public DataTableExtensionsTests()
	{
		//SampleData
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
		{
			var ds = new DataTable();
			ds.Columns.Add("column1", typeof(int));
			ds.Columns.Add("column2", typeof(string));
			ds.Columns.Add("column3", typeof(string));
			ds.Columns.Add("column4", typeof(decimal));
			for (int i = 0; i < 5; i++)
			{
				var position = i + 1;
				var dsrow = ds.NewRow();
				dsrow["column1"] = position;
				dsrow["column2"] = $"data{position}";
				dsrow["column3"] = $"item{position}";
				dsrow["column4"] = (decimal)position;
				ds.Rows.Add(dsrow);
			}

			GroupData = ds;
		}
	}

	#region << CreateAsEmpty >>
	[Fact]
	public void CreateAsEmpty_ShouldReturnEmptyIfEmpty()
	{
		DataTable? originalDt = new DataTable();
		DataTable? resultDt = originalDt.CreateAsEmpty();
		Assert.NotNull(resultDt);
		Assert.Empty(resultDt.Rows);
		Assert.Empty(resultDt.Columns);
	}

	[Fact]
	public void CreateAsEmpty_ShouldCreateWithoutData()
	{
		DataTable? originalDt = SampleData;
		DataTable? resultDt = originalDt.CreateAsEmpty();
		Assert.NotNull(resultDt);
		Assert.NotEmpty(resultDt.Columns);
		Assert.Empty(resultDt.Rows);
	}
	#endregion

	#region << CreateAsnew >>
	[Fact]
	public void Type_ShouldReturnEmptyIfEmpty()
	{
		DataTable? originalDt = new DataTable();
		DataTable? resultDt = originalDt.CreateAsNew();
		Assert.NotNull(resultDt);
		Assert.Empty(resultDt.Rows);
		Assert.Empty(resultDt.Columns);
	}

	[Fact]
	public void Type_ShouldReturnClone()
	{
		DataTable? originalDt = SampleData;
		DataTable? resultDt = originalDt.CreateAsNew();
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
		var rowIndices = new List<int> { 1, 2 };
		DataTable? originalDt = SampleData;
		DataTable? resultDt = originalDt.CreateAsNew(rowIndices);
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
			if (!rowIndices.Contains(i)) continue;

			var rowHashSB = new StringBuilder();
			foreach (DataColumn col in originalDt.Columns)
			{
				string colStringValue = originalDt.Rows[i][col.ColumnName]?.ToString() ?? String.Empty;
				rowHashSB.Append(colStringValue);
			}

			Assert.Equal(rowHashSB.ToString(), resultRowHash[currentResultRowIndex]);
			currentResultRowIndex++;
		}
	}
	#endregion

	#region << Group By >>
	[Fact]
	public void GroupBy_ShouldReturnEmptyIfEmpty()
	{
		DataTable originalDt = new DataTable();
		var groupColumns = new List<string> { "column1", "column2" };
		List<DataTable> resultDt = originalDt.GroupBy(groupColumns);
		Assert.NotNull(resultDt);
		Assert.Single(resultDt);
		Assert.Empty(resultDt[0].Rows);
		Assert.Empty(resultDt[0].Columns);
	}
	[Fact]
	public void GroupBy_ShouldReturnSameOnceIfWrongColumnGroup()
	{
		DataTable originalDt = GroupData;
		var groupColumns = new List<string> { Guid.NewGuid().ToString() };
		List<DataTable> resultDt = originalDt.GroupBy(groupColumns);
		Assert.NotNull(resultDt);
		Assert.Single(resultDt);
		Assert.NotEmpty(resultDt[0].Rows);
		Assert.NotEmpty(resultDt[0].Columns);
		Assert.Equal(GroupData.Columns.Count, resultDt[0].Columns.Count);
		Assert.Equal(GroupData.Rows.Count, resultDt[0].Rows.Count);
		var rowIndex = 0;
		foreach (DataRow row in GroupData.Rows)
		{
			foreach (DataColumn column in GroupData.Columns)
			{
				Assert.Equal(GroupData.Rows[rowIndex][column.ColumnName], resultDt[0].Rows[rowIndex][column.ColumnName]);
			}
			rowIndex++;
		}
	}

	[Fact]
	public void GroupBy_ShouldReturnSameOnceIfEmptyColumnGroup()
	{
		DataTable originalDt = GroupData;
		var groupColumns = new List<string> { String.Empty };
		List<DataTable> resultDt = originalDt.GroupBy(groupColumns);
		Assert.NotNull(resultDt);
		Assert.Single(resultDt);
		Assert.NotEmpty(resultDt[0].Rows);
		Assert.NotEmpty(resultDt[0].Columns);
		Assert.Equal(GroupData.Columns.Count, resultDt[0].Columns.Count);
		Assert.Equal(GroupData.Rows.Count, resultDt[0].Rows.Count);
		var rowIndex = 0;
		foreach (DataRow row in GroupData.Rows)
		{
			foreach (DataColumn column in GroupData.Columns)
			{
				Assert.Equal(GroupData.Rows[rowIndex][column.ColumnName], resultDt[0].Rows[rowIndex][column.ColumnName]);
			}
			rowIndex++;
		}
	}

	[Fact]
	public void GroupBy_ShouldReturnSameCaseSensetive1_NoMatch()
	{
		DataTable originalDt = GroupData;
		originalDt.CaseSensitive = true;
		originalDt.Rows[1]["column1"] = originalDt.Rows[0]["column1"];
		var groupColumns = new List<string> { "Column1" };
		List<DataTable> resultDt = originalDt.GroupBy(groupColumns);
		Assert.NotNull(resultDt);
		Assert.Single(resultDt);
		Assert.NotEmpty(resultDt[0].Rows);
		Assert.NotEmpty(resultDt[0].Columns);
		Assert.Equal(GroupData.Columns.Count, resultDt[0].Columns.Count);
		Assert.Equal(GroupData.Rows.Count, resultDt[0].Rows.Count);
		var rowIndex = 0;
		foreach (DataRow row in GroupData.Rows)
		{
			foreach (DataColumn column in GroupData.Columns)
			{
				Assert.Equal(GroupData.Rows[rowIndex][column.ColumnName], resultDt[0].Rows[rowIndex][column.ColumnName]);
			}
			rowIndex++;
		}
	}

	[Fact]
	public void GroupBy_ShouldReturnSameCaseSensetive1_Match()
	{
		DataTable originalDt = GroupData;
		originalDt.CaseSensitive = false;
		originalDt.Rows[1]["column1"] = originalDt.Rows[0]["column1"];
		var groupColumns = new List<string> { "Column1" };
		List<DataTable> resultDt = originalDt.GroupBy(groupColumns);
		Assert.NotNull(resultDt);
		Assert.Equal(4, resultDt.Count);
		Assert.NotEmpty(resultDt[0].Rows);
		Assert.NotEmpty(resultDt[0].Columns);
		Assert.Equal(2, resultDt[0].Rows.Count);

	}

	[Fact]
	public void GroupBy_MultiGroup()
	{
		DataTable originalDt = GroupData;
		var groupColumns = new List<string> { "column1", "column2", "column3", "column4" };
		List<DataTable> resultDt = originalDt.GroupBy(groupColumns);
		Assert.NotNull(resultDt);
		Assert.Equal(5, resultDt.Count);
	}

	[Fact]
	public void GroupBy_MultiGroupWithWrong()
	{
		DataTable originalDt = GroupData;
		var groupColumns = new List<string> { "column1", Guid.NewGuid().ToString(), "column2", "column3", "column4" };
		List<DataTable> resultDt = originalDt.GroupBy(groupColumns);
		Assert.NotNull(resultDt);
		Assert.Equal(5, resultDt.Count);
	}

	[Fact]
	public void GroupBy_MultiGroup2()
	{
		DataTable originalDt = GroupData;
		originalDt.CaseSensitive = false;
		for (int i = 0; i < 5; i++)
		{
			if (i == 0) continue;
			originalDt.Rows[i]["column1"] = originalDt.Rows[0]["column1"];
			originalDt.Rows[i]["column2"] = originalDt.Rows[0]["column2"];
			originalDt.Rows[i]["column4"] = originalDt.Rows[0]["column4"];
		}

		var groupColumns = new List<string> { "column1", "column2", "column4" };
		List<DataTable> resultDt = originalDt.GroupBy(groupColumns);
		Assert.NotNull(resultDt);
		Assert.Single(resultDt);
		Assert.NotEmpty(resultDt[0].Rows);
		Assert.NotEmpty(resultDt[0].Columns);
		Assert.Equal(5, resultDt[0].Rows.Count);

	}

	#endregion
}
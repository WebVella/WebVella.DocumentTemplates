using WebVella.DocumentTemplates.Engines.ExcelFile.Utility;
using WebVella.DocumentTemplates.Tests.Models;

namespace WebVella.DocumentTemplates.Tests.Engines;
public partial class UtilitiesExcelEngineTests : TestBase
{
	private static readonly object locker = new();
	public UtilitiesExcelEngineTests() : base() { }

	#region << GetRangeFromString >>

	//GetRangeFromString
	[Fact]
	public void GetRangeFromString_Test1()
	{
		lock (locker)
		{
			//Given
			var address = "name!A1:C4";

			//When
			var result = new WvExcelRangeHelpers().GetRangeFromString(address);
			//Then
			Assert.NotNull(result);
			Assert.Equal("name", result.Worksheet);
			Assert.Equal(1, result.FirstRow);
			Assert.Equal(1, result.FirstColumn);
			Assert.Equal(4, result.LastRow);
			Assert.Equal(3, result.LastColumn);
		}
	}

	[Fact]
	public void GetRangeFromString_Test2()
	{
		lock (locker)
		{
			//Given
			var address = "name!C5";

			//When
			var result = new WvExcelRangeHelpers().GetRangeFromString(address);
			//Then
			Assert.NotNull(result);
			Assert.Equal("name", result.Worksheet);
			Assert.Equal(5, result.FirstRow);
			Assert.Equal(3, result.FirstColumn);
			Assert.Equal(5, result.LastRow);
			Assert.Equal(3, result.LastColumn);
		}
	}

	[Fact]
	public void GetRangeFromString_Test3()
	{
		lock (locker)
		{
			//Given
			var address = "A1:C4";

			//When
			var result = new WvExcelRangeHelpers().GetRangeFromString(address);
			//Then
			Assert.NotNull(result);
			Assert.True(String.IsNullOrWhiteSpace(result.Worksheet));
			Assert.Equal(1, result.FirstRow);
			Assert.Equal(1, result.FirstColumn);
			Assert.Equal(4, result.LastRow);
			Assert.Equal(3, result.LastColumn);
		}
	}

	[Fact]
	public void GetRangeFromString_Test4()
	{
		lock (locker)
		{
			//Given
			var address = "C5";

			//When
			var result = new WvExcelRangeHelpers().GetRangeFromString(address);
			//Then
			Assert.NotNull(result);
			Assert.True(String.IsNullOrWhiteSpace(result.Worksheet));
			Assert.Equal(5, result.FirstRow);
			Assert.Equal(3, result.FirstColumn);
			Assert.Equal(5, result.LastRow);
			Assert.Equal(3, result.LastColumn);
		}
	}

	[Fact]
	public void GetRangeFromString_Test5()
	{
		lock (locker)
		{
			//Given
			var address = "invalid";

			//When
			var result = new WvExcelRangeHelpers().GetRangeFromString(address);
			//Then
			Assert.Null(result);
		}
	}

	[Fact]
	public void GetRangeFromString_Test6()
	{
		lock (locker)
		{
			//Given
			var address = "invalid123";

			//When
			var result = new WvExcelRangeHelpers().GetRangeFromString(address);
			//Then
			Assert.Null(result);
		}
	}

	[Fact]
	public void GetRangeFromString_Lock1()
	{
		lock (locker)
		{
			//Given
			var address = "$A1";

			//When
			var result = new WvExcelRangeHelpers().GetRangeFromString(address);
			//Then
			Assert.NotNull(result);

			Assert.Equal(1, result!.FirstColumn);
			Assert.True(result!.FirstColumnLocked);
			Assert.Equal(1, result!.FirstRow);
			Assert.False(result!.FirstRowLocked);

			Assert.Equal(1, result!.LastColumn);
			Assert.True(result!.LastColumnLocked);
			Assert.Equal(1, result!.LastRow);
			Assert.False(result!.LastRowLocked);

		}
	}

	[Fact]
	public void GetRangeFromString_Lock2()
	{
		lock (locker)
		{
			//Given
			var address = "A$1";

			//When
			var result = new WvExcelRangeHelpers().GetRangeFromString(address);
			//Then
			Assert.NotNull(result);

			Assert.Equal(1, result!.FirstColumn);
			Assert.False(result!.FirstColumnLocked);
			Assert.Equal(1, result!.FirstRow);
			Assert.True(result!.FirstRowLocked);

			Assert.Equal(1, result!.LastColumn);
			Assert.False(result!.LastColumnLocked);
			Assert.Equal(1, result!.LastRow);
			Assert.True(result!.LastRowLocked);
		}
	}
	[Fact]
	public void GetRangeFromString_Lock3()
	{
		lock (locker)
		{
			//Given
			var address = "$A$1";

			//When
			var result = new WvExcelRangeHelpers().GetRangeFromString(address);
			//Then
			Assert.NotNull(result);

			Assert.Equal(1, result!.FirstColumn);
			Assert.True(result!.FirstColumnLocked);
			Assert.Equal(1, result!.FirstRow);
			Assert.True(result!.FirstRowLocked);

			Assert.Equal(1, result!.LastColumn);
			Assert.True(result!.LastColumnLocked);
			Assert.Equal(1, result!.LastRow);
			Assert.True(result!.LastRowLocked);
		}
	}

	[Fact]
	public void GetRangeFromString_Lock4()
	{
		lock (locker)
		{
			//Given
			var address = "$A$1:B1";

			//When
			var result = new WvExcelRangeHelpers().GetRangeFromString(address);
			//Then
			Assert.NotNull(result);

			Assert.Equal(1, result!.FirstColumn);
			Assert.True(result!.FirstColumnLocked);
			Assert.Equal(1, result!.FirstRow);
			Assert.True(result!.FirstRowLocked);

			Assert.Equal(2, result!.LastColumn);
			Assert.False(result!.LastColumnLocked);
			Assert.Equal(1, result!.LastRow);
			Assert.False(result!.LastRowLocked);

		}
	}

	[Fact]
	public void GetRangeFromString_Lock5()
	{
		lock (locker)
		{
			//Given
			var address = "$A$1:$B$1";

			//When
			var result = new WvExcelRangeHelpers().GetRangeFromString(address);
			//Then
			Assert.NotNull(result);

			Assert.Equal(1, result!.FirstColumn);
			Assert.True(result!.FirstColumnLocked);
			Assert.Equal(1, result!.FirstRow);
			Assert.True(result!.FirstRowLocked);

			Assert.Equal(2, result!.LastColumn);
			Assert.True(result!.LastColumnLocked);
			Assert.Equal(1, result!.LastRow);
			Assert.True(result!.LastRowLocked);
		}
	}

	#endregion
}

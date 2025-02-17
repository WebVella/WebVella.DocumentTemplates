using ClosedXML.Excel;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Engines.Excel;
using WebVella.DocumentTemplates.Engines.Excel.Utility;
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
			var result = WvExcelRangeHelpers.GetRangeFromString(address);
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
			var result = WvExcelRangeHelpers.GetRangeFromString(address);
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
			var result = WvExcelRangeHelpers.GetRangeFromString(address);
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
			var result = WvExcelRangeHelpers.GetRangeFromString(address);
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
			var result = WvExcelRangeHelpers.GetRangeFromString(address);
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
			var result = WvExcelRangeHelpers.GetRangeFromString(address);
			//Then
			Assert.Null(result);
		}
	}

	#endregion
}

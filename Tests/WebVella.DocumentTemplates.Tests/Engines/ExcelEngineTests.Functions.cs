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
public partial class ExcelEngineTests : TestBase
{
	#region << SUM >>
	//in the general tests
	#endregion
	#region << AVERAGE >>
	[Fact]
	public void Function_AVERAGE_SingleRange()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateFunction-Average-1.xlsx";
			var template = new WvExcelFileTemplate
			{
				Template = LoadWorkbook(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvExcelFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			GeneralResultChecks(result);
			Assert.Single(result!.Template!.Worksheets);
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Result);
			Assert.Single(result!.ResultItems[0]!.Result!.Worksheets);
			var worksheet = result!.ResultItems[0]!.Result!.Worksheets.First();
			Assert.Equal("position", worksheet.Cell(1, 1).Value.ToString());
			Assert.Equal("1", worksheet.Cell(2, 1).Value.ToString());
			Assert.Equal("2", worksheet.Cell(3, 1).Value.ToString());
			Assert.Equal("3", worksheet.Cell(4, 1).Value.ToString());
			Assert.Equal("4", worksheet.Cell(5, 1).Value.ToString());
			Assert.Equal("5", worksheet.Cell(6, 1).Value.ToString());
			Assert.Equal("3", worksheet.Cell(7, 1).Value.ToString());
			Assert.Equal("item1 AVERAGE: 3", worksheet.Cell(8, 1).Value.ToString());
			SaveWorkbook(result!.ResultItems[0]!.Result!, templateFile);
		}
	}
	#endregion

	#region << MIN >>
	[Fact]
	public void Function_MIN_SingleRange()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateFunction-MIN-1.xlsx";
			var template = new WvExcelFileTemplate
			{
				Template = LoadWorkbook(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvExcelFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			GeneralResultChecks(result);
			Assert.Single(result!.Template!.Worksheets);
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Result);
			Assert.Single(result!.ResultItems[0]!.Result!.Worksheets);
			var worksheet = result!.ResultItems[0]!.Result!.Worksheets.First();
			Assert.Equal("position", worksheet.Cell(1, 1).Value.ToString());
			Assert.Equal("1", worksheet.Cell(2, 1).Value.ToString());
			Assert.Equal("2", worksheet.Cell(3, 1).Value.ToString());
			Assert.Equal("3", worksheet.Cell(4, 1).Value.ToString());
			Assert.Equal("4", worksheet.Cell(5, 1).Value.ToString());
			Assert.Equal("5", worksheet.Cell(6, 1).Value.ToString());
			Assert.Equal("1", worksheet.Cell(7, 1).Value.ToString());
			Assert.Equal("item1 MIN: 1", worksheet.Cell(8, 1).Value.ToString());
			SaveWorkbook(result!.ResultItems[0]!.Result!, templateFile);
		}
	}
	#endregion

	#region << MAX >>
	[Fact]
	public void Function_MAX_SingleRange()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateFunction-MAX-1.xlsx";
			var template = new WvExcelFileTemplate
			{
				Template = LoadWorkbook(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvExcelFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			GeneralResultChecks(result);
			Assert.Single(result!.Template!.Worksheets);
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Result);
			Assert.Single(result!.ResultItems[0]!.Result!.Worksheets);
			var worksheet = result!.ResultItems[0]!.Result!.Worksheets.First();
			Assert.Equal("position", worksheet.Cell(1, 1).Value.ToString());
			Assert.Equal("1", worksheet.Cell(2, 1).Value.ToString());
			Assert.Equal("2", worksheet.Cell(3, 1).Value.ToString());
			Assert.Equal("3", worksheet.Cell(4, 1).Value.ToString());
			Assert.Equal("4", worksheet.Cell(5, 1).Value.ToString());
			Assert.Equal("5", worksheet.Cell(6, 1).Value.ToString());
			Assert.Equal("5", worksheet.Cell(7, 1).Value.ToString());
			Assert.Equal("item1 MAX: 5", worksheet.Cell(8, 1).Value.ToString());
			SaveWorkbook(result!.ResultItems[0]!.Result!, templateFile);
		}
	}
	#endregion
}

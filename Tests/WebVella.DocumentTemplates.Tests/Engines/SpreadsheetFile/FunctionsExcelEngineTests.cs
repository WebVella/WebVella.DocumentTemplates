using WebVella.DocumentTemplates.Engines.SpreadsheetFile;
using WebVella.DocumentTemplates.Tests.Models;
using WebVella.DocumentTemplates.Tests.Utils;

namespace WebVella.DocumentTemplates.Tests.Engines;
public partial class FunctionsSpreadsheetEngineTests : TestBase
{
	private static readonly object locker = new();
	public FunctionsSpreadsheetEngineTests() : base() { }

	#region << General >>
	[Fact]
	public void Function_SingleRange()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateFunction1.xlsx";
			var template = new WvSpreadsheetFileTemplate
			{
				Template = new TestUtils().LoadFileAsStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvSpreadsheetFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			new TestUtils().GeneralResultChecks(result);
			Assert.Single(result!.Workbook!.Worksheets);
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Workbook);
			Assert.Single(result!.ResultItems[0]!.Workbook!.Worksheets);
			var worksheet = result!.ResultItems[0]!.Workbook!.Worksheets.First();
			Assert.Equal("position", worksheet.Cell(1, 1).Value.ToString());
			Assert.Equal("1", worksheet.Cell(2, 1).Value.ToString());
			Assert.Equal("2", worksheet.Cell(3, 1).Value.ToString());
			Assert.Equal("3", worksheet.Cell(4, 1).Value.ToString());
			Assert.Equal("4", worksheet.Cell(5, 1).Value.ToString());
			Assert.Equal("5", worksheet.Cell(6, 1).Value.ToString());
			Assert.Equal("15", worksheet.Cell(7, 1).Value.ToString());
			Assert.Equal("item1 TOTAL: 15", worksheet.Cell(8, 1).Value.ToString());


			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
		}
	}

	[Fact]
	public void Function_MultiRange()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateFunction2.xlsx";
			var template = new WvSpreadsheetFileTemplate
			{
				Template = new TestUtils().LoadFileAsStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvSpreadsheetFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			new TestUtils().GeneralResultChecks(result);
			Assert.Single(result!.Workbook!.Worksheets);
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Workbook);
			Assert.Single(result!.ResultItems[0]!.Workbook!.Worksheets);
			var worksheet = result!.ResultItems[0]!.Workbook!.Worksheets.First();
			Assert.Equal("position", worksheet.Cell(1, 1).Value.ToString());
			Assert.Equal("1", worksheet.Cell(2, 1).Value.ToString());
			Assert.Equal("2", worksheet.Cell(3, 1).Value.ToString());
			Assert.Equal("3", worksheet.Cell(4, 1).Value.ToString());
			Assert.Equal("4", worksheet.Cell(5, 1).Value.ToString());
			Assert.Equal("5", worksheet.Cell(6, 1).Value.ToString());
			Assert.Equal("15", worksheet.Cell(7, 1).Value.ToString());
			Assert.Equal("15", worksheet.Cell(7, 2).Value.ToString());


			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
		}
	}

	[Fact]
	public void Function_Dependancies()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateFunction3.xlsx";
			var template = new WvSpreadsheetFileTemplate
			{
				Template = new TestUtils().LoadFileAsStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvSpreadsheetFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			new TestUtils().GeneralResultChecks(result);
			var positionContext1 = result!.TemplateContexts
				.FirstOrDefault(x => x.Range!.RangeAddress.FirstAddress.RowNumber == 2
					&& x.Range!.RangeAddress.FirstAddress.ColumnNumber == 1);
			var positionContext2 = result!.TemplateContexts
				.FirstOrDefault(x => x.Range!.RangeAddress.FirstAddress.RowNumber == 2
					&& x.Range!.RangeAddress.FirstAddress.ColumnNumber == 2);

			var functionContext1 = result!.TemplateContexts
				.FirstOrDefault(x => x.Range!.RangeAddress.FirstAddress.RowNumber == 3
					&& x.Range!.RangeAddress.FirstAddress.ColumnNumber == 1);
			var functionContext2 = result!.TemplateContexts
				.FirstOrDefault(x => x.Range!.RangeAddress.FirstAddress.RowNumber == 3
					&& x.Range!.RangeAddress.FirstAddress.ColumnNumber == 2);
			var functionContext3 = result!.TemplateContexts
				.FirstOrDefault(x => x.Range!.RangeAddress.FirstAddress.RowNumber == 3
					&& x.Range!.RangeAddress.FirstAddress.ColumnNumber == 3);

			Assert.NotNull(functionContext1);
			Assert.NotNull(functionContext2);
			Assert.NotNull(functionContext3);

			Assert.Single(functionContext1.ContextDependencies);
			Assert.Single(functionContext2.ContextDependencies);
			Assert.Equal(2, functionContext3.ContextDependencies.Count);

			Assert.Contains(positionContext1!.Id, functionContext1.ContextDependencies);
			Assert.Contains(positionContext2!.Id, functionContext2.ContextDependencies);

			Assert.Contains(functionContext1!.Id, functionContext3.ContextDependencies);
			Assert.Contains(functionContext2!.Id, functionContext3.ContextDependencies);

			Assert.Single(result!.Workbook!.Worksheets);
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Workbook);
			Assert.Single(result!.ResultItems[0]!.Workbook!.Worksheets);
			var worksheet = result!.ResultItems[0]!.Workbook!.Worksheets.First();
			Assert.Equal("position", worksheet.Cell(1, 1).Value.ToString());
			Assert.Equal("1", worksheet.Cell(2, 1).Value.ToString());
			Assert.Equal("2", worksheet.Cell(3, 1).Value.ToString());
			Assert.Equal("3", worksheet.Cell(4, 1).Value.ToString());
			Assert.Equal("4", worksheet.Cell(5, 1).Value.ToString());
			Assert.Equal("5", worksheet.Cell(6, 1).Value.ToString());
			Assert.Equal("15", worksheet.Cell(7, 1).Value.ToString());
			Assert.Equal("15", worksheet.Cell(7, 2).Value.ToString());
			Assert.Equal("30", worksheet.Cell(7, 3).Value.ToString());


			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
		}
	}
	[Fact]
	public void Function_Dependancies_Range()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateFunction4.xlsx";
			var template = new WvSpreadsheetFileTemplate
			{
				Template = new TestUtils().LoadFileAsStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvSpreadsheetFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			new TestUtils().GeneralResultChecks(result);
			var positionContext1 = result!.TemplateContexts
				.FirstOrDefault(x => x.Range!.RangeAddress.FirstAddress.RowNumber == 2
					&& x.Range!.RangeAddress.FirstAddress.ColumnNumber == 1);
			var positionContext2 = result!.TemplateContexts
				.FirstOrDefault(x => x.Range!.RangeAddress.FirstAddress.RowNumber == 2
					&& x.Range!.RangeAddress.FirstAddress.ColumnNumber == 2);

			var functionContext1 = result!.TemplateContexts
				.FirstOrDefault(x => x.Range!.RangeAddress.FirstAddress.RowNumber == 3
					&& x.Range!.RangeAddress.FirstAddress.ColumnNumber == 1);
			var functionContext2 = result!.TemplateContexts
				.FirstOrDefault(x => x.Range!.RangeAddress.FirstAddress.RowNumber == 3
					&& x.Range!.RangeAddress.FirstAddress.ColumnNumber == 2);
			var functionContext3 = result!.TemplateContexts
				.FirstOrDefault(x => x.Range!.RangeAddress.FirstAddress.RowNumber == 3
					&& x.Range!.RangeAddress.FirstAddress.ColumnNumber == 3);

			Assert.NotNull(functionContext1);
			Assert.NotNull(functionContext2);
			Assert.NotNull(functionContext3);

			Assert.Single(functionContext1.ContextDependencies);
			Assert.Single(functionContext2.ContextDependencies);
			Assert.Equal(2, functionContext3.ContextDependencies.Count);

			Assert.Contains(positionContext1!.Id, functionContext1.ContextDependencies);
			Assert.Contains(positionContext2!.Id, functionContext2.ContextDependencies);

			Assert.Contains(functionContext1!.Id, functionContext3.ContextDependencies);
			Assert.Contains(functionContext2!.Id, functionContext3.ContextDependencies);

			Assert.Single(result!.Workbook!.Worksheets);
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Workbook);
			Assert.Single(result!.ResultItems[0]!.Workbook!.Worksheets);
			var worksheet = result!.ResultItems[0]!.Workbook!.Worksheets.First();
			Assert.Equal("position", worksheet.Cell(1, 1).Value.ToString());
			Assert.Equal("1", worksheet.Cell(2, 1).Value.ToString());
			Assert.Equal("2", worksheet.Cell(3, 1).Value.ToString());
			Assert.Equal("3", worksheet.Cell(4, 1).Value.ToString());
			Assert.Equal("4", worksheet.Cell(5, 1).Value.ToString());
			Assert.Equal("5", worksheet.Cell(6, 1).Value.ToString());
			Assert.Equal("15", worksheet.Cell(7, 1).Value.ToString());
			Assert.Equal("15", worksheet.Cell(7, 2).Value.ToString());
			Assert.Equal("30", worksheet.Cell(7, 3).Value.ToString());


			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
		}
	}

	#endregion

	#region << ABS >>
	[Fact]
	public void Function_ABS_SingleRange()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateFunction-ABS-1.xlsx";
			var template = new WvSpreadsheetFileTemplate
			{
				Template = new TestUtils().LoadFileAsStream(templateFile)
			};
			SampleData.Rows[0]["position"] = -100;
			var dataSource = SampleData;
			//When
			WvSpreadsheetFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			new TestUtils().GeneralResultChecks(result);
			Assert.Single(result!.Workbook!.Worksheets);
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Workbook);
			Assert.Single(result!.ResultItems[0]!.Workbook!.Worksheets);
			var worksheet = result!.ResultItems[0]!.Workbook!.Worksheets.First();
			Assert.Equal("position", worksheet.Cell(1, 1).Value.ToString());
			Assert.Equal(-100, worksheet.Cell(2, 1).Value);
			Assert.Equal("2", worksheet.Cell(3, 1).Value.ToString());
			Assert.Equal("3", worksheet.Cell(4, 1).Value.ToString());
			Assert.Equal("4", worksheet.Cell(5, 1).Value.ToString());
			Assert.Equal("5", worksheet.Cell(6, 1).Value.ToString());
			Assert.Equal("86", worksheet.Cell(7, 1).Value.ToString());
			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
		}
	}
	#endregion

	#region << AVERAGE >>
	[Fact]
	public void Function_AVERAGE_SingleRange()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateFunction-Average-1.xlsx";
			var template = new WvSpreadsheetFileTemplate
			{
				Template = new TestUtils().LoadFileAsStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvSpreadsheetFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			new TestUtils().GeneralResultChecks(result);
			Assert.Single(result!.Workbook!.Worksheets);
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Workbook);
			Assert.Single(result!.ResultItems[0]!.Workbook!.Worksheets);
			var worksheet = result!.ResultItems[0]!.Workbook!.Worksheets.First();
			Assert.Equal("position", worksheet.Cell(1, 1).Value.ToString());
			Assert.Equal("1", worksheet.Cell(2, 1).Value.ToString());
			Assert.Equal("2", worksheet.Cell(3, 1).Value.ToString());
			Assert.Equal("3", worksheet.Cell(4, 1).Value.ToString());
			Assert.Equal("4", worksheet.Cell(5, 1).Value.ToString());
			Assert.Equal("5", worksheet.Cell(6, 1).Value.ToString());
			Assert.Equal("3", worksheet.Cell(7, 1).Value.ToString());
			Assert.Equal("item1 AVERAGE: 3", worksheet.Cell(8, 1).Value.ToString());
			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
		}
	}
	#endregion

	#region << CONCAT >>
	[Fact]
	public void Function_CONCAT_SingleRange()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateFunction-CONCAT-1.xlsx";
			var template = new WvSpreadsheetFileTemplate
			{
				Template = new TestUtils().LoadFileAsStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvSpreadsheetFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			new TestUtils().GeneralResultChecks(result);
			Assert.Single(result!.Workbook!.Worksheets);
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Workbook);
			Assert.Single(result!.ResultItems[0]!.Workbook!.Worksheets);
			var worksheet = result!.ResultItems[0]!.Workbook!.Worksheets.First();
			Assert.Equal("name", worksheet.Cell(1, 1).Value.ToString());
			Assert.Equal("item1", worksheet.Cell(2, 1).Value.ToString());
			Assert.Equal("item2", worksheet.Cell(3, 1).Value.ToString());
			Assert.Equal("item3", worksheet.Cell(4, 1).Value.ToString());
			Assert.Equal("item4", worksheet.Cell(5, 1).Value.ToString());
			Assert.Equal("item5", worksheet.Cell(6, 1).Value.ToString());
			Assert.Equal("item1item2item3item4item5", worksheet.Cell(7, 1).Value.ToString());
			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
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
			var template = new WvSpreadsheetFileTemplate
			{
				Template = new TestUtils().LoadFileAsStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvSpreadsheetFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			new TestUtils().GeneralResultChecks(result);
			Assert.Single(result!.Workbook!.Worksheets);
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Workbook);
			Assert.Single(result!.ResultItems[0]!.Workbook!.Worksheets);
			var worksheet = result!.ResultItems[0]!.Workbook!.Worksheets.First();
			Assert.Equal("position", worksheet.Cell(1, 1).Value.ToString());
			Assert.Equal("1", worksheet.Cell(2, 1).Value.ToString());
			Assert.Equal("2", worksheet.Cell(3, 1).Value.ToString());
			Assert.Equal("3", worksheet.Cell(4, 1).Value.ToString());
			Assert.Equal("4", worksheet.Cell(5, 1).Value.ToString());
			Assert.Equal("5", worksheet.Cell(6, 1).Value.ToString());
			Assert.Equal("5", worksheet.Cell(7, 1).Value.ToString());
			Assert.Equal("item1 MAX: 5", worksheet.Cell(8, 1).Value.ToString());
			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
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
			var template = new WvSpreadsheetFileTemplate
			{
				Template = new TestUtils().LoadFileAsStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvSpreadsheetFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			new TestUtils().GeneralResultChecks(result);
			Assert.Single(result!.Workbook!.Worksheets);
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Workbook);
			Assert.Single(result!.ResultItems[0]!.Workbook!.Worksheets);
			var worksheet = result!.ResultItems[0]!.Workbook!.Worksheets.First();
			Assert.Equal("position", worksheet.Cell(1, 1).Value.ToString());
			Assert.Equal("1", worksheet.Cell(2, 1).Value.ToString());
			Assert.Equal("2", worksheet.Cell(3, 1).Value.ToString());
			Assert.Equal("3", worksheet.Cell(4, 1).Value.ToString());
			Assert.Equal("4", worksheet.Cell(5, 1).Value.ToString());
			Assert.Equal("5", worksheet.Cell(6, 1).Value.ToString());
			Assert.Equal("1", worksheet.Cell(7, 1).Value.ToString());
			Assert.Equal("item1 MIN: 1", worksheet.Cell(8, 1).Value.ToString());
			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
		}
	}
	#endregion

	#region << SUM >>
	[Fact]
	public void Function_SUM_RowRepeat()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateFunction-SUM-1.xlsx";
			var template = new WvSpreadsheetFileTemplate
			{
				Template = new TestUtils().LoadFileAsStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvSpreadsheetFileTemplateProcessResult? result = template.Process(dataSource);
			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);

			//Then
			new TestUtils().GeneralResultChecks(result);
			Assert.Single(result!.Workbook!.Worksheets);
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Workbook);
			Assert.Single(result!.ResultItems[0]!.Workbook!.Worksheets);
			var worksheet = result!.ResultItems[0]!.Workbook!.Worksheets.First();
			Assert.Equal("position", worksheet.Cell(1, 1).Value.ToString());
			Assert.Equal("1", worksheet.Cell(2, 1).Value.ToString());
			Assert.Equal("2", worksheet.Cell(3, 1).Value.ToString());
			Assert.Equal("3", worksheet.Cell(4, 1).Value.ToString());
			Assert.Equal("4", worksheet.Cell(5, 1).Value.ToString());
			Assert.Equal("5", worksheet.Cell(6, 1).Value.ToString());
			Assert.Equal("position", worksheet.Cell(1, 2).Value.ToString());
			Assert.Equal("1", worksheet.Cell(2, 2).Value.ToString());
			Assert.Equal("2", worksheet.Cell(3, 2).Value.ToString());
			Assert.Equal("3", worksheet.Cell(4, 2).Value.ToString());
			Assert.Equal("4", worksheet.Cell(5, 2).Value.ToString());
			Assert.Equal("5", worksheet.Cell(6, 2).Value.ToString());
			Assert.Equal("total", worksheet.Cell(1, 3).Value.ToString());
			Assert.Equal("2", worksheet.Cell(2, 3).Value.ToString());
			Assert.Equal("4", worksheet.Cell(3, 3).Value.ToString());
			Assert.Equal("6", worksheet.Cell(4, 3).Value.ToString());
			Assert.Equal("8", worksheet.Cell(5, 3).Value.ToString());
			Assert.Equal("10", worksheet.Cell(6, 3).Value.ToString());

			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
		}
	}

	[Fact]
	public void Function_SUM_Totals()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateFunction-SUM-2.xlsx";
			var template = new WvSpreadsheetFileTemplate
			{
				Template = new TestUtils().LoadFileAsStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvSpreadsheetFileTemplateProcessResult? result = template.Process(dataSource);
			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);

			//Then
			new TestUtils().GeneralResultChecks(result);
			Assert.Single(result!.Workbook!.Worksheets);
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Workbook);
			Assert.Single(result!.ResultItems[0]!.Workbook!.Worksheets);
			var worksheet = result!.ResultItems[0]!.Workbook!.Worksheets.First();
			Assert.Equal("position", worksheet.Cell(1, 1).Value);
			Assert.Equal(1, worksheet.Cell(2, 1).Value);
			Assert.Equal(2, worksheet.Cell(3, 1).Value);
			Assert.Equal(3, worksheet.Cell(4, 1).Value);
			Assert.Equal(4, worksheet.Cell(5, 1).Value);
			Assert.Equal(5, worksheet.Cell(6, 1).Value);
			Assert.Equal("position", worksheet.Cell(1, 2).Value);
			Assert.Equal(1, worksheet.Cell(2, 2).Value);
			Assert.Equal(2, worksheet.Cell(3, 2).Value);
			Assert.Equal(3, worksheet.Cell(4, 2).Value);
			Assert.Equal(4, worksheet.Cell(5, 2).Value);
			Assert.Equal(5, worksheet.Cell(6, 2).Value);

			Assert.Equal(15, worksheet.Cell(7, 1).Value);
			Assert.Equal(15, worksheet.Cell(7, 2).Value);

			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
		}
	}

	[Fact]
	public void Function_SUM_Static()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateFunction-SUM-3.xlsx";
			var template = new WvSpreadsheetFileTemplate
			{
				Template = new TestUtils().LoadFileAsStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvSpreadsheetFileTemplateProcessResult? result = template.Process(dataSource);
			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);

			//Then
			new TestUtils().GeneralResultChecks(result);
			Assert.Single(result!.Workbook!.Worksheets);
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Workbook);
			Assert.Single(result!.ResultItems[0]!.Workbook!.Worksheets);
			var worksheet = result!.ResultItems[0]!.Workbook!.Worksheets.First();
			Assert.Equal("position", worksheet.Cell(1, 1).Value);
			Assert.Equal(1, worksheet.Cell(2, 1).Value);
			Assert.Equal(2, worksheet.Cell(3, 1).Value);
			Assert.Equal(3, worksheet.Cell(4, 1).Value);
			Assert.Equal(4, worksheet.Cell(5, 1).Value);
			Assert.Equal(5, worksheet.Cell(6, 1).Value);
			Assert.Equal("position", worksheet.Cell(1, 2).Value.ToString());
			Assert.Equal(1, worksheet.Cell(2, 2).Value);
			Assert.Equal(2, worksheet.Cell(3, 2).Value);
			Assert.Equal(3, worksheet.Cell(4, 2).Value);
			Assert.Equal(4, worksheet.Cell(5, 2).Value);
			Assert.Equal(5, worksheet.Cell(6, 2).Value);
			Assert.Equal("total", worksheet.Cell(1, 3).Value);
			Assert.Equal(2, worksheet.Cell(2, 3).Value);
			Assert.Equal(3, worksheet.Cell(3, 3).Value);
			Assert.Equal(4, worksheet.Cell(4, 3).Value);
			Assert.Equal(5, worksheet.Cell(5, 3).Value);
			Assert.Equal(6, worksheet.Cell(6, 3).Value);

			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
		}
	}
	#endregion
}

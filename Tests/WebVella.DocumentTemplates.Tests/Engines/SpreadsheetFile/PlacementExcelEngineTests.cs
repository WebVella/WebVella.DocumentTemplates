using WebVella.DocumentTemplates.Engines.SpreadsheetFile;
using WebVella.DocumentTemplates.Tests.Models;
using WebVella.DocumentTemplates.Tests.Utils;

namespace WebVella.DocumentTemplates.Tests.Engines;
public partial class PlacementSpreadsheetEngineTests : TestBase
{
	private static readonly object locker = new();
	public PlacementSpreadsheetEngineTests() : base() { }
	[Fact]
	public void Placement1_Styles()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplatePlacement1.xlsx";
			var template = new WvSpreadsheetFileTemplate
			{
				Template = new TestUtils().LoadWorkbookAsMemoryStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvSpreadsheetFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			new TestUtils().GeneralResultChecks(result);
			Assert.Single(result!.Workbook!.Worksheets);
			Assert.Equal(2, result.TemplateContexts.Count);
			new TestUtils().CheckRangeDimensions(result!.TemplateContexts[0]!.Range!, 1, 1, 1, 1);
			new TestUtils().CheckRangeDimensions(result!.TemplateContexts[1]!.Range!, 1, 2, 1, 3);

			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Workbook);
			Assert.NotNull(result!.ResultItems[0]!.Workbook!.Worksheets);
			Assert.Single(result!.ResultItems[0]!.Workbook!.Worksheets);
			Assert.Equal(2, result!.ResultItems[0]!.Contexts.Count);
			Assert.Null(result!.ResultItems[0]!.Contexts[0].SpreadsheetCellError);
			Assert.Null(result!.ResultItems[0]!.Contexts[1].SpreadsheetCellError);

			new TestUtils().CheckRangeDimensions(result!.ResultItems[0]!.Contexts[0].Range!, 1, 1, 1, 1);
			new TestUtils().CheckRangeDimensions(result!.ResultItems[0]!.Contexts[1].Range!, 1, 2, 1, 3);
			new TestUtils().CheckCellPropertiesCopy(result.TemplateContexts, result!.ResultItems[0]!);
			var tempWs = result.Workbook.Worksheets.First();
			var resultWs = result!.ResultItems[0]!.Workbook!.Worksheets.First();
			new TestUtils().CompareRowProperties(tempWs.Row(1), resultWs.Row(1));
			new TestUtils().CompareColumnProperties(tempWs.Column(1), resultWs.Column(1));
			new TestUtils().CompareColumnProperties(tempWs.Column(2), resultWs.Column(2));

			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
		}
	}

	[Fact]
	public void Placement2_MultiWorksheet()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplatePlacement2.xlsx";
			var template = new WvSpreadsheetFileTemplate
			{
				Template = new TestUtils().LoadWorkbookAsMemoryStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvSpreadsheetFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			new TestUtils().GeneralResultChecks(result);
			Assert.Equal(2, result!.Workbook!.Worksheets.Count);
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Workbook);
			Assert.Equal(2, result!.ResultItems[0]!.Workbook!.Worksheets.Count);
			Assert.Equal(3, result!.ResultItems[0]!.Contexts.Count);
			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
		}
	}

	[Fact]
	public void Placement3_Data()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplatePlacement3.xlsx";
			var template = new WvSpreadsheetFileTemplate
			{
				Template = new TestUtils().LoadWorkbookAsMemoryStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvSpreadsheetFileTemplateProcessResult? result = template.Process(dataSource);
			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
			//Then
			new TestUtils().GeneralResultChecks(result);
			Assert.Single(result!.Workbook!.Worksheets);
			Assert.Equal(6, result.TemplateContexts.Count);
			new TestUtils().CheckRangeDimensions(result!.TemplateContexts[0]!.Range!, 1, 1, 1, 1);
			new TestUtils().CheckRangeDimensions(result!.TemplateContexts[1]!.Range!, 1, 2, 1, 2);
			new TestUtils().CheckRangeDimensions(result!.TemplateContexts[2]!.Range!, 1, 3, 1, 3);
			new TestUtils().CheckRangeDimensions(result!.TemplateContexts[3]!.Range!, 1, 4, 1, 4);
			new TestUtils().CheckRangeDimensions(result!.TemplateContexts[4]!.Range!, 1, 5, 1, 5);
			new TestUtils().CheckRangeDimensions(result!.TemplateContexts[5]!.Range!, 1, 6, 1, 6);

			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Workbook);
			Assert.Single(result!.ResultItems[0]!.Workbook!.Worksheets);
			Assert.Equal(6, result!.ResultItems[0]!.Contexts.Count);
			Assert.Null(result!.ResultItems[0]!.Contexts[0].SpreadsheetCellError);
			Assert.Null(result!.ResultItems[0]!.Contexts[1].SpreadsheetCellError);
			Assert.Null(result!.ResultItems[0]!.Contexts[2].SpreadsheetCellError);
			Assert.Null(result!.ResultItems[0]!.Contexts[3].SpreadsheetCellError);
			Assert.Null(result!.ResultItems[0]!.Contexts[4].SpreadsheetCellError);
			Assert.Null(result!.ResultItems[0]!.Contexts[5].SpreadsheetCellError);


			new TestUtils().CheckRangeDimensions(result!.ResultItems[0]!.Contexts[0].Range!, 1, 1, 5, 1);
			new TestUtils().CheckRangeDimensions(result!.ResultItems[0]!.Contexts[1].Range!, 1, 2, 5, 2);
			new TestUtils().CheckRangeDimensions(result!.ResultItems[0]!.Contexts[2].Range!, 1, 3, 1, 3);
			new TestUtils().CheckRangeDimensions(result!.ResultItems[0]!.Contexts[3].Range!, 1, 4, 5, 4);
			new TestUtils().CheckRangeDimensions(result!.ResultItems[0]!.Contexts[4].Range!, 1, 5, 1, 5);
			new TestUtils().CheckRangeDimensions(result!.ResultItems[0]!.Contexts[5].Range!, 1, 6, 1, 6);


			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
		}
	}

	[Fact]
	public void Placement4_WrongColumnName()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplatePlacement4.xlsx";
			var template = new WvSpreadsheetFileTemplate
			{
				Template = new TestUtils().LoadWorkbookAsMemoryStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvSpreadsheetFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			new TestUtils().GeneralResultChecks(result);
			Assert.Single(result!.Workbook!.Worksheets);
			Assert.Single(result.TemplateContexts);
			new TestUtils().CheckRangeDimensions(result!.TemplateContexts[0]!.Range!, 1, 1, 1, 1);

			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Workbook);
			Assert.Single(result!.ResultItems[0]!.Workbook!.Worksheets);
			Assert.Single(result!.ResultItems[0]!.Contexts);
			Assert.Null(result!.ResultItems[0]!.Contexts[0].SpreadsheetCellError);
			new TestUtils().CheckRangeDimensions(result!.ResultItems[0]!.Contexts[0].Range!, 1, 1, 5, 1);

			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
		}
	}

	[Fact]
	public void Placement5_EmptyComlumnsWithStyles()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplatePlacement5.xlsx";
			var template = new WvSpreadsheetFileTemplate
			{
				Template = new TestUtils().LoadWorkbookAsMemoryStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvSpreadsheetFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			new TestUtils().GeneralResultChecks(result);
			Assert.Single(result!.Workbook!.Worksheets);
			Assert.Equal(4, result.TemplateContexts.Count);
			new TestUtils().CheckRangeDimensions(result!.TemplateContexts[0]!.Range!, 1, 1, 1, 1);
			new TestUtils().CheckRangeDimensions(result!.TemplateContexts[1]!.Range!, 1, 2, 1, 2);
			new TestUtils().CheckRangeDimensions(result!.TemplateContexts[2]!.Range!, 2, 1, 2, 1);
			new TestUtils().CheckRangeDimensions(result!.TemplateContexts[3]!.Range!, 2, 2, 2, 2);

			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Workbook);
			Assert.Single(result!.ResultItems[0]!.Workbook!.Worksheets);
			Assert.Equal(4, result!.ResultItems[0]!.Contexts.Count);
			Assert.Null(result!.ResultItems[0]!.Contexts[0].SpreadsheetCellError);
			Assert.Null(result!.ResultItems[0]!.Contexts[1].SpreadsheetCellError);
			Assert.Null(result!.ResultItems[0]!.Contexts[2].SpreadsheetCellError);
			Assert.Null(result!.ResultItems[0]!.Contexts[3].SpreadsheetCellError);

			new TestUtils().CheckRangeDimensions(result!.ResultItems[0]!.Contexts[0].Range!, 1, 1, 5, 1);
			new TestUtils().CheckRangeDimensions(result!.ResultItems[0]!.Contexts[1].Range!, 1, 2, 1, 2);
			new TestUtils().CheckRangeDimensions(result!.ResultItems[0]!.Contexts[2].Range!, 6, 1, 10, 1);
			new TestUtils().CheckRangeDimensions(result!.ResultItems[0]!.Contexts[3].Range!, 6, 2, 6, 2);

			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
		}
	}

	[Fact]
	public void Placement6_Data_HorizontalFlow()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplatePlacement6.xlsx";
			var template = new WvSpreadsheetFileTemplate
			{
				Template = new TestUtils().LoadWorkbookAsMemoryStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvSpreadsheetFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			new TestUtils().GeneralResultChecks(result);
			Assert.Single(result!.Workbook!.Worksheets);
			Assert.Equal(6, result.TemplateContexts.Count);
			new TestUtils().CheckRangeDimensions(result!.TemplateContexts[0]!.Range!, 1, 1, 1, 1);
			new TestUtils().CheckRangeDimensions(result!.TemplateContexts[1]!.Range!, 1, 2, 1, 2);
			new TestUtils().CheckRangeDimensions(result!.TemplateContexts[2]!.Range!, 1, 3, 1, 3);
			new TestUtils().CheckRangeDimensions(result!.TemplateContexts[3]!.Range!, 2, 1, 2, 1);
			new TestUtils().CheckRangeDimensions(result!.TemplateContexts[4]!.Range!, 2, 2, 2, 2);
			new TestUtils().CheckRangeDimensions(result!.TemplateContexts[5]!.Range!, 2, 3, 2, 3);


			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Workbook);
			Assert.Single(result!.ResultItems[0]!.Workbook!.Worksheets);
			Assert.Equal(6, result!.ResultItems[0]!.Contexts.Count);
			Assert.Null(result!.ResultItems[0]!.Contexts[0].SpreadsheetCellError);
			Assert.Null(result!.ResultItems[0]!.Contexts[1].SpreadsheetCellError);
			Assert.Null(result!.ResultItems[0]!.Contexts[2].SpreadsheetCellError);
			Assert.Null(result!.ResultItems[0]!.Contexts[3].SpreadsheetCellError);
			Assert.Null(result!.ResultItems[0]!.Contexts[4].SpreadsheetCellError);
			Assert.Null(result!.ResultItems[0]!.Contexts[5].SpreadsheetCellError);

			new TestUtils().CheckRangeDimensions(result!.ResultItems[0]!.Contexts[0].Range!, 1, 1, 1, 5);
			new TestUtils().CheckRangeDimensions(result!.ResultItems[0]!.Contexts[1].Range!, 1, 6, 1, 6);
			new TestUtils().CheckRangeDimensions(result!.ResultItems[0]!.Contexts[2].Range!, 1, 7, 5, 7);
			new TestUtils().CheckRangeDimensions(result!.ResultItems[0]!.Contexts[3].Range!, 6, 1, 6, 5);
			new TestUtils().CheckRangeDimensions(result!.ResultItems[0]!.Contexts[4].Range!, 6, 6, 6, 6);
			new TestUtils().CheckRangeDimensions(result!.ResultItems[0]!.Contexts[5].Range!, 6, 7, 6, 7);
			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);

		}
	}

	[Fact]
	public void Placement7_Merge()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplatePlacement7.xlsx";
			var template = new WvSpreadsheetFileTemplate
			{
				Template = new TestUtils().LoadWorkbookAsMemoryStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvSpreadsheetFileTemplateProcessResult? result = template.Process(dataSource);
			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
			//Then
			new TestUtils().GeneralResultChecks(result);

			//Check template
			Assert.NotNull(result.Workbook);
			Assert.NotNull(result.TemplateContexts);
			Assert.Equal(3, result.TemplateContexts.Count);

			new TestUtils().CheckRangeDimensions(result!.TemplateContexts[0].Range!, 1, 1, 1, 2);
			new TestUtils().CheckRangeDimensions(result!.TemplateContexts[1].Range!, 1, 3, 1, 3);
			new TestUtils().CheckRangeDimensions(result!.TemplateContexts[2].Range!, 1, 4, 1, 4);


			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			var resultItem = result!.ResultItems[0];

			Assert.NotNull(resultItem.Workbook);
			Assert.Single(resultItem.Workbook!.Worksheets);
			Assert.Equal(3, resultItem.Contexts.Count);

			new TestUtils().CheckRangeDimensions(resultItem.Contexts[0].Range!, 1, 1, 5, 2);
			new TestUtils().CheckRangeDimensions(resultItem.Contexts[1].Range!, 1, 3, 5, 3);
			new TestUtils().CheckRangeDimensions(resultItem.Contexts[2].Range!, 1, 4, 1, 4);

			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);

		}
	}

}

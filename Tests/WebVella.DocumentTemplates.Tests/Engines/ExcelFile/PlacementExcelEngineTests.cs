using WebVella.DocumentTemplates.Engines.ExcelFile;
using WebVella.DocumentTemplates.Tests.Models;

namespace WebVella.DocumentTemplates.Tests.Engines;
public partial class PlacementExcelEngineTests : TestBase
{
	private static readonly object locker = new();
	public PlacementExcelEngineTests() : base() { }
	[Fact]
	public void Placement1_Styles()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplatePlacement1.xlsx";
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
			Assert.Equal(2, result.TemplateContexts.Count);
			CheckRangeDimensions(result!.TemplateContexts[0]!.Range!, 1, 1, 1, 1);
			CheckRangeDimensions(result!.TemplateContexts[1]!.Range!, 1, 2, 1, 3);

			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Result);
			Assert.NotNull(result!.ResultItems[0]!.Result!.Worksheets);
			Assert.Single(result!.ResultItems[0]!.Result!.Worksheets);
			Assert.Equal(2, result!.ResultItems[0]!.Contexts.Count);
			Assert.Null(result!.ResultItems[0]!.Contexts[0].ExcelCellError);
			Assert.Null(result!.ResultItems[0]!.Contexts[1].ExcelCellError);

			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[0].Range!, 1, 1, 1, 1);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[1].Range!, 1, 2, 1, 3);
			CheckCellPropertiesCopy(result.TemplateContexts, result!.ResultItems[0]!);
			var tempWs = result.Template.Worksheets.First();
			var resultWs = result!.ResultItems[0]!.Result!.Worksheets.First();
			CompareRowProperties(tempWs.Row(1), resultWs.Row(1));
			CompareColumnProperties(tempWs.Column(1), resultWs.Column(1));
			CompareColumnProperties(tempWs.Column(2), resultWs.Column(2));

			SaveWorkbook(result!.ResultItems[0]!.Result!, templateFile);
		}
	}

	[Fact]
	public void Placement2_MultiWorksheet()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplatePlacement2.xlsx";
			var template = new WvExcelFileTemplate
			{
				Template = LoadWorkbook(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvExcelFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			GeneralResultChecks(result);
			Assert.Equal(2, result!.Template!.Worksheets.Count);
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Result);
			Assert.Equal(2, result!.ResultItems[0]!.Result!.Worksheets.Count);
			Assert.Equal(3, result!.ResultItems[0]!.Contexts.Count);
			SaveWorkbook(result!.ResultItems[0]!.Result!, templateFile);
		}
	}

	[Fact]
	public void Placement3_Data()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplatePlacement3.xlsx";
			var template = new WvExcelFileTemplate
			{
				Template = LoadWorkbook(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvExcelFileTemplateProcessResult? result = template.Process(dataSource);
			SaveWorkbook(result!.ResultItems[0]!.Result!, templateFile);
			//Then
			GeneralResultChecks(result);
			Assert.Single(result!.Template!.Worksheets);
			Assert.Equal(6, result.TemplateContexts.Count);
			CheckRangeDimensions(result!.TemplateContexts[0]!.Range!, 1, 1, 1, 1);
			CheckRangeDimensions(result!.TemplateContexts[1]!.Range!, 1, 2, 1, 2);
			CheckRangeDimensions(result!.TemplateContexts[2]!.Range!, 1, 3, 1, 3);
			CheckRangeDimensions(result!.TemplateContexts[3]!.Range!, 1, 4, 1, 4);
			CheckRangeDimensions(result!.TemplateContexts[4]!.Range!, 1, 5, 1, 5);
			CheckRangeDimensions(result!.TemplateContexts[5]!.Range!, 1, 6, 1, 6);

			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Result);
			Assert.Single(result!.ResultItems[0]!.Result!.Worksheets);
			Assert.Equal(6, result!.ResultItems[0]!.Contexts.Count);
			Assert.Null(result!.ResultItems[0]!.Contexts[0].ExcelCellError);
			Assert.Null(result!.ResultItems[0]!.Contexts[1].ExcelCellError);
			Assert.Null(result!.ResultItems[0]!.Contexts[2].ExcelCellError);
			Assert.Null(result!.ResultItems[0]!.Contexts[3].ExcelCellError);
			Assert.Null(result!.ResultItems[0]!.Contexts[4].ExcelCellError);
			Assert.Null(result!.ResultItems[0]!.Contexts[5].ExcelCellError);


			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[0].Range!, 1, 1, 5, 1);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[1].Range!, 1, 2, 5, 2);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[2].Range!, 1, 3, 1, 3);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[3].Range!, 1, 4, 5, 4);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[4].Range!, 1, 5, 1, 5);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[5].Range!, 1, 6, 1, 6);


			SaveWorkbook(result!.ResultItems[0]!.Result!, templateFile);
		}
	}

	[Fact]
	public void Placement4_WrongColumnName()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplatePlacement4.xlsx";
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
			Assert.Single(result.TemplateContexts);
			CheckRangeDimensions(result!.TemplateContexts[0]!.Range!, 1, 1, 1, 1);

			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Result);
			Assert.Single(result!.ResultItems[0]!.Result!.Worksheets);
			Assert.Single(result!.ResultItems[0]!.Contexts);
			Assert.Null(result!.ResultItems[0]!.Contexts[0].ExcelCellError);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[0].Range!, 1, 1, 5, 1);

			SaveWorkbook(result!.ResultItems[0]!.Result!, templateFile);
		}
	}

	[Fact]
	public void Placement5_EmptyComlumnsWithStyles()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplatePlacement5.xlsx";
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
			Assert.Equal(4, result.TemplateContexts.Count);
			CheckRangeDimensions(result!.TemplateContexts[0]!.Range!, 1, 1, 1, 1);
			CheckRangeDimensions(result!.TemplateContexts[1]!.Range!, 1, 2, 1, 2);
			CheckRangeDimensions(result!.TemplateContexts[2]!.Range!, 2, 1, 2, 1);
			CheckRangeDimensions(result!.TemplateContexts[3]!.Range!, 2, 2, 2, 2);

			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Result);
			Assert.Single(result!.ResultItems[0]!.Result!.Worksheets);
			Assert.Equal(4, result!.ResultItems[0]!.Contexts.Count);
			Assert.Null(result!.ResultItems[0]!.Contexts[0].ExcelCellError);
			Assert.Null(result!.ResultItems[0]!.Contexts[1].ExcelCellError);
			Assert.Null(result!.ResultItems[0]!.Contexts[2].ExcelCellError);
			Assert.Null(result!.ResultItems[0]!.Contexts[3].ExcelCellError);

			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[0].Range!, 1, 1, 5, 1);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[1].Range!, 1, 2, 1, 2);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[2].Range!, 6, 1, 10, 1);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[3].Range!, 6, 2, 6, 2);

			SaveWorkbook(result!.ResultItems[0]!.Result!, templateFile);
		}
	}

	[Fact]
	public void Placement6_Data_HorizontalFlow()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplatePlacement6.xlsx";
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
			Assert.Equal(6, result.TemplateContexts.Count);
			CheckRangeDimensions(result!.TemplateContexts[0]!.Range!, 1, 1, 1, 1);
			CheckRangeDimensions(result!.TemplateContexts[1]!.Range!, 1, 2, 1, 2);
			CheckRangeDimensions(result!.TemplateContexts[2]!.Range!, 1, 3, 1, 3);
			CheckRangeDimensions(result!.TemplateContexts[3]!.Range!, 2, 1, 2, 1);
			CheckRangeDimensions(result!.TemplateContexts[4]!.Range!, 2, 2, 2, 2);
			CheckRangeDimensions(result!.TemplateContexts[5]!.Range!, 2, 3, 2, 3);


			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Result);
			Assert.Single(result!.ResultItems[0]!.Result!.Worksheets);
			Assert.Equal(6, result!.ResultItems[0]!.Contexts.Count);
			Assert.Null(result!.ResultItems[0]!.Contexts[0].ExcelCellError);
			Assert.Null(result!.ResultItems[0]!.Contexts[1].ExcelCellError);
			Assert.Null(result!.ResultItems[0]!.Contexts[2].ExcelCellError);
			Assert.Null(result!.ResultItems[0]!.Contexts[3].ExcelCellError);
			Assert.Null(result!.ResultItems[0]!.Contexts[4].ExcelCellError);
			Assert.Null(result!.ResultItems[0]!.Contexts[5].ExcelCellError);

			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[0].Range!, 1, 1, 1, 5);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[1].Range!, 1, 6, 1, 6);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[2].Range!, 1, 7, 5, 7);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[3].Range!, 6, 1, 6, 5);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[4].Range!, 6, 6, 6, 6);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[5].Range!, 6, 7, 6, 7);
			SaveWorkbook(result!.ResultItems[0]!.Result!, templateFile);

		}
	}

	[Fact]
	public void Placement7_Merge()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplatePlacement7.xlsx";
			var template = new WvExcelFileTemplate
			{
				Template = LoadWorkbook(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvExcelFileTemplateProcessResult? result = template.Process(dataSource);
			SaveWorkbook(result!.ResultItems[0]!.Result!, templateFile);
			//Then
			GeneralResultChecks(result);

			//Check template
			Assert.NotNull(result.Template);
			Assert.NotNull(result.TemplateContexts);
			Assert.Equal(3, result.TemplateContexts.Count);

			CheckRangeDimensions(result!.TemplateContexts[0].Range!, 1, 1, 1, 2);
			CheckRangeDimensions(result!.TemplateContexts[1].Range!, 1, 3, 1, 3);
			CheckRangeDimensions(result!.TemplateContexts[2].Range!, 1, 4, 1, 4);


			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			var resultItem = result!.ResultItems[0];

			Assert.NotNull(resultItem.Result);
			Assert.Single(resultItem.Result!.Worksheets);
			Assert.Equal(3, resultItem.Contexts.Count);

			CheckRangeDimensions(resultItem.Contexts[0].Range!, 1, 1, 5, 2);
			CheckRangeDimensions(resultItem.Contexts[1].Range!, 1, 3, 5, 3);
			CheckRangeDimensions(resultItem.Contexts[2].Range!, 1, 4, 1, 4);

			SaveWorkbook(result!.ResultItems[0]!.Result!, templateFile);

		}
	}

}

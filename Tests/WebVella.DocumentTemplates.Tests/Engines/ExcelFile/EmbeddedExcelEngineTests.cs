using WebVella.DocumentTemplates.Engines.ExcelFile;
using WebVella.DocumentTemplates.Tests.Models;

namespace WebVella.DocumentTemplates.Tests.Engines;
public partial class EmbeddedExcelEngineTests : TestBase
{
	private static readonly object locker = new();
	public EmbeddedExcelEngineTests() : base() { }

	[Fact]
	public void Image_1()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateData5.xlsx";
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

			SaveWorkbook(result!.ResultItems[0]!.Result!, templateFile);
		}
	}
	[Fact]
	public void Image_2()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateData6.xlsx";
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

			SaveWorkbook(result!.ResultItems[0]!.Result!, templateFile);
		}
	}

}

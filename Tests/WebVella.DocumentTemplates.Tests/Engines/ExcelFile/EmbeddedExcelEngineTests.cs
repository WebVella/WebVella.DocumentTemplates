using WebVella.DocumentTemplates.Engines.ExcelFile;
using WebVella.DocumentTemplates.Tests.Models;
using WebVella.DocumentTemplates.Tests.Utils;

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
				Template = new TestUtils().LoadWorkbookAsMemoryStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvExcelFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			new TestUtils().GeneralResultChecks(result);
			Assert.Single(result!.Workbook!.Worksheets);
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Workbook);
			Assert.Single(result!.ResultItems[0]!.Workbook!.Worksheets);

			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
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
				Template = new TestUtils().LoadWorkbookAsMemoryStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvExcelFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			new TestUtils().GeneralResultChecks(result);
			Assert.Single(result!.Workbook!.Worksheets);
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Workbook);
			Assert.Single(result!.ResultItems[0]!.Workbook!.Worksheets);

			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
		}
	}

}

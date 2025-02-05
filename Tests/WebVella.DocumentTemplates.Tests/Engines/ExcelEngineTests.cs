using WebVella.DocumentTemplates.Engines.Excel;
using WebVella.DocumentTemplates.Tests.Models;

namespace WebVella.DocumentTemplates.Tests.Engines;
public class ExcelEngineTests : TestBase
{
	private static readonly object locker = new();
	public ExcelEngineTests() : base() { }

	#region << Arguments >>
	//[Fact]
	//public void ShouldHaveDataSource()
	//{
	//	var templateFile = "TemplatePlacement1.xlsx";
	//	lock (locker)
	//	{
	//		var template = new WvExcelFileTemplate
	//		{
	//			Template = LoadFile(templateFile)
	//		};
	//		var result = template.Process(SampleData, encoding: Encoding.UTF8);
	//		Assert.NotNull(result);
	//		Assert.NotNull(result.Template);
	//		Assert.NotNull(result.Result);
	//		var resultString = Encoding.UTF8.GetString(result.Result);
	//		Assert.Equal(expectedResultText, resultString);

	//		var result = new TfExcelTemplateProcessResult();
	//		result.TemplateWorkbook = LoadWorkbook("TemplatePlacement1.xlsx");
	//		Func<bool> action = () => { result.ProcessExcelTemplate(null); return true; };
	//		action.Should().Throw<Exception>("No datasource provided!");
	//	}
	//}
	#endregion
}

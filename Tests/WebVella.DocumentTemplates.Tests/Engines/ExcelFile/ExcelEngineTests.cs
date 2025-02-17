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
	private static readonly object locker = new();
	public ExcelEngineTests() : base() { }

	#region << Arguments >>
	[Fact]
	public void ShouldHaveDataSource()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplatePlacement1.xlsx";
			var template = new WvExcelFileTemplate
			{
				Template = LoadWorkbook(templateFile)
			};
			WvExcelFileTemplateProcessResult? result = null;
			//When
			var action = () => result = template.Process(null, DefaultCulture);
			//Then
			var ex = Record.Exception(action);
			Assert.NotNull(ex);
			Assert.IsType<ArgumentException>(ex);
			var argEx = (ArgumentException)ex;
			Assert.Equal("dataSource", argEx.ParamName);
			Assert.StartsWith("No datasource provided!", argEx.Message);
		}
	}
	#endregion

	#region << error >>
	[Fact]
	public void Error_1()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateError1.xlsx";
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
			var errorCell = result!.ResultItems[0]!.Result!.Worksheets.First().Cell(6, 1);
			Assert.NotNull(errorCell);
			Assert.True(errorCell.Value.IsError);
			var error = errorCell.Value.GetError();
			Assert.Equal(XLError.IncompatibleValue, error);
			SaveWorkbook(result!.ResultItems[0]!.Result!, templateFile);
		}
	}


	#endregion

	
}

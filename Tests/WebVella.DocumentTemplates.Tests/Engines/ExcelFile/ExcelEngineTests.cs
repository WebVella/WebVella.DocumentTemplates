using ClosedXML.Excel;
using System.Data;
using WebVella.DocumentTemplates.Engines.ExcelFile;
using WebVella.DocumentTemplates.Tests.Models;
using WebVella.DocumentTemplates.Tests.Utils;

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
				Template = new TestUtils().LoadWorkbookAsMemoryStream(templateFile)
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
				Template = new TestUtils().LoadWorkbookAsMemoryStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvExcelFileTemplateProcessResult? result = template.Process(dataSource);
			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
			//Then
			new TestUtils().GeneralResultChecks(result);
			Assert.Single(result!.Workbook!.Worksheets);
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Workbook);
			Assert.Single(result!.ResultItems[0]!.Workbook!.Worksheets);
			var errorCell = result!.ResultItems[0]!.Workbook!.Worksheets.First().Cell(6, 1);
			Assert.NotNull(errorCell);
			Assert.True(errorCell.Value.IsError);
			var error = errorCell.Value.GetError();
			Assert.Equal(XLError.IncompatibleValue, error);
			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
		}
	}


	#endregion

	#region << Docs >>
	[Fact]
	public void Docs1()
	{
		lock (locker)
		{
			var templateFile = "TemplateDoc1.xlsx";
			var ds = new DataTable();
			ds.Columns.Add("position",typeof(int));
			ds.Columns.Add("sku",typeof(string));
			ds.Columns.Add("item",typeof(string));
			ds.Columns.Add("price",typeof(decimal));

			for (int i = 1; i < 6; i++) {
				var dsrow = ds.NewRow();
				dsrow["position"] = i;
				dsrow["sku"] = $"SKU{i}";
				dsrow["item"] = $"item {i} description text";
				dsrow["price"] = i * (decimal)0.98;
				ds.Rows.Add(dsrow);
			}

			var template = new WvExcelFileTemplate
			{
				Template = new TestUtils().LoadWorkbookAsMemoryStream(templateFile)
			};
			WvExcelFileTemplateProcessResult? result = template.Process(ds);
			
			new TestUtils().SaveWorkbookFromMemoryStream(result.ResultItems[0].Result!,"TemplateDoc1.xlsx");
		}
	}
	#endregion
}

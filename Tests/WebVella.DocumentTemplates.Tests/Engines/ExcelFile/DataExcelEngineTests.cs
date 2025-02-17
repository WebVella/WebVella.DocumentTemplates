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
public partial class DataExcelEngineTests : TestBase
{
	private static readonly object locker = new();
	public DataExcelEngineTests() : base() { }

	#region << Data >>
	[Fact]
	public void ExcelData1_RepeatAndWorksheetName()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateData1.xlsx";
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
			var ws = result!.ResultItems[0]!.Result!.Worksheets.First();
			Assert.Equal((string)SampleData.Rows[0]["name"], ws.Name);

			for (int i = 0; i < SampleData.Rows.Count; i++)
			{
				var cellValueString = ws.Cell(i + 1, 1).Value.ToString();
				var rowValueString = SampleData.Rows[i]["position"]?.ToString();
				Assert.Equal(rowValueString, cellValueString);
			}

			SaveWorkbook(result!.ResultItems[0]!.Result!, templateFile);
		}
	}

	[Fact]
	public void ExcelData1_RepeatAndWorksheetName_Horizontal()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateData2.xlsx";
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
			var ws = result!.ResultItems[0]!.Result!.Worksheets.First();
			Assert.Equal((string)SampleData.Rows[0]["name"], ws.Name);

			for (int i = 0; i < SampleData.Rows.Count; i++)
			{
				var cellValueString = ws.Cell(1, i + 1).Value.ToString();
				var rowValueString = SampleData.Rows[i]["position"]?.ToString();
				Assert.Equal(rowValueString, cellValueString);
			}

			SaveWorkbook(result!.ResultItems[0]!.Result!, templateFile);
		}
	}

	[Fact]
	public void ExcelData1_RepeatAndWorksheetName_Mixed()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateData3.xlsx";
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
			var ws = result!.ResultItems[0]!.Result!.Worksheets.First();
			Assert.Equal((string)SampleData.Rows[0]["name"], ws.Name);

			for (int i = 0; i < SampleData.Rows.Count; i++)
			{
				var cellValueString = ws.Cell(1, i + 1).Value.ToString();
				var rowValueString = SampleData.Rows[i]["position"]?.ToString();
				Assert.Equal(rowValueString, cellValueString);
			}
			for (int i = 0; i < SampleData.Rows.Count; i++)
			{
				var cellValueString = ws.Cell(i + 2, 1).Value.ToString();
				var rowValueString = SampleData.Rows[i]["name"]?.ToString();
				Assert.Equal(rowValueString, cellValueString);
			}

			SaveWorkbook(result!.ResultItems[0]!.Result!, templateFile);
		}
	}

	[Fact]
	public void ExcelData1_RepeatAndWorksheetName_HighLoad()
	{
		lock (locker)
		{
			//Given
			var ds = new DataTable();
			var colCount = 100;
			//var rowCount = 5;
			var rowCount = 10000;
			var wb = new XLWorkbook();
			var ws = wb.Worksheets.Add();
			//Data
			{
				for (int i = 0; i < colCount; i++)
				{
					var colName = $"col{i}";
					ds.Columns.Add(colName, typeof(string));
					ws.Cell(1, i + 1).Value = "{{" + colName + "}}";
				}
				for (int i = 0; i < rowCount; i++)
				{
					var position = i + 1;
					var dsrow = ds.NewRow();
					for (int j = 0; j < colCount; j++)
					{
						dsrow[$"col{j}"] = $"cell-{i}-{j}";
					}
					ds.Rows.Add(dsrow);
				}
			}


			var templateFile = "TemplateDataGenerated.xlsx";
			var result = new WvExcelFileTemplateProcessResult
			{
				Template = wb
			};
			var resultItem = new WvExcelFileTemplateProcessResultItem()
			{
				Result = new XLWorkbook()
			};
			var dataSource = ds;
			var culture = new CultureInfo("en-US");
			var engine = new WvExcelFileTemplate();
			//When
			var sw = new Stopwatch();
			long timeMS = 0;
			sw.Start();
			WvExcelFileEngineUtility.ProcessExcelTemplateInitTemplateContexts(result);
			timeMS += sw.ElapsedMilliseconds;
			sw.Restart();
			WvExcelFileEngineUtility.ProcessExcelTemplateCalculateDependencies(result);
			timeMS += sw.ElapsedMilliseconds;
			sw.Restart();
			var templContextDict = result.TemplateContexts.ToDictionary(x => x.Id);
			WvExcelFileEngineUtility.ProcessExcelTemplateGenerateResultContexts(result, resultItem, ds, culture, templContextDict);
			timeMS += sw.ElapsedMilliseconds;
			sw.Stop();
			//Then
			SaveWorkbook(resultItem.Result!, templateFile);
			Assert.True(10 * 1000 > timeMS);
		}
	}


	[Fact]
	public void Placement2_MultiWorksheet12()
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
			//Then
			GeneralResultChecks(result);
			Assert.Single(result!.Template!.Worksheets);
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Result);
			Assert.Single(result!.ResultItems[0]!.Result!.Worksheets);
			Assert.Equal(6, result!.ResultItems[0]!.ResultContexts.Count);
			SaveWorkbook(result!.ResultItems[0]!.Result!, templateFile);
		}
	}


	#endregion

	#region << Group By >>
	[Fact]
	public void ExcelData1_GroupBy()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateData1.xlsx";
			var template = new WvExcelFileTemplate
			{
				Template = LoadWorkbook(templateFile),
				GroupDataByColumns = new List<string> { "sku" }
			};
			var dataSource = SampleData;
			dataSource.Rows[1]["sku"] = dataSource.Rows[0]["sku"];
			//When
			WvExcelFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			GeneralResultChecks(result);
			Assert.Single(result!.Template!.Worksheets);
			Assert.NotNull(result!.ResultItems);
			Assert.Equal(4, result!.ResultItems.Count);
			Assert.NotNull(result!.ResultItems[0]!.Result);
			Assert.Single(result!.ResultItems[0]!.Result!.Worksheets);
			var ws = result!.ResultItems[0]!.Result!.Worksheets.First();
			Assert.Equal((string)SampleData.Rows[0]["name"], ws.Name);

			var row1SkuValueString = ws.Cell(1, 2).Value.ToString();
			var row2SkuValueString = ws.Cell(2, 2).Value.ToString();
			Assert.Equal(row1SkuValueString, row2SkuValueString);
		}
	}
	#endregion

}

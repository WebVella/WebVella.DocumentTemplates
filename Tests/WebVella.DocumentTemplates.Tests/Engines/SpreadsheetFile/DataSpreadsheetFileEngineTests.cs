using ClosedXML.Excel;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using WebVella.DocumentTemplates.Engines.SpreadsheetFile;
using WebVella.DocumentTemplates.Engines.SpreadsheetFile.Utility;
using WebVella.DocumentTemplates.Tests.Models;
using WebVella.DocumentTemplates.Tests.Utils;

namespace WebVella.DocumentTemplates.Tests.Engines;
public partial class DataSpreadsheetFileEngineTests : TestBase
{
	private static readonly object locker = new();
	public DataSpreadsheetFileEngineTests() : base() { }

	#region << Data >>
	[Fact]
	public void SpreadsheetData1_RepeatAndWorksheetName()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateData1.xlsx";
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
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Workbook);
			Assert.Single(result!.ResultItems[0]!.Workbook!.Worksheets);
			var ws = result!.ResultItems[0]!.Workbook!.Worksheets.First();
			Assert.Equal((string)SampleData.Rows[0]["name"], ws.Name);

			for (int i = 0; i < SampleData.Rows.Count; i++)
			{
				var data = SampleData.Rows[i]["position"];
				var cell = ws.Cell(i + 1, 1).Value;
				Assert.Equal(data.ToString(), cell.ToString());
			}

			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
		}
	}

	[Fact]
	public void SpreadsheetData1_RepeatAndWorksheetName_Horizontal()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateData2.xlsx";
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
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Workbook);
			Assert.Single(result!.ResultItems[0]!.Workbook!.Worksheets);
			var ws = result!.ResultItems[0]!.Workbook!.Worksheets.First();
			Assert.Equal((string)SampleData.Rows[0]["name"], ws.Name);

			for (int i = 0; i < SampleData.Rows.Count; i++)
			{
				var cellValueString = ws.Cell(1, i + 1).Value.ToString();
				var rowValueString = SampleData.Rows[i]["position"]?.ToString();
				Assert.Equal(rowValueString, cellValueString);
			}

			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
		}
	}

	[Fact]
	public void SpreadsheetData1_RepeatAndWorksheetName_Mixed()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateData3.xlsx";
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
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Workbook);
			Assert.Single(result!.ResultItems[0]!.Workbook!.Worksheets);
			var ws = result!.ResultItems[0]!.Workbook!.Worksheets.First();
			Assert.Equal((string)SampleData.Rows[0]["name"], ws.Name);

			for (int i = 0; i < SampleData.Rows.Count; i++)
			{
				var cellValueString = ws.Cell(1, i + 1).Value.ToString();
				var rowValueString = SampleData.Rows[i]["position"]?.ToString();
				Assert.Equal(rowValueString, cellValueString);
			}
			for (int i = 0; i < SampleData.Rows.Count; i++)
			{
				var cellValueString = ws.Cell(2, i + 1).Value.ToString();
				var rowValueString = SampleData.Rows[i]["name"]?.ToString();
				Assert.Equal(rowValueString, cellValueString);
			}

			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
		}
	}

	[Fact]
	public void SpreadsheetData1_RepeatAndWorksheetName_HighLoad()
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
			MemoryStream ms = new();
			wb.SaveAs(ms);
			WvSpreadsheetFileTemplateProcessResult result = new()
			{
				Template = ms
			};
			var resultItem = new WvSpreadsheetFileTemplateProcessResultItem()
			{
				Workbook = new XLWorkbook()
			};
			var dataSource = ds;
			var culture = new CultureInfo("en-US");
			var engine = new WvSpreadsheetFileTemplate();
			//When
			var sw = new Stopwatch();
			long timeMS = 0;
			sw.Start();
			new WvSpreadsheetFileEngineUtility().ProcessSpreadsheetTemplateInitTemplateContexts(result);
			timeMS += sw.ElapsedMilliseconds;
			sw.Restart();
			new WvSpreadsheetFileEngineUtility().ProcessSpreadsheetTemplateCalculateDependencies(result);
			timeMS += sw.ElapsedMilliseconds;
			sw.Restart();
			var templContextDict = result.TemplateContexts.ToDictionary(x => x.Id);
			new WvSpreadsheetFileEngineUtility().ProcessSpreadsheetTemplateGenerateResultContexts(result, resultItem, ds, culture, templContextDict);
			timeMS += sw.ElapsedMilliseconds;
			sw.Stop();
			//Then
			new TestUtils().SaveWorkbook(resultItem.Workbook!, templateFile);
			Assert.True(15 * 1000 > timeMS);
		}
	}


	[Fact]
	public void Placement2_MultiWorksheet12()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplatePlacement3MultiWs.xlsx";
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
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Workbook);
			Assert.Single(result!.ResultItems[0]!.Workbook!.Worksheets);
			Assert.Equal(6, result!.ResultItems[0]!.Contexts.Count);
			var tempWs = result!.Workbook!.Worksheets.First();
			var resultWs = result!.ResultItems[0]!.Workbook!.Worksheets.First();

			var tempA1 = tempWs.Cell(1, 1);
			var resultA1 = tempWs.Cell(1, 1);

			var tempB1 = tempWs.Cell(1, 2);
			var resultB1 = tempWs.Cell(1, 2);

			var tempC1 = tempWs.Cell(1, 3);
			var resultC1 = tempWs.Cell(1, 3);

			var tempD1 = tempWs.Cell(1, 4);
			var resultD1 = tempWs.Cell(1, 4);

			var tempE1 = tempWs.Cell(1, 5);
			var resultE1 = tempWs.Cell(1, 5);

			var tempF1 = tempWs.Cell(1, 6);
			var resultF1 = tempWs.Cell(1, 6);

			new TestUtils().CompareStyle(tempA1, resultA1);
			new TestUtils().CompareStyle(tempB1, resultB1);
			new TestUtils().CompareStyle(tempC1, resultC1);
			new TestUtils().CompareStyle(tempD1, resultD1);
			new TestUtils().CompareStyle(tempE1, resultE1);
			new TestUtils().CompareStyle(tempF1, resultF1);

			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
		}
	}


	#endregion

	#region << Group By >>
	[Fact]
	public void SpreadsheetData1_GroupBy()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateData1.xlsx";
			var template = new WvSpreadsheetFileTemplate
			{
				Template = new TestUtils().LoadWorkbookAsMemoryStream(templateFile),
				GroupDataByColumns = new List<string> { "sku" }
			};
			var dataSource = SampleData;
			dataSource.Rows[1]["sku"] = dataSource.Rows[0]["sku"];
			//When
			WvSpreadsheetFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			new TestUtils().GeneralResultChecks(result);
			Assert.Single(result!.Workbook!.Worksheets);
			Assert.NotNull(result!.ResultItems);
			Assert.Equal(4, result!.ResultItems.Count);
			Assert.NotNull(result!.ResultItems[0]!.Workbook);
			Assert.Single(result!.ResultItems[0]!.Workbook!.Worksheets);
			var ws = result!.ResultItems[0]!.Workbook!.Worksheets.First();
			Assert.Equal((string)SampleData.Rows[0]["name"], ws.Name);

			var row1SkuValueString = ws.Cell(1, 2).Value.ToString();
			var row2SkuValueString = ws.Cell(2, 2).Value.ToString();
			Assert.Equal(row1SkuValueString, row2SkuValueString);
		}
	}
	#endregion


}

﻿using ClosedXML.Excel;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using WebVella.DocumentTemplates.Engines.Excel;
using WebVella.DocumentTemplates.Tests.Models;

namespace WebVella.DocumentTemplates.Tests.Engines;
public class ExcelEngineTests : TestBase
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
			WvExcelFileTemplateResult? result = null;
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

	#region << Placement >>
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
			WvExcelFileTemplateResult? result = template.Process(dataSource);
			//Then
			GeneralResultChecks(result);
			Assert.Single(result!.Template!.Worksheets);
			Assert.Single(result!.Result!.Worksheets);
			Assert.Equal(2, result!.Contexts.Count);

			CheckRangeDimensions(result.Contexts[0].TemplateRange!, 1, 1, 1, 1);
			CheckRangeDimensions(result.Contexts[0].ResultRange!, 1, 1, 1, 1);
			CheckRangeDimensions(result.Contexts[1].TemplateRange!, 1, 2, 1, 3);
			CheckRangeDimensions(result.Contexts[1].ResultRange!, 1, 2, 1, 3);
			CheckCellPropertiesCopy(result);
			var tempWs = result.Template.Worksheets.First();
			var resultWs = result.Result.Worksheets.First();
			CompareRowProperties(tempWs.Row(1), resultWs.Row(1));
			CompareColumnProperties(tempWs.Column(1), resultWs.Column(1));
			CompareColumnProperties(tempWs.Column(2), resultWs.Column(2));

			SaveWorkbook(result.Result, templateFile);
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
			WvExcelFileTemplateResult? result = template.Process(dataSource);
			//Then
			GeneralResultChecks(result);
			Assert.Equal(2, result!.Template!.Worksheets.Count);
			Assert.Equal(2, result!.Result!.Worksheets.Count);
			Assert.Equal(3, result!.Contexts.Count);
			SaveWorkbook(result.Result, templateFile);
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
			WvExcelFileTemplateResult? result = template.Process(dataSource);
			//Then
			GeneralResultChecks(result);
			Assert.Single(result!.Template!.Worksheets);
			Assert.Single(result!.Result!.Worksheets);
			Assert.Equal(6, result!.Contexts.Count);
			CheckRangeDimensions(result.Contexts[0].TemplateRange!, 1, 1, 1, 1);
			CheckRangeDimensions(result.Contexts[0].ResultRange!, 1, 1, 5, 1);

			CheckRangeDimensions(result.Contexts[1].TemplateRange!, 1, 2, 1, 2);
			CheckRangeDimensions(result.Contexts[1].ResultRange!, 1, 2, 5, 2);

			CheckRangeDimensions(result.Contexts[2].TemplateRange!, 1, 3, 1, 3);
			CheckRangeDimensions(result.Contexts[2].ResultRange!, 1, 3, 1, 3);

			CheckRangeDimensions(result.Contexts[3].TemplateRange!, 1, 4, 1, 4);
			CheckRangeDimensions(result.Contexts[3].ResultRange!, 1, 4, 5, 4);

			CheckRangeDimensions(result.Contexts[4].TemplateRange!, 1, 5, 1, 5);
			CheckRangeDimensions(result.Contexts[4].ResultRange!, 1, 5, 1, 5);

			CheckRangeDimensions(result.Contexts[5].TemplateRange!, 1, 6, 1, 6);
			CheckRangeDimensions(result.Contexts[5].ResultRange!, 1, 6, 1, 6);


			SaveWorkbook(result.Result, templateFile);
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
			WvExcelFileTemplateResult? result = template.Process(dataSource);
			//Then
			GeneralResultChecks(result);
			Assert.Single(result!.Template!.Worksheets);
			Assert.Single(result!.Result!.Worksheets);
			Assert.Single(result!.Contexts);

			CheckRangeDimensions(result.Contexts[0].TemplateRange!, 1, 1, 1, 1);
			CheckRangeDimensions(result.Contexts[0].ResultRange!, 1, 1, 5, 1);

			SaveWorkbook(result.Result, templateFile);
		}
	}

	[Fact]
	public void Placement6_Data_Multiline()
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
			WvExcelFileTemplateResult? result = template.Process(dataSource);
			//Then
			GeneralResultChecks(result);
			Assert.Single(result!.Template!.Worksheets);
			Assert.Single(result!.Result!.Worksheets);
			Assert.Equal(2, result!.Contexts.Count);
			CheckRangeDimensions(result.Contexts[0].TemplateRange!, 1, 1, 1, 1);
			CheckRangeDimensions(result.Contexts[0].ResultRange!, 1, 1, 5, 1);
			CheckRangeDimensions(result.Contexts[1].TemplateRange!, 2, 1, 2, 1);
			CheckRangeDimensions(result.Contexts[1].ResultRange!, 6, 1, 10, 1);

			SaveWorkbook(result.Result, templateFile);
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
			WvExcelFileTemplateResult? result = template.Process(dataSource);
			//Then
			GeneralResultChecks(result);
			Assert.Single(result!.Template!.Worksheets);
			Assert.Single(result!.Result!.Worksheets);
			Assert.Equal(6, result!.Contexts.Count);
			CheckRangeDimensions(result.Contexts[0].TemplateRange!, 1, 1, 1, 1);
			CheckRangeDimensions(result.Contexts[0].ResultRange!, 1, 1, 1, 5);
			CheckRangeDimensions(result.Contexts[1].TemplateRange!, 1, 2, 1, 2);
			CheckRangeDimensions(result.Contexts[1].ResultRange!, 1, 6, 1, 6);
			CheckRangeDimensions(result.Contexts[2].TemplateRange!, 1, 3, 1, 3);
			CheckRangeDimensions(result.Contexts[2].ResultRange!, 1, 7, 5, 7);
			CheckRangeDimensions(result.Contexts[3].TemplateRange!, 2, 1, 2, 1);
			CheckRangeDimensions(result.Contexts[3].ResultRange!, 6, 1, 6, 5);
			CheckRangeDimensions(result.Contexts[4].TemplateRange!, 2, 2, 2, 2);
			CheckRangeDimensions(result.Contexts[4].ResultRange!, 6, 6, 6, 6);
			CheckRangeDimensions(result.Contexts[5].TemplateRange!, 2, 3, 2, 3);
			CheckRangeDimensions(result.Contexts[5].ResultRange!, 6, 7, 6, 7);
			SaveWorkbook(result.Result, templateFile);

		}
	}
	#endregion

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
			WvExcelFileTemplateResult? result = template.Process(dataSource);
			//Then
			GeneralResultChecks(result);
			Assert.Single(result!.Template!.Worksheets);
			Assert.Single(result!.Result!.Worksheets);
			var ws = result.Result.Worksheets.First();
			Assert.Equal((string)SampleData.Rows[0]["name"], ws.Name);

			for (int i = 0; i < SampleData.Rows.Count; i++)
			{
				var cellValueString = ws.Cell(i + 1, 1).Value.ToString();
				var rowValueString = SampleData.Rows[i]["position"]?.ToString();
				Assert.Equal(rowValueString, cellValueString);
			}

			SaveWorkbook(result.Result, templateFile);
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
			WvExcelFileTemplateResult? result = template.Process(dataSource);
			//Then
			GeneralResultChecks(result);
			Assert.Single(result!.Template!.Worksheets);
			Assert.Single(result!.Result!.Worksheets);
			var ws = result.Result.Worksheets.First();
			Assert.Equal((string)SampleData.Rows[0]["name"], ws.Name);

			for (int i = 0; i < SampleData.Rows.Count; i++)
			{
				var cellValueString = ws.Cell(1, i + 1).Value.ToString();
				var rowValueString = SampleData.Rows[i]["position"]?.ToString();
				Assert.Equal(rowValueString, cellValueString);
			}

			SaveWorkbook(result.Result, templateFile);
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
			WvExcelFileTemplateResult? result = template.Process(dataSource);
			//Then
			GeneralResultChecks(result);
			Assert.Single(result!.Template!.Worksheets);
			Assert.Single(result!.Result!.Worksheets);
			var ws = result.Result.Worksheets.First();
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

			SaveWorkbook(result.Result, templateFile);
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
					ds.Columns.Add(colName,typeof(string));
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
			var template = new WvExcelFileTemplate
			{
				Template = wb
			};
			var templateResult = new WvExcelFileTemplateResult() { Template = template.Template};
			var dataSource = ds;
			var culture = new CultureInfo("en-US");
			var engine = new WvExcelFileTemplate();
			//When
			var sw = new Stopwatch();
			long timeMS = 0;
			sw.Start();
			engine.ProcessExcelTemplatePlacement(templateResult,ds,culture);
			timeMS+=sw.ElapsedMilliseconds;
			sw.Restart();
			engine.ProcessExcelTemplateDependencies(templateResult,ds);
			timeMS+=sw.ElapsedMilliseconds;
			sw.Restart();
			engine.ProcessExcelTemplateData(templateResult,ds,culture);
			timeMS+=sw.ElapsedMilliseconds;
			sw.Stop();
			//Then
			SaveWorkbook(templateResult.Result!, templateFile);
			Assert.True(15 * 1000 > timeMS);
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
			WvExcelFileTemplateResult? result = template.Process(dataSource);
			//Then
			GeneralResultChecks(result);
			Assert.Single(result!.Template!.Worksheets);
			Assert.Single(result!.Result!.Worksheets);
			Assert.Equal(6, result!.Contexts.Count);
			SaveWorkbook(result.Result, templateFile);
		}
	}


	#endregion

	#region << Excel Function >>
	#endregion
}

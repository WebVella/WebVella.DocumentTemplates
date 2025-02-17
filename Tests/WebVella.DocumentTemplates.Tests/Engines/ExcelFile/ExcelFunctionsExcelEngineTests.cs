﻿using ClosedXML.Excel;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Engines.Excel;
using WebVella.DocumentTemplates.Engines.Excel.Utility;
using WebVella.DocumentTemplates.Tests.Models;

namespace WebVella.DocumentTemplates.Tests.Engines;
public partial class ExcelFunctionsExcelEngineTests : TestBase
{
	private static readonly object locker = new();
	public ExcelFunctionsExcelEngineTests() : base() { }

	#region << General >>
	[Fact]
	public void Excel_Function_SingleRange()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateExcelFunction1.xlsx";
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
			var worksheet = result!.ResultItems[0]!.Result!.Worksheets.First();
			Assert.Equal("position", worksheet.Cell(1, 1).Value.ToString());
			Assert.Equal("1", worksheet.Cell(2, 1).Value.ToString());
			Assert.Equal("2", worksheet.Cell(3, 1).Value.ToString());
			Assert.Equal("3", worksheet.Cell(4, 1).Value.ToString());
			Assert.Equal("4", worksheet.Cell(5, 1).Value.ToString());
			Assert.Equal("5", worksheet.Cell(6, 1).Value.ToString());
			IXLCell excelFunctionCell = worksheet.Cell(7, 1);
			Assert.True(excelFunctionCell.HasFormula);
			Assert.Equal("SUM(A2:A6)", excelFunctionCell.FormulaA1);

			SaveWorkbook(result!.ResultItems[0]!.Result!, templateFile);
		}
	}
	[Fact]
	public void Excel_Function_MultiRange()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateExcelFunction2.xlsx";
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
			var worksheet = result!.ResultItems[0]!.Result!.Worksheets.First();
			Assert.Equal("position", worksheet.Cell(1, 1).Value.ToString());
			Assert.Equal("1", worksheet.Cell(2, 1).Value.ToString());
			Assert.Equal("2", worksheet.Cell(3, 1).Value.ToString());
			Assert.Equal("3", worksheet.Cell(4, 1).Value.ToString());
			Assert.Equal("4", worksheet.Cell(5, 1).Value.ToString());
			Assert.Equal("5", worksheet.Cell(6, 1).Value.ToString());
			IXLCell excelFunctionCell = worksheet.Cell(7, 1);
			Assert.True(excelFunctionCell.HasFormula);
			Assert.Equal("SUM(A2:A6,B2:B6)", excelFunctionCell.FormulaA1);

			SaveWorkbook(result!.ResultItems[0]!.Result!, templateFile);
		}
	}

	[Fact]
	public void Excel_Function_Dependencies()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateExcelFunction3.xlsx";
			var template = new WvExcelFileTemplate
			{
				Template = LoadWorkbook(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvExcelFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			GeneralResultChecks(result);

			var positionContext1 = result!.TemplateContexts
				.FirstOrDefault(x => x.Range!.RangeAddress.FirstAddress.RowNumber == 2
					&& x.Range!.RangeAddress.FirstAddress.ColumnNumber == 1);
			var positionContext2 = result!.TemplateContexts
				.FirstOrDefault(x => x.Range!.RangeAddress.FirstAddress.RowNumber == 2
					&& x.Range!.RangeAddress.FirstAddress.ColumnNumber == 2);

			var functionContext1 = result!.TemplateContexts
				.FirstOrDefault(x => x.Range!.RangeAddress.FirstAddress.RowNumber == 3
					&& x.Range!.RangeAddress.FirstAddress.ColumnNumber == 1);
			var functionContext2 = result!.TemplateContexts
				.FirstOrDefault(x => x.Range!.RangeAddress.FirstAddress.RowNumber == 3
					&& x.Range!.RangeAddress.FirstAddress.ColumnNumber == 2);
			var functionContext3 = result!.TemplateContexts
				.FirstOrDefault(x => x.Range!.RangeAddress.FirstAddress.RowNumber == 3
					&& x.Range!.RangeAddress.FirstAddress.ColumnNumber == 3);

			Assert.NotNull(functionContext1);
			Assert.NotNull(functionContext2);
			Assert.NotNull(functionContext3);

			Assert.Single(functionContext1.ContextDependencies);
			Assert.Single(functionContext2.ContextDependencies);
			Assert.Equal(2, functionContext3.ContextDependencies.Count);

			Assert.Contains(positionContext1!.Id, functionContext1.ContextDependencies);
			Assert.Contains(positionContext2!.Id, functionContext2.ContextDependencies);

			Assert.Contains(functionContext1!.Id, functionContext3.ContextDependencies);
			Assert.Contains(functionContext2!.Id, functionContext3.ContextDependencies);


			Assert.Single(result!.Template!.Worksheets);
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Result);
			Assert.Single(result!.ResultItems[0]!.Result!.Worksheets);
			var worksheet = result!.ResultItems[0]!.Result!.Worksheets.First();
			Assert.Equal("position", worksheet.Cell(1, 1).Value.ToString());
			Assert.Equal("1", worksheet.Cell(2, 1).Value.ToString());
			Assert.Equal("2", worksheet.Cell(3, 1).Value.ToString());
			Assert.Equal("3", worksheet.Cell(4, 1).Value.ToString());
			Assert.Equal("4", worksheet.Cell(5, 1).Value.ToString());
			Assert.Equal("5", worksheet.Cell(6, 1).Value.ToString());
			IXLCell excelFunctionCell1 = worksheet.Cell(7, 1);
			IXLCell excelFunctionCell2 = worksheet.Cell(7, 2);
			IXLCell excelFunctionCell3 = worksheet.Cell(7, 3);
			Assert.True(excelFunctionCell1.HasFormula);
			Assert.True(excelFunctionCell2.HasFormula);
			Assert.True(excelFunctionCell3.HasFormula);
			Assert.Equal("SUM(A2:A6)", excelFunctionCell1.FormulaA1);
			Assert.Equal("SUM(B2:B6)", excelFunctionCell2.FormulaA1);
			Assert.Equal("SUM(A7:A7,B7:B7)", excelFunctionCell3.FormulaA1);

			SaveWorkbook(result!.ResultItems[0]!.Result!, templateFile);
		}
	}

	[Fact]
	public void Excel_Function_Dependencies_range()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateExcelFunction4.xlsx";
			var template = new WvExcelFileTemplate
			{
				Template = LoadWorkbook(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvExcelFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			GeneralResultChecks(result);

			var positionContext1 = result!.TemplateContexts
				.FirstOrDefault(x => x.Range!.RangeAddress.FirstAddress.RowNumber == 2
					&& x.Range!.RangeAddress.FirstAddress.ColumnNumber == 1);
			var positionContext2 = result!.TemplateContexts
				.FirstOrDefault(x => x.Range!.RangeAddress.FirstAddress.RowNumber == 2
					&& x.Range!.RangeAddress.FirstAddress.ColumnNumber == 2);

			var functionContext1 = result!.TemplateContexts
				.FirstOrDefault(x => x.Range!.RangeAddress.FirstAddress.RowNumber == 3
					&& x.Range!.RangeAddress.FirstAddress.ColumnNumber == 1);
			var functionContext2 = result!.TemplateContexts
				.FirstOrDefault(x => x.Range!.RangeAddress.FirstAddress.RowNumber == 3
					&& x.Range!.RangeAddress.FirstAddress.ColumnNumber == 2);
			var functionContext3 = result!.TemplateContexts
				.FirstOrDefault(x => x.Range!.RangeAddress.FirstAddress.RowNumber == 3
					&& x.Range!.RangeAddress.FirstAddress.ColumnNumber == 3);

			Assert.NotNull(functionContext1);
			Assert.NotNull(functionContext2);
			Assert.NotNull(functionContext3);

			Assert.Single(functionContext1.ContextDependencies);
			Assert.Single(functionContext2.ContextDependencies);
			Assert.Equal(2, functionContext3.ContextDependencies.Count);

			Assert.Contains(positionContext1!.Id, functionContext1.ContextDependencies);
			Assert.Contains(positionContext2!.Id, functionContext2.ContextDependencies);

			Assert.Contains(functionContext1!.Id, functionContext3.ContextDependencies);
			Assert.Contains(functionContext2!.Id, functionContext3.ContextDependencies);


			Assert.Single(result!.Template!.Worksheets);
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Result);
			Assert.Single(result!.ResultItems[0]!.Result!.Worksheets);
			var worksheet = result!.ResultItems[0]!.Result!.Worksheets.First();
			Assert.Equal("position", worksheet.Cell(1, 1).Value.ToString());
			Assert.Equal("1", worksheet.Cell(2, 1).Value.ToString());
			Assert.Equal("2", worksheet.Cell(3, 1).Value.ToString());
			Assert.Equal("3", worksheet.Cell(4, 1).Value.ToString());
			Assert.Equal("4", worksheet.Cell(5, 1).Value.ToString());
			Assert.Equal("5", worksheet.Cell(6, 1).Value.ToString());
			IXLCell excelFunctionCell1 = worksheet.Cell(7, 1);
			IXLCell excelFunctionCell2 = worksheet.Cell(7, 2);
			IXLCell excelFunctionCell3 = worksheet.Cell(7, 3);
			Assert.True(excelFunctionCell1.HasFormula);
			Assert.True(excelFunctionCell2.HasFormula);
			Assert.True(excelFunctionCell3.HasFormula);
			Assert.Equal("SUM(A2:A6)", excelFunctionCell1.FormulaA1);
			Assert.Equal("SUM(B2:B6)", excelFunctionCell2.FormulaA1);
			Assert.Equal("SUM(A7:A7,B7:B7)", excelFunctionCell3.FormulaA1);

			SaveWorkbook(result!.ResultItems[0]!.Result!, templateFile);
		}
	}

	#endregion

	#region << ABS >>
	[Fact]
	public void Excel_Function_ABS_SingleRange()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateExcelFunction-ABS-1.xlsx";
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
			var worksheet = result!.ResultItems[0]!.Result!.Worksheets.First();
			Assert.Equal("position", worksheet.Cell(1, 1).Value.ToString());
			Assert.Equal("1", worksheet.Cell(2, 1).Value.ToString());
			Assert.Equal("2", worksheet.Cell(3, 1).Value.ToString());
			Assert.Equal("3", worksheet.Cell(4, 1).Value.ToString());
			Assert.Equal("4", worksheet.Cell(5, 1).Value.ToString());
			Assert.Equal("5", worksheet.Cell(6, 1).Value.ToString());
			IXLCell excelFunctionCell = worksheet.Cell(7, 1);
			Assert.True(excelFunctionCell.HasFormula);
			Assert.Equal("ABS(A2:A6)", excelFunctionCell.FormulaA1);

			SaveWorkbook(result!.ResultItems[0]!.Result!, templateFile);
		}
	}
	#endregion

	#region << AVERAGE >>
	[Fact]
	public void Excel_Function_AVERAGE_SingleRange()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateExcelFunction-AVERAGE-1.xlsx";
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
			var worksheet = result!.ResultItems[0]!.Result!.Worksheets.First();
			Assert.Equal("position", worksheet.Cell(1, 1).Value.ToString());
			Assert.Equal("1", worksheet.Cell(2, 1).Value.ToString());
			Assert.Equal("2", worksheet.Cell(3, 1).Value.ToString());
			Assert.Equal("3", worksheet.Cell(4, 1).Value.ToString());
			Assert.Equal("4", worksheet.Cell(5, 1).Value.ToString());
			Assert.Equal("5", worksheet.Cell(6, 1).Value.ToString());
			IXLCell excelFunctionCell = worksheet.Cell(7, 1);
			Assert.True(excelFunctionCell.HasFormula);
			Assert.Equal("AVERAGE(A2:A6)", excelFunctionCell.FormulaA1);

			SaveWorkbook(result!.ResultItems[0]!.Result!, templateFile);
		}
	}
	#endregion

	#region << CONCAT >>
	[Fact]
	public void Excel_Function_CONCAT_SingleRange()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateExcelFunction-CONCAT-1.xlsx";
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
			var worksheet = result!.ResultItems[0]!.Result!.Worksheets.First();
			Assert.Equal("name", worksheet.Cell(1, 1).Value.ToString());
			Assert.Equal("item1", worksheet.Cell(2, 1).Value.ToString());
			Assert.Equal("item2", worksheet.Cell(3, 1).Value.ToString());
			Assert.Equal("item3", worksheet.Cell(4, 1).Value.ToString());
			Assert.Equal("item4", worksheet.Cell(5, 1).Value.ToString());
			Assert.Equal("item5", worksheet.Cell(6, 1).Value.ToString());
			IXLCell excelFunctionCell = worksheet.Cell(7, 1);
			Assert.True(excelFunctionCell.HasFormula);
			Assert.Equal("CONCAT(A2:A6)", excelFunctionCell.FormulaA1);

			SaveWorkbook(result!.ResultItems[0]!.Result!, templateFile);
		}
	}
	#endregion

	#region << MAX >>
	[Fact]
	public void Excel_Function_MAX_SingleRange()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateExcelFunction-MAX-1.xlsx";
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
			var worksheet = result!.ResultItems[0]!.Result!.Worksheets.First();
			Assert.Equal("position", worksheet.Cell(1, 1).Value.ToString());
			Assert.Equal("1", worksheet.Cell(2, 1).Value.ToString());
			Assert.Equal("2", worksheet.Cell(3, 1).Value.ToString());
			Assert.Equal("3", worksheet.Cell(4, 1).Value.ToString());
			Assert.Equal("4", worksheet.Cell(5, 1).Value.ToString());
			Assert.Equal("5", worksheet.Cell(6, 1).Value.ToString());
			IXLCell excelFunctionCell = worksheet.Cell(7, 1);
			Assert.True(excelFunctionCell.HasFormula);
			Assert.Equal("MAX(A2:A6)", excelFunctionCell.FormulaA1);

			SaveWorkbook(result!.ResultItems[0]!.Result!, templateFile);
		}
	}
	#endregion

	#region << MIN >>
	[Fact]
	public void Excel_Function_MIN_SingleRange()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateExcelFunction-MIN-1.xlsx";
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
			var worksheet = result!.ResultItems[0]!.Result!.Worksheets.First();
			Assert.Equal("position", worksheet.Cell(1, 1).Value.ToString());
			Assert.Equal("1", worksheet.Cell(2, 1).Value.ToString());
			Assert.Equal("2", worksheet.Cell(3, 1).Value.ToString());
			Assert.Equal("3", worksheet.Cell(4, 1).Value.ToString());
			Assert.Equal("4", worksheet.Cell(5, 1).Value.ToString());
			Assert.Equal("5", worksheet.Cell(6, 1).Value.ToString());
			IXLCell excelFunctionCell = worksheet.Cell(7, 1);
			Assert.True(excelFunctionCell.HasFormula);
			Assert.Equal("MIN(A2:A6)", excelFunctionCell.FormulaA1);

			SaveWorkbook(result!.ResultItems[0]!.Result!, templateFile);
		}
	}
	#endregion

	#region << SUM >>
	//in the general tests
	#endregion
}

using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Engines.Excel;
using WebVella.DocumentTemplates.Engines.Excel.Utility;
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

	#region << Template Context >>
	//Arguments
	[Fact]
	public void TemplateContext_Arguments()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateContext1.xlsx";
			var template = new WvExcelFileTemplate
			{
				Template = LoadWorkbook(templateFile)
			};
			WvExcelFileTemplateProcessResult? result = null;
			//When
			var action = () => template.ProcessExcelTemplateInitTemplateContexts(result);
			//Then
			var ex = Record.Exception(action);
			Assert.NotNull(ex);
			Assert.IsType<ArgumentException>(ex);
			var argEx = (ArgumentException)ex;
			Assert.Equal("result", argEx.ParamName);
			Assert.StartsWith("No result provided!", argEx.Message);
		}
	}

	[Fact]
	public void TemplateContext_Arguments_Success()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateContext1.xlsx";
			var template = new WvExcelFileTemplate
			{
				Template = LoadWorkbook(templateFile)
			};
			WvExcelFileTemplateProcessResult? result = new()
			{
				Template = template.Template
			};
			//When
			var action = () => template.ProcessExcelTemplateInitTemplateContexts(result);
			//Then
			var ex = Record.Exception(action);
			Assert.Null(ex);
		}
	}

	//Flow
	[Fact]
	public void TemplateContext_Flow_ShouldBeVerticalByDefault()
	{
		lock (locker)
		{
			//Given
			var wb = new XLWorkbook();
			var ws = wb.Worksheets.Add();
			ws.Cell(1, 1).Value = "test";

			var template = new WvExcelFileTemplate
			{
				Template = wb
			};
			WvExcelFileTemplateProcessResult result = new()
			{
				Template = template.Template
			};
			//When
			template.ProcessExcelTemplateInitTemplateContexts(result);
			//Then
			Assert.NotNull(result);
			Assert.Single(result.TemplateContexts);
			var context = result.TemplateContexts[0];
			Assert.NotEqual(Guid.Empty, context.Id);
			Assert.Equal(1, context.WorksheetPosition);
			Assert.NotNull(context.Range);
			Assert.Equal(WvExcelFileTemplateContextType.CellRange, context.Type);
			Assert.Equal(WvTemplateTagDataFlow.Vertical, context.Flow);
			Assert.Null(context.LeftContext);
			Assert.Null(context.TopContext);
			Assert.Null(context.ParentContext);
			Assert.Equal("A1:A1", context.Range.RangeAddress.ToString());
		}
	}

	[Fact]
	public void TemplateContext_Flow_ShouldBeVerticalIfNotAllHorizontal()
	{
		lock (locker)
		{
			//Given
			var wb = new XLWorkbook();
			var ws = wb.Worksheets.Add();
			ws.Cell(1, 1).Value = "{{test1}}{{test(F=H)}}";

			var template = new WvExcelFileTemplate
			{
				Template = wb
			};
			WvExcelFileTemplateProcessResult result = new()
			{
				Template = template.Template
			};
			//When
			template.ProcessExcelTemplateInitTemplateContexts(result);
			//Then
			Assert.NotNull(result);
			Assert.Single(result.TemplateContexts);
			var context = result.TemplateContexts[0];
			Assert.NotEqual(Guid.Empty, context.Id);
			Assert.Equal(1, context.WorksheetPosition);
			Assert.NotNull(context.Range);
			Assert.Equal(WvExcelFileTemplateContextType.CellRange, context.Type);
			Assert.Equal(WvTemplateTagDataFlow.Vertical, context.Flow);
			Assert.Null(context.LeftContext);
			Assert.Null(context.TopContext);
			Assert.Null(context.ParentContext);
			Assert.Equal("A1:A1", context.Range.RangeAddress.ToString());
		}
	}
	[Fact]
	public void TemplateContext_Flow_ShouldBeHorizontalIfExplicit()
	{
		lock (locker)
		{
			//Given
			var wb = new XLWorkbook();
			var ws = wb.Worksheets.Add();
			ws.Cell(1, 1).Value = "{{test(F=H)}}";

			var template = new WvExcelFileTemplate
			{
				Template = wb
			};
			WvExcelFileTemplateProcessResult result = new()
			{
				Template = template.Template
			};
			//When
			template.ProcessExcelTemplateInitTemplateContexts(result);
			//Then
			Assert.NotNull(result);
			Assert.Single(result.TemplateContexts);
			var context = result.TemplateContexts[0];
			Assert.NotEqual(Guid.Empty, context.Id);
			Assert.Equal(1, context.WorksheetPosition);
			Assert.NotNull(context.Range);
			Assert.Equal(WvExcelFileTemplateContextType.CellRange, context.Type);
			Assert.Equal(WvTemplateTagDataFlow.Horizontal, context.Flow);
			Assert.Null(context.LeftContext);
			Assert.Null(context.TopContext);
			Assert.Null(context.ParentContext);
			Assert.Equal("A1:A1", context.Range.RangeAddress.ToString());
		}
	}
	[Fact]
	public void TemplateContext_Flow_ShouldBeHorizontalIfAllHorizontal()
	{
		lock (locker)
		{
			//Given
			var wb = new XLWorkbook();
			var ws = wb.Worksheets.Add();
			ws.Cell(1, 1).Value = "{{test(F=H)}} some text {{test1(F=H)}}";

			var template = new WvExcelFileTemplate
			{
				Template = wb
			};
			WvExcelFileTemplateProcessResult result = new()
			{
				Template = template.Template
			};
			//When
			template.ProcessExcelTemplateInitTemplateContexts(result);
			//Then
			Assert.NotNull(result);
			Assert.Single(result.TemplateContexts);
			var context = result.TemplateContexts[0];
			Assert.NotEqual(Guid.Empty, context.Id);
			Assert.Equal(1, context.WorksheetPosition);
			Assert.NotNull(context.Range);
			Assert.Equal(WvExcelFileTemplateContextType.CellRange, context.Type);
			Assert.Equal(WvTemplateTagDataFlow.Horizontal, context.Flow);
			Assert.Null(context.LeftContext);
			Assert.Null(context.TopContext);
			Assert.Null(context.ParentContext);
			Assert.Equal("A1:A1", context.Range.RangeAddress.ToString());
		}
	}

	//Cell Context
	[Fact]
	public void TemplateContext_Context_LeftForVertical()
	{
		lock (locker)
		{
			//Given
			var wb = new XLWorkbook();
			var ws = wb.Worksheets.Add();
			ws.Cell(1, 1).Value = "test";
			ws.Cell(1, 2).Value = "test";

			var template = new WvExcelFileTemplate
			{
				Template = wb
			};
			WvExcelFileTemplateProcessResult result = new()
			{
				Template = template.Template
			};
			//When
			template.ProcessExcelTemplateInitTemplateContexts(result);
			//Then
			Assert.NotNull(result);
			Assert.Equal(2, result.TemplateContexts.Count);
			var contextA1 = result.TemplateContexts.GetByAddress(ws.Position, 1, 1).FirstOrDefault();
			var contextB1 = result.TemplateContexts.GetByAddress(ws.Position, 1, 2).FirstOrDefault();
			Assert.NotNull(contextA1);
			Assert.NotNull(contextB1);
			Assert.NotNull(contextB1.ParentContext);
			Assert.Equal(contextA1.Id, contextB1.ParentContext.Id);
		}
	}

	[Fact]
	public void TemplateContext_Context_LeftForVerticalTopIsFallBackForTheFirstColumn()
	{
		lock (locker)
		{
			//Given
			var wb = new XLWorkbook();
			var ws = wb.Worksheets.Add();
			ws.Cell(1, 1).Value = "test";
			ws.Cell(2, 1).Value = "test";

			var template = new WvExcelFileTemplate
			{
				Template = wb
			};
			WvExcelFileTemplateProcessResult result = new()
			{
				Template = template.Template
			};
			//When
			template.ProcessExcelTemplateInitTemplateContexts(result);
			//Then
			Assert.NotNull(result);
			Assert.Equal(2, result.TemplateContexts.Count);
			var contextA1 = result.TemplateContexts.GetByAddress(ws.Position, 1, 1).FirstOrDefault();
			var contextB1 = result.TemplateContexts.GetByAddress(ws.Position, 2, 1).FirstOrDefault();
			Assert.NotNull(contextA1);
			Assert.NotNull(contextB1);
			Assert.NotNull(contextB1.ParentContext);
			Assert.Equal(contextA1.Id, contextB1.ParentContext.Id);
		}
	}

	[Fact]
	public void TemplateContext_Context_LeftForVerticalParentIsMerge()
	{
		lock (locker)
		{
			//Given
			var wb = new XLWorkbook();
			var ws = wb.Worksheets.Add();
			var range = ws.Range(1, 1, 1, 3);
			range.Merge();
			range.Value = "test";
			ws.Cell(1, 4).Value = "test";

			var template = new WvExcelFileTemplate
			{
				Template = wb
			};
			WvExcelFileTemplateProcessResult result = new()
			{
				Template = template.Template
			};
			//When
			template.ProcessExcelTemplateInitTemplateContexts(result);
			//Then
			Assert.NotNull(result);
			Assert.Equal(2, result.TemplateContexts.Count);
			var contextA1 = result.TemplateContexts.GetByAddress(ws.Position, 1, 1).FirstOrDefault();
			var contextB1 = result.TemplateContexts.GetByAddress(ws.Position, 1, 4).FirstOrDefault();
			Assert.NotNull(contextA1);
			Assert.NotNull(contextB1);
			Assert.NotNull(contextB1.ParentContext);
			Assert.Equal(contextA1.Id, contextB1.ParentContext.Id);
		}
	}

	[Fact]
	public void TemplateContext_Context_MergeLeftForVerticalParent()
	{
		lock (locker)
		{
			//Given
			var wb = new XLWorkbook();
			var ws = wb.Worksheets.Add();
			ws.Cell(1, 1).Value = "test";
			var range = ws.Range(1, 2, 1, 4);
			range.Merge();
			range.Value = "test";

			var template = new WvExcelFileTemplate
			{
				Template = wb
			};
			WvExcelFileTemplateProcessResult result = new()
			{
				Template = template.Template
			};
			//When
			template.ProcessExcelTemplateInitTemplateContexts(result);
			//Then
			Assert.NotNull(result);
			Assert.Equal(2, result.TemplateContexts.Count);
			var contextA1 = result.TemplateContexts.GetByAddress(ws.Position, 1, 1).FirstOrDefault();
			var contextB1 = result.TemplateContexts.GetByAddress(ws.Position, 1, 2).FirstOrDefault();
			Assert.NotNull(contextA1);
			Assert.NotNull(contextB1);
			Assert.NotNull(contextB1.ParentContext);
			Assert.Equal(contextA1.Id, contextB1.ParentContext.Id);
		}
	}

	[Fact]
	public void TemplateContext_Context_TopForHorizontalParentIsMerge()
	{
		lock (locker)
		{
			//Given
			var wb = new XLWorkbook();
			var ws = wb.Worksheets.Add();
			var range = ws.Range(1, 1, 1, 3);
			range.Merge();
			range.Value = "test";
			ws.Cell(2, 1).Value = "test";

			var template = new WvExcelFileTemplate
			{
				Template = wb
			};
			WvExcelFileTemplateProcessResult result = new()
			{
				Template = template.Template
			};
			//When
			template.ProcessExcelTemplateInitTemplateContexts(result);
			//Then
			Assert.NotNull(result);
			Assert.Equal(2, result.TemplateContexts.Count);
			var contextA1 = result.TemplateContexts.GetByAddress(ws.Position, 1, 1).FirstOrDefault();
			var contextB1 = result.TemplateContexts.GetByAddress(ws.Position, 2, 1).FirstOrDefault();
			Assert.NotNull(contextA1);
			Assert.NotNull(contextB1);
			Assert.NotNull(contextB1.ParentContext);
			Assert.Equal(contextA1.Id, contextB1.ParentContext.Id);
		}
	}

	//Picture Context
	[Fact]
	public void TemplateContext_PictureContext_IsPresent()
	{
		lock (locker)
		{
			//Given
			var wb = new XLWorkbook();
			var ws = wb.Worksheets.Add();
			ws.Cell(1, 1).Value = "test";
			ws.Cell(1, 2).Value = "test";
			var imageFilename = "wv-logo.jpg";
			var picture = ws.AddPicture(LoadFileAsStream(imageFilename), imageFilename);
			picture.MoveTo(ws.Cell(3, 2));

			var template = new WvExcelFileTemplate
			{
				Template = wb
			};
			WvExcelFileTemplateProcessResult result = new()
			{
				Template = template.Template
			};
			//When
			template.ProcessExcelTemplateInitTemplateContexts(result);
			//Then
			Assert.NotNull(result);
			var pictureContextList = result.TemplateContexts.Where(x => x.Type == WvExcelFileTemplateContextType.Picture).ToList();
			Assert.NotNull(pictureContextList);
			Assert.Single(pictureContextList);
			var context = pictureContextList[0];
			Assert.NotNull(context.Picture);
			Assert.NotNull(context.TopContext);
			Assert.NotNull(context.TopContext.Range);
			Assert.NotNull(context.LeftContext);
			Assert.NotNull(context.LeftContext.Range);
			Assert.Equal(context.TopContext.Range.RangeAddress.FirstAddress.RowNumber, picture.TopLeftCell.Address.RowNumber - 1);
			Assert.Equal(context.LeftContext.Range.RangeAddress.FirstAddress.ColumnNumber, picture.TopLeftCell.Address.ColumnNumber - 1);
		}
	}
	#endregion

	#region << Template Context Dependencies >>
	[Fact]
	public void TemplateContext_Context_CheckDependencies()
	{
		lock (locker)
		{
			//Given
			var wb = new XLWorkbook();
			var ws = wb.Worksheets.Add();
			ws.Cell(1, 1).Value = "{{position}}";
			ws.Cell(1, 2).Value = "{{=SUM(A1:A1)}}";

			var template = new WvExcelFileTemplate
			{
				Template = wb
			};
			//When
			var result = template.Process(TypedData, DefaultCulture);
			//Then
			Assert.NotNull(result);
			Assert.Equal(2, result.TemplateContexts.Count);
			var a1Context = result.TemplateContexts.GetByAddress(ws.Position, 1, 1).FirstOrDefault();
			var b1Context = result.TemplateContexts.GetByAddress(ws.Position, 1, 2).FirstOrDefault();
			Assert.NotNull(a1Context);
			Assert.NotNull(b1Context);
			Assert.Single(b1Context.ContextDependencies);
			Assert.Equal(a1Context.Id, b1Context.ContextDependencies.First());
		}
	}

	[Fact]
	public void TemplateContext_Context_CheckDependencies2()
	{
		lock (locker)
		{
			//Given
			var wb = new XLWorkbook();
			var ws = wb.Worksheets.Add();
			ws.Cell(1, 1).Value = "{{position}}";
			ws.Cell(1, 2).Value = "{{=SUM(B100:C200)}}";

			var template = new WvExcelFileTemplate
			{
				Template = wb
			};
			//When
			var result = template.Process(TypedData, DefaultCulture);
			//Then
			Assert.NotNull(result);
			Assert.Equal(2, result.TemplateContexts.Count);
			var a1Context = result.TemplateContexts.GetByAddress(ws.Position, 1, 1).FirstOrDefault();
			var b1Context = result.TemplateContexts.GetByAddress(ws.Position, 1, 2).FirstOrDefault();
			Assert.NotNull(a1Context);
			Assert.NotNull(b1Context);
			Assert.Empty(b1Context.ContextDependencies);
		}
	}

	[Fact]
	public void TemplateContext_Context_CheckDependencies3()
	{
		lock (locker)
		{
			//Given
			var wb = new XLWorkbook();
			var ws = wb.Worksheets.Add();
			ws.Cell(1, 1).Value = "{{position}}";
			ws.Cell(1, 2).Value = "{{=SUM(A1:C5)}}";

			var template = new WvExcelFileTemplate
			{
				Template = wb
			};
			//When
			var result = template.Process(TypedData, DefaultCulture);
			//Then
			Assert.NotNull(result);
			Assert.Equal(2, result.TemplateContexts.Count);
			var a1Context = result.TemplateContexts.GetByAddress(ws.Position, 1, 1).FirstOrDefault();
			var b1Context = result.TemplateContexts.GetByAddress(ws.Position, 1, 2).FirstOrDefault();
			Assert.NotNull(a1Context);
			Assert.NotNull(b1Context);
			Assert.Single(b1Context.ContextDependencies);
			Assert.Equal(a1Context.Id, b1Context.ContextDependencies.First());
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
			WvExcelFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			GeneralResultChecks(result);
			Assert.Single(result!.Template!.Worksheets);
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Result);
			Assert.NotNull(result!.ResultItems[0]!.Result!.Worksheets);
			Assert.Single(result!.ResultItems[0]!.Result!.Worksheets);
			Assert.Equal(2, result!.ResultItems[0]!.Contexts.Count);

			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[0].TemplateRange!, 1, 1, 1, 1);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[0].ResultRange!, 1, 1, 1, 1);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[1].TemplateRange!, 1, 2, 1, 3);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[1].ResultRange!, 1, 2, 1, 3);
			CheckCellPropertiesCopy(result!.ResultItems[0]!);
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
			//Then
			GeneralResultChecks(result);
			Assert.Single(result!.Template!.Worksheets);
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Result);

			Assert.Single(result!.ResultItems[0]!.Result!.Worksheets);
			Assert.Equal(6, result!.ResultItems[0]!.Contexts.Count);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[0].TemplateRange!, 1, 1, 1, 1);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[0].ResultRange!, 1, 1, 5, 1);

			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[1].TemplateRange!, 1, 2, 1, 2);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[1].ResultRange!, 1, 2, 5, 2);

			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[2].TemplateRange!, 1, 3, 1, 3);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[2].ResultRange!, 1, 3, 1, 3);

			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[3].TemplateRange!, 1, 4, 1, 4);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[3].ResultRange!, 1, 4, 5, 4);

			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[4].TemplateRange!, 1, 5, 1, 5);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[4].ResultRange!, 1, 5, 1, 5);

			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[5].TemplateRange!, 1, 6, 1, 6);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[5].ResultRange!, 1, 6, 1, 6);


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
			Assert.Single(result!.Template!.Worksheets);
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Result);
			Assert.Single(result!.ResultItems[0]!.Result!.Worksheets);
			Assert.Single(result!.ResultItems[0]!.Contexts);

			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[0].TemplateRange!, 1, 1, 1, 1);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[0].ResultRange!, 1, 1, 5, 1);

			SaveWorkbook(result!.ResultItems[0]!.Result!, templateFile);
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
			WvExcelFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			GeneralResultChecks(result);
			Assert.Single(result!.Template!.Worksheets);
			Assert.Single(result!.Template!.Worksheets);
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Result);
			Assert.Single(result!.ResultItems[0]!.Result!.Worksheets);
			Assert.Equal(2, result!.ResultItems[0]!.Contexts.Count);
			CheckRangeDimensions(result!.ResultItems[0].Contexts[0].TemplateRange!, 1, 1, 1, 1);
			CheckRangeDimensions(result!.ResultItems[0].Contexts[0].ResultRange!, 1, 1, 5, 1);
			CheckRangeDimensions(result!.ResultItems[0].Contexts[1].TemplateRange!, 2, 1, 2, 1);
			CheckRangeDimensions(result!.ResultItems[0].Contexts[1].ResultRange!, 6, 1, 10, 1);

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
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Result);
			Assert.Single(result!.ResultItems[0]!.Result!.Worksheets);
			Assert.Equal(6, result!.ResultItems[0]!.Contexts.Count);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[0].TemplateRange!, 1, 1, 1, 1);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[0].ResultRange!, 1, 1, 1, 5);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[1].TemplateRange!, 1, 2, 1, 2);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[1].ResultRange!, 1, 6, 1, 6);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[2].TemplateRange!, 1, 3, 1, 3);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[2].ResultRange!, 1, 7, 5, 7);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[3].TemplateRange!, 2, 1, 2, 1);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[3].ResultRange!, 6, 1, 6, 5);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[4].TemplateRange!, 2, 2, 2, 2);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[4].ResultRange!, 6, 6, 6, 6);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[5].TemplateRange!, 2, 3, 2, 3);
			CheckRangeDimensions(result!.ResultItems[0]!.Contexts[5].ResultRange!, 6, 7, 6, 7);
			SaveWorkbook(result!.ResultItems[0]!.Result!, templateFile);

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
			var rowCount = 100;
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
			var template = new WvExcelFileTemplate
			{
				Template = wb
			};
			var templateResultItem = new WvExcelFileTemplateProcessResultItem();
			var dataSource = ds;
			var culture = new CultureInfo("en-US");
			var engine = new WvExcelFileTemplate();
			//When
			var sw = new Stopwatch();
			long timeMS = 0;
			sw.Start();
			engine.ProcessExcelTemplatePlacementAndContexts(template.Template, templateResultItem, ds, culture);
			timeMS += sw.ElapsedMilliseconds;
			sw.Restart();
			engine.ProcessExcelTemplateDependencies(templateResultItem, ds);
			timeMS += sw.ElapsedMilliseconds;
			sw.Restart();
			engine.ProcessExcelTemplateData(templateResultItem, ds, culture);
			timeMS += sw.ElapsedMilliseconds;
			sw.Stop();
			//Then
			SaveWorkbook(templateResultItem.Result!, templateFile);
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
			WvExcelFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			GeneralResultChecks(result);
			Assert.Single(result!.Template!.Worksheets);
			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Result);
			Assert.Single(result!.ResultItems[0]!.Result!.Worksheets);
			Assert.Equal(6, result!.ResultItems[0]!.Contexts.Count);
			SaveWorkbook(result!.ResultItems[0]!.Result!, templateFile);
		}
	}


	#endregion

	#region << Excel Function >>
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

	#region << Embedded items >>
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
	#endregion

	#region << Utilities >>

	//GetApplicableRangeForFunctionTag
	[Fact]
	public void GetApplicableRangeForFunctionTag_Test1()
	{
		lock (locker)
		{
			//Given
			var tags = WvTemplateUtility.GetTagsFromTemplate("{{=SUM(A1)}}");
			Assert.NotNull(tags);
			Assert.Single(tags);
			//When
			var ranges = tags[0].GetApplicableRangeForFunctionTag();
			//Then
			Assert.NotNull(ranges);
			Assert.Single(ranges);
			var range = ranges[0];
			Assert.Equal(1, range.FirstRow);
			Assert.Equal(1, range.FirstColumn);
			Assert.Equal(1, range.LastRow);
			Assert.Equal(1, range.LastColumn);
		}
	}

	[Fact]
	public void GetApplicableRangeForFunctionTag_Test2()
	{
		lock (locker)
		{
			//Given
			var tags = WvTemplateUtility.GetTagsFromTemplate("{{=SUM(A1:A1)}}");
			Assert.NotNull(tags);
			Assert.Single(tags);
			//When
			var ranges = tags[0].GetApplicableRangeForFunctionTag();
			//Then
			Assert.NotNull(ranges);
			Assert.Single(ranges);
			var range = ranges[0];
			Assert.Equal(1, range.FirstRow);
			Assert.Equal(1, range.FirstColumn);
			Assert.Equal(1, range.LastRow);
			Assert.Equal(1, range.LastColumn);
		}
	}

	[Fact]
	public void GetApplicableRangeForFunctionTag_Test3()
	{
		lock (locker)
		{
			//Given
			var tags = WvTemplateUtility.GetTagsFromTemplate("{{=SUM(B1:C3)}}");
			Assert.NotNull(tags);
			Assert.Single(tags);
			//When
			var ranges = tags[0].GetApplicableRangeForFunctionTag();
			//Then
			Assert.NotNull(ranges);
			Assert.Single(ranges);
			var range = ranges[0];
			Assert.Equal(1, range.FirstRow);
			Assert.Equal(2, range.FirstColumn);
			Assert.Equal(3, range.LastRow);
			Assert.Equal(3, range.LastColumn);
		}
	}


	//GetRangeFromString
	[Fact]
	public void GetRangeFromString_Test1()
	{
		lock (locker)
		{
			//Given
			var address = "name!A1:C4";

			//When
			var result = WvExcelRangeHelpers.GetRangeFromString(address);
			//Then
			Assert.NotNull(result);
			Assert.Equal("name", result.Worksheet);
			Assert.Equal(1, result.FirstRow);
			Assert.Equal(1, result.FirstColumn);
			Assert.Equal(4, result.LastRow);
			Assert.Equal(3, result.LastColumn);
		}
	}

	[Fact]
	public void GetRangeFromString_Test2()
	{
		lock (locker)
		{
			//Given
			var address = "name!C5";

			//When
			var result = WvExcelRangeHelpers.GetRangeFromString(address);
			//Then
			Assert.NotNull(result);
			Assert.Equal("name", result.Worksheet);
			Assert.Equal(5, result.FirstRow);
			Assert.Equal(3, result.FirstColumn);
			Assert.Equal(5, result.LastRow);
			Assert.Equal(3, result.LastColumn);
		}
	}

	[Fact]
	public void GetRangeFromString_Test3()
	{
		lock (locker)
		{
			//Given
			var address = "A1:C4";

			//When
			var result = WvExcelRangeHelpers.GetRangeFromString(address);
			//Then
			Assert.NotNull(result);
			Assert.True(String.IsNullOrWhiteSpace(result.Worksheet));
			Assert.Equal(1, result.FirstRow);
			Assert.Equal(1, result.FirstColumn);
			Assert.Equal(4, result.LastRow);
			Assert.Equal(3, result.LastColumn);
		}
	}

	[Fact]
	public void GetRangeFromString_Test4()
	{
		lock (locker)
		{
			//Given
			var address = "C5";

			//When
			var result = WvExcelRangeHelpers.GetRangeFromString(address);
			//Then
			Assert.NotNull(result);
			Assert.True(String.IsNullOrWhiteSpace(result.Worksheet));
			Assert.Equal(5, result.FirstRow);
			Assert.Equal(3, result.FirstColumn);
			Assert.Equal(5, result.LastRow);
			Assert.Equal(3, result.LastColumn);
		}
	}

	[Fact]
	public void GetRangeFromString_Test5()
	{
		lock (locker)
		{
			//Given
			var address = "invalid";

			//When
			var result = WvExcelRangeHelpers.GetRangeFromString(address);
			//Then
			Assert.Null(result);
		}
	}

	[Fact]
	public void GetRangeFromString_Test6()
	{
		lock (locker)
		{
			//Given
			var address = "invalid123";

			//When
			var result = WvExcelRangeHelpers.GetRangeFromString(address);
			//Then
			Assert.Null(result);
		}
	}

	#endregion
}

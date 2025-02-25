using ClosedXML.Excel;
using System.Data;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Engines.ExcelFile;
using WebVella.DocumentTemplates.Engines.ExcelFile.Utility;
using WebVella.DocumentTemplates.Tests.Models;
using WebVella.DocumentTemplates.Tests.Utils;

namespace WebVella.DocumentTemplates.Tests.Engines;
public partial class TemplateContextsExcelEngineTests : TestBase
{
	private static readonly object locker = new();
	public TemplateContextsExcelEngineTests() : base() { }

	#region << Template Context >>
	//Arguments
	[Fact]
	public void TemplateContext_Arguments()
	{
		lock (locker)
		{
			//Given
			WvExcelFileTemplateProcessResult? result = null;
			//When
			var action = () => new WvExcelFileEngineUtility().ProcessExcelTemplateInitTemplateContexts(result);
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
			WvExcelFileTemplateProcessResult? result = new()
			{
				Template = new TestUtils().LoadWorkbookAsMemoryStream(templateFile)
			};
			//When
			var action = () => new WvExcelFileEngineUtility().ProcessExcelTemplateInitTemplateContexts(result);
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

			MemoryStream ms = new();
			wb.SaveAs(ms);
			WvExcelFileTemplateProcessResult result = new()
			{
				Template = ms
			};
			//When
			new WvExcelFileEngineUtility().ProcessExcelTemplateInitTemplateContexts(result);
			//Then
			Assert.NotNull(result);
			Assert.Single(result.TemplateContexts);
			var context = result.TemplateContexts[0];
			Assert.NotEqual(Guid.Empty, context.Id);
			Assert.Equal(1, context.Worksheet!.Position);
			Assert.NotNull(context.Range);
			Assert.Equal(WvExcelFileTemplateContextType.CellRange, context.Type);
			Assert.Equal(WvTemplateTagDataFlow.Vertical, context.ForcedFlow);
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

			MemoryStream ms = new();
			wb.SaveAs(ms);
			WvExcelFileTemplateProcessResult result = new()
			{
				Template = ms
			};
			//When
			new WvExcelFileEngineUtility().ProcessExcelTemplateInitTemplateContexts(result);
			//Then
			Assert.NotNull(result);
			Assert.Single(result.TemplateContexts);
			var context = result.TemplateContexts[0];
			Assert.NotEqual(Guid.Empty, context.Id);
			Assert.Equal(1, context.Worksheet!.Position);
			Assert.NotNull(context.Range);
			Assert.Equal(WvExcelFileTemplateContextType.CellRange, context.Type);
			Assert.Equal(WvTemplateTagDataFlow.Vertical, context.ForcedFlow);
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
			MemoryStream ms = new();
			wb.SaveAs(ms);
			WvExcelFileTemplateProcessResult result = new()
			{
				Template = ms
			};
			//When
			new WvExcelFileEngineUtility().ProcessExcelTemplateInitTemplateContexts(result);
			//Then
			Assert.NotNull(result);
			Assert.Single(result.TemplateContexts);
			var context = result.TemplateContexts[0];
			Assert.NotEqual(Guid.Empty, context.Id);
			Assert.Equal(1, context.Worksheet!.Position);
			Assert.NotNull(context.Range);
			Assert.Equal(WvExcelFileTemplateContextType.CellRange, context.Type);
			Assert.Equal(WvTemplateTagDataFlow.Horizontal, context.ForcedFlow);
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

			MemoryStream ms = new();
			wb.SaveAs(ms);
			WvExcelFileTemplateProcessResult result = new()
			{
				Template = ms
			};
			//When
			new WvExcelFileEngineUtility().ProcessExcelTemplateInitTemplateContexts(result);
			//Then
			Assert.NotNull(result);
			Assert.Single(result.TemplateContexts);
			var context = result.TemplateContexts[0];
			Assert.NotEqual(Guid.Empty, context.Id);
			Assert.Equal(1, context.Worksheet!.Position);
			Assert.NotNull(context.Range);
			Assert.Equal(WvExcelFileTemplateContextType.CellRange, context.Type);
			Assert.Equal(WvTemplateTagDataFlow.Horizontal, context.ForcedFlow);
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
			MemoryStream ms = new();
			wb.SaveAs(ms);
			WvExcelFileTemplateProcessResult result = new()
			{
				Template = ms
			};
			//When
			new WvExcelFileEngineUtility().ProcessExcelTemplateInitTemplateContexts(result);
			//Then
			Assert.NotNull(result);
			Assert.Equal(2, result.TemplateContexts.Count);
			var contextA1 = result.TemplateContexts.GetByAddress(ws.Position, 1, 1).FirstOrDefault();
			var contextB1 = result.TemplateContexts.GetByAddress(ws.Position, 1, 2).FirstOrDefault();
			Assert.NotNull(contextA1);
			Assert.NotNull(contextB1);
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

			MemoryStream ms = new();
			wb.SaveAs(ms);
			WvExcelFileTemplateProcessResult result = new()
			{
				Template = ms
			};
			//When
			new WvExcelFileEngineUtility().ProcessExcelTemplateInitTemplateContexts(result);
			//Then
			Assert.NotNull(result);
			Assert.Equal(2, result.TemplateContexts.Count);
			var contextA1 = result.TemplateContexts.GetByAddress(ws.Position, 1, 1).FirstOrDefault();
			var contextB1 = result.TemplateContexts.GetByAddress(ws.Position, 2, 1).FirstOrDefault();
			Assert.NotNull(contextA1);
			Assert.NotNull(contextB1);
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
			MemoryStream ms = new();
			wb.SaveAs(ms);
			WvExcelFileTemplateProcessResult result = new()
			{
				Template = ms
			};
			//When
			new WvExcelFileEngineUtility().ProcessExcelTemplateInitTemplateContexts(result);
			//Then
			Assert.NotNull(result);
			Assert.Equal(2, result.TemplateContexts.Count);
			var contextA1 = result.TemplateContexts.GetByAddress(ws.Position, 1, 1).FirstOrDefault();
			var contextB1 = result.TemplateContexts.GetByAddress(ws.Position, 1, 4).FirstOrDefault();
			Assert.NotNull(contextA1);
			Assert.NotNull(contextB1);
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

			MemoryStream ms = new();
			wb.SaveAs(ms);
			WvExcelFileTemplateProcessResult result = new()
			{
				Template = ms
			};
			//When
			new WvExcelFileEngineUtility().ProcessExcelTemplateInitTemplateContexts(result);
			//Then
			Assert.NotNull(result);
			Assert.Equal(2, result.TemplateContexts.Count);
			var contextA1 = result.TemplateContexts.GetByAddress(ws.Position, 1, 1).FirstOrDefault();
			var contextB1 = result.TemplateContexts.GetByAddress(ws.Position, 1, 2).FirstOrDefault();
			Assert.NotNull(contextA1);
			Assert.NotNull(contextB1);
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

			MemoryStream ms = new();
			wb.SaveAs(ms);
			WvExcelFileTemplateProcessResult result = new()
			{
				Template = ms
			};
			//When
			new WvExcelFileEngineUtility().ProcessExcelTemplateInitTemplateContexts(result);
			//Then
			Assert.NotNull(result);
			Assert.Equal(2, result.TemplateContexts.Count);
			var contextA1 = result.TemplateContexts.GetByAddress(ws.Position, 1, 1).FirstOrDefault();
			var contextB1 = result.TemplateContexts.GetByAddress(ws.Position, 2, 1).FirstOrDefault();
			Assert.NotNull(contextA1);
			Assert.NotNull(contextB1);
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
			var picture = ws.AddPicture(new TestUtils().LoadFileAsStream(imageFilename), imageFilename);
			picture.MoveTo(ws.Cell(3, 2));

			MemoryStream ms = new();
			wb.SaveAs(ms);
			WvExcelFileTemplateProcessResult result = new()
			{
				Template = ms
			};
			//When
			new WvExcelFileEngineUtility().ProcessExcelTemplateInitTemplateContexts(result);
			//Then
			Assert.NotNull(result);
			var pictureContextList = result.TemplateContexts.Where(x => x.Type == WvExcelFileTemplateContextType.Picture).ToList();
			Assert.NotNull(pictureContextList);
			Assert.Single(pictureContextList);
			var context = pictureContextList[0];
			Assert.NotNull(context.Picture);
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
			
			MemoryStream? ms = new();
			wb.SaveAs(ms);
			var template = new WvExcelFileTemplate
			{
				Template = ms
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

			MemoryStream? ms = new();
			wb.SaveAs(ms);
			var template = new WvExcelFileTemplate
			{
				Template = ms
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

			MemoryStream? ms = new();
			wb.SaveAs(ms);
			var template = new WvExcelFileTemplate
			{
				Template = ms
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

	#region << Template Context Expansion>>
	[Fact]
	public void DataFlow_Test1()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateContextFlow1.xlsx";
			var template = new WvExcelFileTemplate
			{
				Template = new TestUtils().LoadWorkbookAsMemoryStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvExcelFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			new TestUtils().GeneralResultChecks(result);
			Assert.NotNull(result);
			Assert.NotNull(result.Workbook);
			Assert.Single(result.Workbook.Worksheets);
			var ws = result.Workbook.Worksheets.First();
			Assert.Equal(10, result.TemplateContexts.Count);
			var a1Context = result.TemplateContexts.GetByAddress(ws.Position, 1, 1).FirstOrDefault();
			var b1Context = result.TemplateContexts.GetByAddress(ws.Position, 1, 2).FirstOrDefault();
			var a2Context = result.TemplateContexts.GetByAddress(ws.Position, 2, 1).FirstOrDefault();
			var b2Context = result.TemplateContexts.GetByAddress(ws.Position, 2, 2).FirstOrDefault();
			var a3Context = result.TemplateContexts.GetByAddress(ws.Position, 3, 1).FirstOrDefault();
			var b3Context = result.TemplateContexts.GetByAddress(ws.Position, 3, 2).FirstOrDefault();
			var a4Context = result.TemplateContexts.GetByAddress(ws.Position, 4, 1).FirstOrDefault();
			var b4Context = result.TemplateContexts.GetByAddress(ws.Position, 4, 2).FirstOrDefault();
			var a5Context = result.TemplateContexts.GetByAddress(ws.Position, 5, 1).FirstOrDefault();
			var b5Context = result.TemplateContexts.GetByAddress(ws.Position, 5, 2).FirstOrDefault();

			Assert.NotNull(a1Context);
			Assert.NotNull(b1Context);
			Assert.NotNull(a2Context);
			Assert.NotNull(b2Context);
			Assert.NotNull(a3Context);
			Assert.NotNull(b3Context);
			Assert.NotNull(a4Context);
			Assert.NotNull(b4Context);
			Assert.NotNull(a5Context);
			Assert.NotNull(b5Context);

			Assert.Equal(WvTemplateTagDataFlow.Vertical, a1Context.Flow);
			Assert.Equal(WvTemplateTagDataFlow.Vertical, b1Context.Flow);
			Assert.Equal(WvTemplateTagDataFlow.Vertical, a2Context.Flow);
			Assert.Equal(WvTemplateTagDataFlow.Vertical, b2Context.Flow);
			Assert.Equal(WvTemplateTagDataFlow.Vertical, a3Context.Flow);
			Assert.Equal(WvTemplateTagDataFlow.Vertical, b3Context.Flow);
			Assert.Equal(WvTemplateTagDataFlow.Vertical, a4Context.Flow);
			Assert.Equal(WvTemplateTagDataFlow.Vertical, b4Context.Flow);
			Assert.Equal(WvTemplateTagDataFlow.Vertical, a5Context.Flow);
			Assert.Equal(WvTemplateTagDataFlow.Vertical, b5Context.Flow);
			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
		}
	}
	[Fact]
	public void DataFlow_Test2()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateContextFlow2.xlsx";
			var template = new WvExcelFileTemplate
			{
				Template = new TestUtils().LoadWorkbookAsMemoryStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvExcelFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			new TestUtils().GeneralResultChecks(result);
			Assert.NotNull(result);
			Assert.NotNull(result.Workbook);
			Assert.Single(result.Workbook.Worksheets);
			var ws = result.Workbook.Worksheets.First();
			Assert.Equal(8, result.TemplateContexts.Count);
			var a1Context = result.TemplateContexts.GetByAddress(ws.Position, 1, 1).FirstOrDefault();
			var b1Context = result.TemplateContexts.GetByAddress(ws.Position, 1, 2).FirstOrDefault();
			var a2Context = result.TemplateContexts.GetByAddress(ws.Position, 2, 1).FirstOrDefault();
			var b2Context = result.TemplateContexts.GetByAddress(ws.Position, 2, 2).FirstOrDefault();
			var a3Context = result.TemplateContexts.GetByAddress(ws.Position, 3, 1).FirstOrDefault();
			var b3Context = result.TemplateContexts.GetByAddress(ws.Position, 3, 2).FirstOrDefault();
			var a4Context = result.TemplateContexts.GetByAddress(ws.Position, 4, 1).FirstOrDefault();
			var b4Context = result.TemplateContexts.GetByAddress(ws.Position, 4, 2).FirstOrDefault();

			Assert.NotNull(a1Context);
			Assert.NotNull(b1Context);
			Assert.NotNull(a2Context);
			Assert.NotNull(b2Context);
			Assert.NotNull(a3Context);
			Assert.NotNull(b3Context);
			Assert.NotNull(a4Context);
			Assert.NotNull(b4Context);

			Assert.Equal(WvTemplateTagDataFlow.Vertical, a1Context.Flow);
			Assert.Equal(WvTemplateTagDataFlow.Horizontal, b1Context.ForcedFlow);
			Assert.Equal(WvTemplateTagDataFlow.Horizontal, b1Context.Flow);
			Assert.Equal(WvTemplateTagDataFlow.Vertical, a2Context.Flow);
			Assert.Equal(WvTemplateTagDataFlow.Vertical, b2Context.Flow);
			Assert.Equal(WvTemplateTagDataFlow.Vertical, a3Context.Flow);
			Assert.Equal(WvTemplateTagDataFlow.Vertical, b3Context.Flow);
			Assert.Equal(WvTemplateTagDataFlow.Vertical, a4Context.Flow);
			Assert.Equal(WvTemplateTagDataFlow.Vertical, b4Context.Flow);
			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
		}
	}
	[Fact]
	public void DataFlow_Test3()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateContextFlow3.xlsx";
			var template = new WvExcelFileTemplate
			{
				Template = new TestUtils().LoadWorkbookAsMemoryStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvExcelFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			new TestUtils().GeneralResultChecks(result);
			Assert.NotNull(result);
			Assert.NotNull(result.Workbook);
			Assert.Single(result.Workbook.Worksheets);
			var ws = result.Workbook.Worksheets.First();
			Assert.Equal(8, result.TemplateContexts.Count);
			var a1Context = result.TemplateContexts.GetByAddress(ws.Position, 1, 1).FirstOrDefault();
			var b1Context = result.TemplateContexts.GetByAddress(ws.Position, 1, 2).FirstOrDefault();
			var a2Context = result.TemplateContexts.GetByAddress(ws.Position, 2, 1).FirstOrDefault();
			var b2Context = result.TemplateContexts.GetByAddress(ws.Position, 2, 2).FirstOrDefault();
			var a3Context = result.TemplateContexts.GetByAddress(ws.Position, 3, 1).FirstOrDefault();
			var b3Context = result.TemplateContexts.GetByAddress(ws.Position, 3, 2).FirstOrDefault();
			var a4Context = result.TemplateContexts.GetByAddress(ws.Position, 4, 1).FirstOrDefault();
			var b4Context = result.TemplateContexts.GetByAddress(ws.Position, 4, 2).FirstOrDefault();

			Assert.NotNull(a1Context);
			Assert.NotNull(b1Context);
			Assert.NotNull(a2Context);
			Assert.NotNull(b2Context);
			Assert.NotNull(a3Context);
			Assert.NotNull(b3Context);
			Assert.NotNull(a4Context);
			Assert.NotNull(b4Context);

			Assert.Equal(WvTemplateTagDataFlow.Vertical, a1Context.Flow);
			Assert.Equal(WvTemplateTagDataFlow.Horizontal, b1Context.ForcedFlow);
			Assert.Equal(WvTemplateTagDataFlow.Horizontal, b1Context.Flow);
			Assert.Equal(WvTemplateTagDataFlow.Vertical, a2Context.Flow);
			Assert.Equal(WvTemplateTagDataFlow.Horizontal, b2Context.ForcedFlow);
			Assert.Equal(WvTemplateTagDataFlow.Horizontal, b2Context.Flow);
			Assert.Equal(WvTemplateTagDataFlow.Vertical, a3Context.Flow);
			Assert.Equal(WvTemplateTagDataFlow.Horizontal, b3Context.ForcedFlow);
			Assert.Equal(WvTemplateTagDataFlow.Horizontal, b3Context.Flow);
			Assert.Equal(WvTemplateTagDataFlow.Vertical, a4Context.Flow);
			Assert.Equal(WvTemplateTagDataFlow.Horizontal, b4Context.ForcedFlow);
			Assert.Equal(WvTemplateTagDataFlow.Horizontal, b4Context.Flow);
			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
		}
	}
	[Fact]
	public void DataFlow_Test4()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateContextFlow4.xlsx";
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
			Assert.NotNull(result);
			Assert.NotNull(result.Workbook);
			Assert.Single(result.Workbook.Worksheets);
			var ws = result.Workbook.Worksheets.First();
			Assert.Equal(16, result.TemplateContexts.Count);
			var a1Context = result.TemplateContexts.GetByAddress(ws.Position, 1, 1).FirstOrDefault();
			var b1Context = result.TemplateContexts.GetByAddress(ws.Position, 1, 2).FirstOrDefault();
			var c1Context = result.TemplateContexts.GetByAddress(ws.Position, 1, 3).FirstOrDefault();
			var d1Context = result.TemplateContexts.GetByAddress(ws.Position, 1, 4).FirstOrDefault();
			var a2Context = result.TemplateContexts.GetByAddress(ws.Position, 2, 1).FirstOrDefault();
			var b2Context = result.TemplateContexts.GetByAddress(ws.Position, 2, 2).FirstOrDefault();
			var a3Context = result.TemplateContexts.GetByAddress(ws.Position, 3, 1).FirstOrDefault();
			var b3Context = result.TemplateContexts.GetByAddress(ws.Position, 3, 2).FirstOrDefault();
			var a4Context = result.TemplateContexts.GetByAddress(ws.Position, 4, 1).FirstOrDefault();
			var b4Context = result.TemplateContexts.GetByAddress(ws.Position, 4, 2).FirstOrDefault();

			Assert.NotNull(a1Context);
			Assert.NotNull(b1Context);
			Assert.NotNull(c1Context);
			Assert.NotNull(d1Context);
			Assert.NotNull(a2Context);
			Assert.NotNull(b2Context);
			Assert.NotNull(a3Context);
			Assert.NotNull(b3Context);
			Assert.NotNull(a4Context);
			Assert.NotNull(b4Context);

			//A1
			Assert.Null(a1Context.ForcedContext);
			Assert.Null(a1Context.ParentContext);

			Assert.NotNull(a1Context.ForcedFlow);
			Assert.Equal(WvTemplateTagDataFlow.Vertical, a1Context.ForcedFlow);
			Assert.Equal(WvTemplateTagDataFlow.Vertical, a1Context.Flow);

			//B1
			Assert.Null(b1Context.ForcedContext);
			Assert.NotNull(b1Context.ParentContext);
			Assert.Equal(a1Context.Id, b1Context.ParentContext.Id);

			Assert.NotNull(b1Context.ForcedFlow);
			Assert.Equal(WvTemplateTagDataFlow.Horizontal, b1Context.ForcedFlow);
			Assert.Equal(WvTemplateTagDataFlow.Horizontal, b1Context.Flow);

			//C1
			Assert.Null(c1Context.ForcedContext);
			Assert.Null(c1Context.ParentContext);
			Assert.True(c1Context.IsNullContextForced);

			Assert.Null(c1Context.ForcedFlow);
			Assert.Equal(WvTemplateTagDataFlow.Vertical, c1Context.Flow);

			//D1
			Assert.Null(d1Context.ForcedContext);
			Assert.Null(d1Context.ParentContext);

			Assert.NotNull(d1Context.ForcedFlow);
			Assert.Equal(WvTemplateTagDataFlow.Horizontal, d1Context.ForcedFlow);
			Assert.Equal(WvTemplateTagDataFlow.Horizontal, d1Context.Flow);

			//A2
			Assert.Null(a2Context.ForcedContext);
			Assert.NotNull(a2Context.ParentContext);
			Assert.Equal(a1Context.Id, a2Context.ParentContext.Id);
			Assert.Equal(WvTemplateTagDataFlow.Vertical, a2Context.Flow);

			//B2
			Assert.Null(b2Context.ForcedFlow);
			Assert.NotNull(b2Context.ForcedContext);
			Assert.Equal(b1Context.Id, b2Context.ForcedContext.Id);
			Assert.Equal(b1Context.Flow, b2Context.Flow);

			//A3
			Assert.Null(a3Context.ForcedContext);
			Assert.NotNull(a3Context.ParentContext);
			Assert.Equal(a2Context.Id, a3Context.ParentContext.Id);
			Assert.Equal(WvTemplateTagDataFlow.Vertical, a3Context.Flow);

			//B3
			Assert.Null(b3Context.ForcedFlow);
			Assert.Null(b3Context.ForcedContext);
			Assert.Equal(WvTemplateTagDataFlow.Vertical, b3Context.Flow);
			Assert.NotNull(b3Context.ParentContext);
			Assert.Equal(a3Context.Id, b3Context.ParentContext.Id);

			//A4
			Assert.Null(a4Context.ForcedContext);
			Assert.NotNull(a4Context.ParentContext);
			Assert.Equal(a3Context.Id, a4Context.ParentContext.Id);
			Assert.Equal(WvTemplateTagDataFlow.Vertical, a4Context.Flow);

			//B2
			Assert.Null(b4Context.ForcedFlow);
			Assert.NotNull(b4Context.ForcedContext);
			Assert.Equal(b1Context.Id, b4Context.ForcedContext.Id);
			Assert.Equal(b1Context.Flow, b4Context.Flow);
			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
		}
	}

	[Fact]
	public void DataFlow_Test5()
	{
		lock (locker)
		{
			//Given
			var templateFile = "TemplateContextFlow5.xlsx";
			var template = new WvExcelFileTemplate
			{
				Template = new TestUtils().LoadWorkbookAsMemoryStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvExcelFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			new TestUtils().GeneralResultChecks(result);
			Assert.NotNull(result);
			Assert.NotNull(result.Workbook);
			Assert.Single(result.Workbook.Worksheets);
			Assert.Equal(2, result.TemplateContexts.Count);

			Assert.NotNull(result!.ResultItems);
			Assert.Single(result!.ResultItems);
			Assert.NotNull(result!.ResultItems[0]!.Workbook);
			Assert.NotNull(result!.ResultItems[0]!.Workbook!.Worksheets);
			Assert.Single(result!.ResultItems[0]!.Workbook!.Worksheets);
			var ws = result!.ResultItems[0]!.Workbook!.Worksheets.First();
			Assert.Equal(2, result!.ResultItems[0]!.Contexts.Count);
			new TestUtils().CheckRangeDimensions(result!.ResultItems[0]!.Contexts[0].Range!, 1, 1, 1, 1);
			new TestUtils().CheckRangeDimensions(result!.ResultItems[0]!.Contexts[1].Range!, 2, 1, 6, 1);

			Assert.Equal("position", ws.Cell(1, 1).Value.ToString());
			Assert.Equal(1, ws.Cell(2, 1).Value);
			Assert.Equal(2, ws.Cell(3, 1).Value);
			Assert.Equal(3, ws.Cell(4, 1).Value);
			Assert.Equal(4, ws.Cell(5, 1).Value);
			Assert.Equal(5, ws.Cell(6, 1).Value);

			new TestUtils().SaveWorkbook(result!.ResultItems[0]!.Workbook!, templateFile);
		}
	}
	#endregion
}

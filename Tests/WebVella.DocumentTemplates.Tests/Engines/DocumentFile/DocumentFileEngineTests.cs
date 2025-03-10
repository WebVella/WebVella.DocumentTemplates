using WebVella.DocumentTemplates.Engines.DocumentFile;
using WebVella.DocumentTemplates.Tests.Models;
using WebVella.DocumentTemplates.Tests.Utils;
using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace WebVella.DocumentTemplates.Tests.Engines;
public class DocumentFileEngineTests : TestBase
{
	private static readonly object locker = new();
	public DocumentFileEngineTests() : base() { }

	#region << Paragraph >>
	[Fact]
	public void DocumentFile_Paragraph_1()
	{
		lock (locker)
		{
			//Given
			var utils = new TestUtils();
			var templateFile = "Template-Paragraph-1.docx";
			var template = new WvDocumentFileTemplate
			{
				Template = new TestUtils().LoadFileAsStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvDocumentFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			utils.GeneralResultChecks(result);
			var paragraphs = result.ResultItems[0].WordDocument.MainDocumentPart.Document.Body.Descendants<Word.Paragraph>();
			Assert.NotNull(paragraphs);
			Assert.Single(paragraphs);
			var paragraph = paragraphs.First();
			Assert.Equal("item1,item2,item3,item4,item5",paragraph.InnerText);
			utils.SaveFileFromStream(result!.ResultItems[0]!.Result!, templateFile);
		}
	}

	[Fact]
	public void DocumentFile_Paragraph_2()
	{
		lock (locker)
		{
			//Given
			var utils = new TestUtils();
			var templateFile = "Template-Paragraph-2.docx";
			var template = new WvDocumentFileTemplate
			{
				Template = new TestUtils().LoadFileAsStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvDocumentFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			utils.GeneralResultChecks(result);
			var paragraphs = result.ResultItems[0].WordDocument.MainDocumentPart.Document.Body.Descendants<Word.Paragraph>();
			Assert.NotNull(paragraphs);
			Assert.Single(paragraphs);
			var paragraph = paragraphs.First();
			Assert.Equal("item1,item2,item3,item4,item5",paragraph.InnerText);
			var markRunProp = paragraph.Descendants<Word.ParagraphMarkRunProperties>().FirstOrDefault();
			Assert.NotNull(markRunProp);
			var boldProp = paragraph.Descendants<Word.Bold>().FirstOrDefault();
			Assert.NotNull(boldProp);
			var fontSizeProp = paragraph.Descendants<Word.FontSize>().FirstOrDefault();
			Assert.NotNull(fontSizeProp);
			Assert.Equal("32",fontSizeProp.Val);
			var colorProp = paragraph.Descendants<Word.Color>().FirstOrDefault();
			Assert.NotNull(colorProp);
			Assert.Equal("FF0000",colorProp.Val);
			utils.SaveFileFromStream(result!.ResultItems[0]!.Result!, templateFile);
		}
	}

	[Fact]
	public void DocumentFile_Paragraph_3()
	{
		lock (locker)
		{
			//Given
			var utils = new TestUtils();
			var templateFile = "Template-Paragraph-3.docx";
			var template = new WvDocumentFileTemplate
			{
				Template = new TestUtils().LoadFileAsStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvDocumentFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			utils.GeneralResultChecks(result);
			var paragraphs = result.ResultItems[0].WordDocument.MainDocumentPart.Document.Body.Descendants<Word.Paragraph>();
			Assert.NotNull(paragraphs);
			Assert.Single(paragraphs);
			var paragraph = paragraphs.First();
			Assert.Equal("Lorem ipsum item1,item2,item3,item4,item5 Lorem ipsum",paragraph.InnerText);
			utils.SaveFileFromStream(result!.ResultItems[0]!.Result!, templateFile);
		}
	}

	[Fact]
	public void DocumentFile_Paragraph_4()
	{
		lock (locker)
		{
			//Given
			var utils = new TestUtils();
			var templateFile = "Template-Paragraph-4.docx";
			var template = new WvDocumentFileTemplate
			{
				Template = new TestUtils().LoadFileAsStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvDocumentFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			utils.GeneralResultChecks(result);
			var paragraphs = result.ResultItems[0].WordDocument.MainDocumentPart.Document.Body.Descendants<Word.Paragraph>();
			Assert.NotNull(paragraphs);
			Assert.Single(paragraphs);
			var paragraph = paragraphs.First();
			Assert.Equal("Lorem ipsum item1,item2,item3,item4,item5 Lorem ipsum",paragraph.InnerText);

			var paragraphRuns = paragraph.Descendants<Word.Run>().ToList();
			Assert.Equal(3,paragraphRuns.Count());
			Assert.Equal("Lorem ipsum ",paragraphRuns[0].InnerText);
			Assert.Equal("item1,item2,item3,item4,item5",paragraphRuns[1].InnerText);
			Assert.Equal(" Lorem ipsum",paragraphRuns[2].InnerText);

			utils.SaveFileFromStream(result!.ResultItems[0]!.Result!, templateFile);
		}
	}

	[Fact]
	public void DocumentFile_Paragraph_5()
	{
		lock (locker)
		{
			//Given
			var utils = new TestUtils();
			var templateFile = "Template-Paragraph-5.docx";
			var template = new WvDocumentFileTemplate
			{
				Template = new TestUtils().LoadFileAsStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvDocumentFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			utils.GeneralResultChecks(result);
			var paragraphs = result.ResultItems[0].WordDocument.MainDocumentPart.Document.Body.Descendants<Word.Paragraph>().ToList();
			Assert.NotNull(paragraphs);
			Assert.Equal(6,paragraphs.Count);
			Assert.Equal("{{name",paragraphs[0].InnerText);
			Assert.Equal("(S=’,’)}}",paragraphs[1].InnerText);
			Assert.Equal("{{Gubergren et in.",paragraphs[2].InnerText);
			Assert.Equal("{{Gubergren.}}",paragraphs[3].InnerText);
			Assert.Equal("{{Veniam.",paragraphs[4].InnerText);
			Assert.Equal("option.}}",paragraphs[5].InnerText);


			utils.SaveFileFromStream(result!.ResultItems[0]!.Result!, templateFile);
		}
	}

	[Fact]
	public void DocumentFile_Paragraph_6()
	{
		lock (locker)
		{
			//Given
			var utils = new TestUtils();
			var templateFile = "Template-Paragraph-6.docx";
			var template = new WvDocumentFileTemplate
			{
				Template = new TestUtils().LoadFileAsStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvDocumentFileTemplateProcessResult? result = template.Process(dataSource);
			utils.SaveFileFromStream(result!.ResultItems[0]!.Result!, templateFile);
			//Then
			utils.GeneralResultChecks(result);
			var paragraphs = result.ResultItems[0].WordDocument.MainDocumentPart.Document.Body.Descendants<Word.Paragraph>().ToList();
			Assert.NotNull(paragraphs);
			Assert.Equal(3,paragraphs.Count);
			Assert.Equal("item1 volutpat.",paragraphs[0].InnerText);
			Assert.Equal("Gubergren item1 erat",paragraphs[1].InnerText);
			Assert.Equal("Gubergren lorem item1 erat.",paragraphs[2].InnerText);


			utils.SaveFileFromStream(result!.ResultItems[0]!.Result!, templateFile);
		}
	}
	#endregion
	#region << HeaderFooter >>
	[Fact]
	public void DocumentFile_HeaderFooter_1()
	{
		lock (locker)
		{
			//Given
			var utils = new TestUtils();
			var templateFile = "Template-HeaderFooter-1.docx";
			var template = new WvDocumentFileTemplate
			{
				Template = new TestUtils().LoadFileAsStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvDocumentFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			utils.GeneralResultChecks(result);
			var mainPart = result.ResultItems[0].WordDocument.MainDocumentPart;
			Assert.Single(mainPart.HeaderParts);
			Assert.Single(mainPart.FooterParts);
			var header = mainPart.HeaderParts.First().Header;
			var footer = mainPart.FooterParts.First().Footer;

			var headerParagraph = header.Descendants<Word.Paragraph>().FirstOrDefault();
			Assert.NotNull(headerParagraph);
			Assert.Equal("item1",headerParagraph.InnerText);

			var footerParagraph = footer.Descendants<Word.Paragraph>().FirstOrDefault();
			Assert.NotNull(footerParagraph);
			Assert.Equal("item2",footerParagraph.InnerText);

			utils.SaveFileFromStream(result!.ResultItems[0]!.Result!, templateFile);
		}
	}

	#endregion

	#region << Table >>
	[Fact]
	public void DocumentFile_Table()
	{
		lock (locker)
		{
			//Given
			var utils = new TestUtils();
			var templateFile = "Template-Table1.docx";
			var template = new WvDocumentFileTemplate
			{
				Template = new TestUtils().LoadFileAsStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvDocumentFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			utils.GeneralResultChecks(result);
			utils.SaveFileFromStream(result!.ResultItems[0]!.Result!, templateFile);
		}
	}

	#endregion

	#region << Complex >>
	[Fact]
	public void DocumentFile_Complex_1()
	{
		lock (locker)
		{
			//Given
			var utils = new TestUtils();
			var templateFile = "Template-Complex-1.docx";
			var template = new WvDocumentFileTemplate
			{
				Template = new TestUtils().LoadFileAsStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvDocumentFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			utils.GeneralResultChecks(result);
			utils.SaveFileFromStream(result!.ResultItems[0]!.Result!, templateFile);
		}
	}
	#endregion
}

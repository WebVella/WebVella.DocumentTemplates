using WebVella.DocumentTemplates.Engines.DocumentFile;
using WebVella.DocumentTemplates.Tests.Models;
using WebVella.DocumentTemplates.Tests.Utils;
using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace WebVella.DocumentTemplates.Tests.Engines;

public class DocumentFileEngineTests : TestBase
{
    private static readonly object locker = new();

    public DocumentFileEngineTests() : base()
    {
    }

    #region << Paragraph >>

    [Fact]
    public void DocumentFile_Flow_Horizontal()
    {
        lock (locker)
        {
            //Given
            var utils = new TestUtils();
            var templateFile = "Template-Flow-Horizontal.docx";
            var template = new WvDocumentFileTemplate
            {
                Template = new TestUtils().LoadFileAsStream(templateFile)
            };
            var dataSource = SampleData;
            //When
            WvDocumentFileTemplateProcessResult? result = template.Process(dataSource);
            //Then
            utils.GeneralResultChecks(result);
            var paragraphs = result.ResultItems[0].WordDocument.MainDocumentPart.Document.Body
                .Descendants<Word.Paragraph>();
            Assert.NotNull(paragraphs);
            Assert.Single(paragraphs);
            var paragraph = paragraphs.First();
            Assert.Equal("item1item2item3item4item5", paragraph.InnerText);
            utils.SaveFileFromStream(result!.ResultItems[0]!.Result!, templateFile);
        }
    }

    [Fact]
    public void DocumentFile_Flow_Vertical()
    {
        lock (locker)
        {
            //Given
            var utils = new TestUtils();
            var templateFile = "Template-Flow-Vertical.docx";
            var template = new WvDocumentFileTemplate
            {
                Template = new TestUtils().LoadFileAsStream(templateFile)
            };
            var dataSource = SampleData;
            //When
            WvDocumentFileTemplateProcessResult? result = template.Process(dataSource);
            //Then
            utils.GeneralResultChecks(result);
            var paragraphs = result.ResultItems[0].WordDocument.MainDocumentPart.Document.Body
                .Descendants<Word.Paragraph>();
            Assert.NotNull(paragraphs);
            Assert.Single(paragraphs);
            var paragraph = paragraphs.First();
            Assert.Single(paragraph.ChildElements);
            var run = paragraph.ChildElements[0];
            Assert.Equal(9, run.ChildElements.Count);
            Assert.IsType<Word.Text>(run.ChildElements[0]);
            Assert.IsType<Word.Break>(run.ChildElements[1]);
            Assert.IsType<Word.Text>(run.ChildElements[2]);
            Assert.IsType<Word.Break>(run.ChildElements[3]);
            Assert.IsType<Word.Text>(run.ChildElements[4]);
            Assert.IsType<Word.Break>(run.ChildElements[5]);
            Assert.IsType<Word.Text>(run.ChildElements[6]);
            Assert.IsType<Word.Break>(run.ChildElements[7]);
            Assert.IsType<Word.Text>(run.ChildElements[8]);
            utils.SaveFileFromStream(result!.ResultItems[0]!.Result!, templateFile);
        }
    }

    [Fact]
    public void DocumentFile_Flow_Vertical_2()
    {
        lock (locker)
        {
            //Given
            var utils = new TestUtils();
            var templateFile = "Template-Flow-Vertical-2.docx";
            var template = new WvDocumentFileTemplate
            {
                Template = new TestUtils().LoadFileAsStream(templateFile)
            };
            var dataSource = SampleData;
            //When
            WvDocumentFileTemplateProcessResult? result = template.Process(dataSource);
            //Then
            utils.GeneralResultChecks(result);
            var paragraphs = result.ResultItems[0].WordDocument.MainDocumentPart.Document.Body
                .Descendants<Word.Paragraph>();
            Assert.NotNull(paragraphs);
            Assert.Single(paragraphs);
            var paragraph = paragraphs.First();
            Assert.Single(paragraph.ChildElements);
            var run = paragraph.ChildElements[0];
            Assert.Equal(9, run.ChildElements.Count);
            Assert.IsType<Word.Text>(run.ChildElements[0]);
            Assert.IsType<Word.Break>(run.ChildElements[1]);
            Assert.IsType<Word.Text>(run.ChildElements[2]);
            Assert.IsType<Word.Break>(run.ChildElements[3]);
            Assert.IsType<Word.Text>(run.ChildElements[4]);
            Assert.IsType<Word.Break>(run.ChildElements[5]);
            Assert.IsType<Word.Text>(run.ChildElements[6]);
            Assert.IsType<Word.Break>(run.ChildElements[7]);
            Assert.IsType<Word.Text>(run.ChildElements[8]);
            utils.SaveFileFromStream(result!.ResultItems[0]!.Result!, templateFile);
        }
    }

    [Fact]
    public void DocumentFile_Flow_Vertical_3()
    {
        lock (locker)
        {
            //Given
            var utils = new TestUtils();
            var templateFile = "Template-Flow-Vertical-3.docx";
            var template = new WvDocumentFileTemplate
            {
                Template = new TestUtils().LoadFileAsStream(templateFile)
            };
            var dataSource = SampleData;
            //When
            WvDocumentFileTemplateProcessResult? result = template.Process(dataSource);
            //Then
            utils.GeneralResultChecks(result);
            var paragraphs = result.ResultItems[0].WordDocument.MainDocumentPart.Document.Body
                .Descendants<Word.Paragraph>();
            Assert.NotNull(paragraphs);
            Assert.Single(paragraphs);
            var paragraph = paragraphs.First();
            Assert.Equal(4, paragraph.ChildElements.Count);
            var run1 = paragraph.ChildElements[0];
            var run2 = paragraph.ChildElements[1];
            var run3 = paragraph.ChildElements[2];
            var run4 = paragraph.ChildElements[3];

            Assert.Single(run1.ChildElements);
            Assert.IsType<Word.Text>(run1.ChildElements[0]);

            Assert.Equal(9, run2.ChildElements.Count);
            Assert.IsType<Word.Text>(run2.ChildElements[0]);
            Assert.IsType<Word.Break>(run2.ChildElements[1]);
            Assert.IsType<Word.Text>(run2.ChildElements[2]);
            Assert.IsType<Word.Break>(run2.ChildElements[3]);
            Assert.IsType<Word.Text>(run2.ChildElements[4]);
            Assert.IsType<Word.Break>(run2.ChildElements[5]);
            Assert.IsType<Word.Text>(run2.ChildElements[6]);
            Assert.IsType<Word.Break>(run2.ChildElements[7]);
            Assert.IsType<Word.Text>(run2.ChildElements[8]);

            Assert.Single(run3.ChildElements);
            Assert.IsType<Word.Text>(run3.ChildElements[0]);

            Assert.Equal(9, run2.ChildElements.Count);
            Assert.IsType<Word.Text>(run4.ChildElements[0]);
            Assert.IsType<Word.Break>(run4.ChildElements[1]);
            Assert.IsType<Word.Text>(run4.ChildElements[2]);
            Assert.IsType<Word.Break>(run4.ChildElements[3]);
            Assert.IsType<Word.Text>(run4.ChildElements[4]);
            Assert.IsType<Word.Break>(run4.ChildElements[5]);
            Assert.IsType<Word.Text>(run4.ChildElements[6]);
            Assert.IsType<Word.Break>(run4.ChildElements[7]);
            Assert.IsType<Word.Text>(run4.ChildElements[8]);
            utils.SaveFileFromStream(result!.ResultItems[0]!.Result!, templateFile);
        }
    }

    #endregion


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
            var paragraphs = result.ResultItems[0].WordDocument.MainDocumentPart.Document.Body
                .Descendants<Word.Paragraph>();
            Assert.NotNull(paragraphs);
            Assert.Single(paragraphs);
            var paragraph = paragraphs.First();
            Assert.Equal("item1,item2,item3,item4,item5", paragraph.InnerText);
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
            var paragraphs = result.ResultItems[0].WordDocument.MainDocumentPart.Document.Body
                .Descendants<Word.Paragraph>();
            Assert.NotNull(paragraphs);
            Assert.Single(paragraphs);
            var paragraph = paragraphs.First();
            Assert.Equal("item1,item2,item3,item4,item5", paragraph.InnerText);
            Assert.Equal(2,paragraph.ChildElements.Count);
            var run = paragraph.Descendants<Word.Run>().FirstOrDefault();
            Assert.NotNull(run);
            Assert.Equal(2, run.ChildElements.Count);
            var runProp = run.Descendants<Word.RunProperties>().FirstOrDefault();
            Assert.NotNull(runProp);
            var boldProp = runProp.Descendants<Word.Bold>().FirstOrDefault();
            Assert.NotNull(boldProp);
            var fontSizeProp = runProp.Descendants<Word.FontSize>().FirstOrDefault();
            Assert.NotNull(fontSizeProp);
            Assert.Equal("32", fontSizeProp.Val);
            var colorProp = runProp.Descendants<Word.Color>().FirstOrDefault();
            Assert.NotNull(colorProp);
            Assert.Equal("FF0000", colorProp.Val);
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
            var paragraphs = result.ResultItems[0].WordDocument.MainDocumentPart.Document.Body
                .Descendants<Word.Paragraph>();
            Assert.NotNull(paragraphs);
            Assert.Single(paragraphs);
            var paragraph = paragraphs.First();
            Assert.Equal("Lorem ipsum item1,item2,item3,item4,item5 Lorem ipsum", paragraph.InnerText);
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
            var paragraphs = result.ResultItems[0].WordDocument.MainDocumentPart.Document.Body
                .Descendants<Word.Paragraph>();
            Assert.NotNull(paragraphs);
            Assert.Single(paragraphs);
            var paragraph = paragraphs.First();
            Assert.Equal("Lorem ipsum item1,item2,item3,item4,item5 Lorem ipsum", paragraph.InnerText);

            var paragraphRuns = paragraph.Descendants<Word.Run>().ToList();
            Assert.Equal(3, paragraphRuns.Count());
            Assert.Equal("Lorem ipsum ", paragraphRuns[0].InnerText);
            Assert.Equal("item1,item2,item3,item4,item5", paragraphRuns[1].InnerText);
            Assert.Equal(" Lorem ipsum", paragraphRuns[2].InnerText);

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
            var paragraphs = result.ResultItems[0].WordDocument.MainDocumentPart.Document.Body
                .Descendants<Word.Paragraph>().ToList();
            Assert.NotNull(paragraphs);
            Assert.Equal(6, paragraphs.Count);
            Assert.Equal("{{name", paragraphs[0].InnerText);
            Assert.Equal("(S=’,’)}}", paragraphs[1].InnerText);
            Assert.Equal("{{Gubergren et in.", paragraphs[2].InnerText);
            Assert.Equal("{{Gubergren.}}", paragraphs[3].InnerText);
            Assert.Equal("{{Veniam.", paragraphs[4].InnerText);
            Assert.Equal("option.}}", paragraphs[5].InnerText);


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
            var paragraphs = result.ResultItems[0].WordDocument.MainDocumentPart.Document.Body
                .Descendants<Word.Paragraph>().ToList();
            Assert.NotNull(paragraphs);
            Assert.Equal(3, paragraphs.Count);
            Assert.Equal("item1 volutpat.", paragraphs[0].InnerText);
            Assert.Equal("Gubergren item1 erat", paragraphs[1].InnerText);
            Assert.Equal("Gubergren lorem item1 erat.", paragraphs[2].InnerText);


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
            Assert.Equal("item1", headerParagraph.InnerText);

            var footerParagraph = footer.Descendants<Word.Paragraph>().FirstOrDefault();
            Assert.NotNull(footerParagraph);
            Assert.Equal("item2", footerParagraph.InnerText);

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
            utils.SaveFileFromStream(result!.ResultItems[0]!.Result!, templateFile);
            //Then
            utils.GeneralResultChecks(result);
            var tableElList = result.ResultItems[0].WordDocument.MainDocumentPart.Document.Descendants<Word.Table>();
            Assert.Single(tableElList);
            var tableEl = tableElList.First();
            var tableRows = tableEl.Descendants<Word.TableRow>().ToList();
            Assert.Equal(6, tableRows.Count);
            for (int i = 0; i < tableRows.Count; i++)
            {
                if (i == 0)
                    utils.CheckWordTableRowContents(tableRows[i], new List<string> { "Position", "Name", "price" });
                else
                    utils.CheckWordTableRowContents(tableRows[i], new List<string>
                    {
                        SampleData.Rows[i - 1]["position"].ToString() ?? String.Empty,
                        SampleData.Rows[i - 1]["name"].ToString() ?? String.Empty,
                        SampleData.Rows[i - 1]["price"].ToString() ?? String.Empty,
                    });
            }

            utils.SaveFileFromStream(result!.ResultItems[0]!.Result!, templateFile);
        }
    }

    [Fact]
    public void DocumentFile_Table2()
    {
        lock (locker)
        {
            //Given
            var utils = new TestUtils();
            var templateFile = "Template-Table2.docx";
            var template = new WvDocumentFileTemplate
            {
                Template = utils.LoadFileAsStream(templateFile)
            };
            var dataSource = SampleData;
            //When
            WvDocumentFileTemplateProcessResult? result = template.Process(dataSource);
            //Then
            utils.GeneralResultChecks(result);
            var tableElList = result.ResultItems[0].WordDocument.MainDocumentPart.Document.Descendants<Word.Table>();
            Assert.Single(tableElList);
            var tableEl = tableElList.First();
            var tableRows = tableEl.Descendants<Word.TableRow>().ToList();
            Assert.Equal(2, tableRows.Count);
            var firstRow = tableRows[0];
            var secondRow = tableRows[1];
            utils.CheckWordTableRowContents(firstRow, new List<string> { "Position", "", "", "", "", });
            utils.CheckWordTableRowContents(secondRow, new List<string> { "1", "2", "3", "4", "5", });

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

    #region <<Inline >>

    [Fact]
    public void DocumentFile_Inline()
    {
        lock (locker)
        {
            //Given
            var utils = new TestUtils();
            var parentFile = "Template-Inline.docx";
            var parentTemplate = new WvDocumentFileTemplate
            {
                Template = new TestUtils().LoadFileAsStream(parentFile)
            };
            var dataSource = SampleData;
            //When
            WvDocumentFileTemplateProcessResult? result = parentTemplate.Process(
                dataSource: dataSource,
                culture: null);
            //Then
            utils.GeneralResultChecks(result);
            var paragraphs = result.ResultItems[0].WordDocument.MainDocumentPart.Document.Body
                .ChildElements.Where(x=> x.GetType().FullName == typeof(Word.Paragraph).FullName);
            Assert.NotNull(paragraphs);
            Assert.Equal(29, paragraphs.Count());
            utils.SaveFileFromStream(result!.ResultItems[0]!.Result!, parentFile);
        }
    }

    [Fact]
    public void DocumentFile_Inline1()
    {
        lock (locker)
        {
            //Given
            var utils = new TestUtils();
            var parentFile = "Template-Inline-1.docx";
            var parentTemplate = new WvDocumentFileTemplate
            {
                Template = new TestUtils().LoadFileAsStream(parentFile)
            };
            var dataSource = SampleData;
            //When
            WvDocumentFileTemplateProcessResult? result = parentTemplate.Process(
                dataSource: dataSource,
                culture: null);
            //Then
            utils.GeneralResultChecks(result);
            var paragraphs = result.ResultItems[0].WordDocument.MainDocumentPart.Document.Body
                .ChildElements.Where(x=> x.GetType().FullName == typeof(Word.Paragraph).FullName);
            Assert.NotNull(paragraphs);
            Assert.Empty(paragraphs);
            utils.SaveFileFromStream(result!.ResultItems[0]!.Result!, parentFile);
        }
    }

    [Fact]
    public void DocumentFile_Inline2()
    {
        lock (locker)
        {
            //Given
            var utils = new TestUtils();
            var parentFile = "Template-Inline-2.docx";
            var parentTemplate = new WvDocumentFileTemplate
            {
                Template = new TestUtils().LoadFileAsStream(parentFile)
            };
            var dataSource = SampleData;
            //When
            WvDocumentFileTemplateProcessResult? result = parentTemplate.Process(
                dataSource: dataSource,
                culture: null);
            //Then
            utils.GeneralResultChecks(result);
            var paragraphs = result.ResultItems[0].WordDocument.MainDocumentPart.Document.Body
                .ChildElements.Where(x=> x.GetType().FullName == typeof(Word.Paragraph).FullName);
            Assert.NotNull(paragraphs);
            Assert.Equal(7, paragraphs.Count());
            utils.SaveFileFromStream(result!.ResultItems[0]!.Result!, parentFile);
        }
    }

    [Fact]
    public void DocumentFile_Inline3()
    {
        lock (locker)
        {
            //Given
            var utils = new TestUtils();
            var parentFile = "Template-Inline-3.docx";
            var parentTemplate = new WvDocumentFileTemplate
            {
                Template = new TestUtils().LoadFileAsStream(parentFile)
            };
            var dataSource = SampleData;
            //When
            WvDocumentFileTemplateProcessResult? result = parentTemplate.Process(
                dataSource: dataSource,
                culture: null);
            //Then
            utils.GeneralResultChecks(result);
            var paragraphs = result.ResultItems[0].WordDocument.MainDocumentPart.Document.Body
                .ChildElements.Where(x=> x.GetType().FullName == typeof(Word.Paragraph).FullName);
            Assert.NotNull(paragraphs);
            Assert.Equal(3, paragraphs.Count());
            utils.SaveFileFromStream(result!.ResultItems[0]!.Result!, parentFile);
        }
    }
    
    [Fact]
    public void DocumentFile_Inline4()
    {
        lock (locker)
        {
            //Given
            var utils = new TestUtils();
            var parentFile = "Template-Inline-4.docx";
            var parentTemplate = new WvDocumentFileTemplate
            {
                Template = new TestUtils().LoadFileAsStream(parentFile)
            };
            var dataSource = SampleData;
            //When
            WvDocumentFileTemplateProcessResult? result = parentTemplate.Process(
                dataSource: dataSource,
                culture: null);
            //Then
            utils.GeneralResultChecks(result);
            var paragraphs = result.ResultItems[0].WordDocument.MainDocumentPart.Document.Body
                .ChildElements.Where(x=> x.GetType().FullName == typeof(Word.Paragraph).FullName);
            var tables = result.ResultItems[0].WordDocument.MainDocumentPart.Document.Body
                .ChildElements.Where(x=> x.GetType().FullName == typeof(Word.Table).FullName);   
            Assert.NotNull(paragraphs);
            Assert.Equal(2, paragraphs.Count());
            Assert.Equal(2, tables.Count());
            utils.SaveFileFromStream(result!.ResultItems[0]!.Result!, parentFile);
        }
    }    
    
    [Fact]
    public void DocumentFile_Inline5()
    {
        //Inline does not work in tables
        lock (locker)
        {
            //Given
            var utils = new TestUtils();
            var parentFile = "Template-Inline-5.docx";
            var parentTemplate = new WvDocumentFileTemplate
            {
                Template = new TestUtils().LoadFileAsStream(parentFile)
            };
            var dataSource = SampleData;
            //When
            WvDocumentFileTemplateProcessResult? result = parentTemplate.Process(
                dataSource: dataSource,
                culture: null);
            //Then
            utils.GeneralResultChecks(result);
            var paragraphs = result.ResultItems[0].WordDocument!.MainDocumentPart!.Document!.Body!
                .ChildElements.Where(x=> x.GetType().FullName == typeof(Word.Paragraph).FullName).ToList();
            var tables = result.ResultItems[0].WordDocument!.MainDocumentPart!.Document!.Body!
                .ChildElements.Where(x=> x.GetType().FullName == typeof(Word.Table).FullName).ToList();             
            Assert.NotNull(paragraphs);
            Assert.NotNull(tables);
            Assert.Single(paragraphs);
            Assert.Single(tables);
            var table = (Word.Table)tables.First();
            var tableRows = table
                .ChildElements.Where(x=> x.GetType().FullName == typeof(Word.TableRow).FullName).ToList();               
            Assert.Equal(5,tableRows.Count);
            utils.SaveFileFromStream(result!.ResultItems[0]!.Result!, parentFile);
        }
    }        

    #endregion

    #region <<Edge >>

    [Fact]
    public void DocumentFile_Images_Header_And_Footer()
    {
        //Inline does not work in tables
        lock (locker)
        {
            //Given
            var utils = new TestUtils();
            var parentFile = "Template-Header-Footer-Images.docx";
            var parentTemplate = new WvDocumentFileTemplate
            {
                Template = new TestUtils().LoadFileAsStream(parentFile)
            };
            var dataSource = SampleData;
            //When
            WvDocumentFileTemplateProcessResult? result = parentTemplate.Process(
                dataSource: dataSource,
                culture: null);
            //Then
            utils.GeneralResultChecks(result);
          
            utils.SaveFileFromStream(result!.ResultItems[0]!.Result!, parentFile);
        }
    }         

    #endregion
    [Fact]
    public void DocumentFile_Test()
    {
        //Inline does not work in tables
        lock (locker)
        {
            //Given
            var utils = new TestUtils();
            var parentFile = "test1.docx";
            var parentTemplate = new WvDocumentFileTemplate
            {
                Template = new TestUtils().LoadFileAsStream(parentFile)
            };
            var dataSource = SampleData;
            //When
            WvDocumentFileTemplateProcessResult? result = parentTemplate.Process(
                dataSource: dataSource,
                culture: null);
            //Then
            utils.GeneralResultChecks(result);
          
            utils.SaveFileFromStream(result!.ResultItems[0]!.Result!, parentFile);
        }
    }          
}
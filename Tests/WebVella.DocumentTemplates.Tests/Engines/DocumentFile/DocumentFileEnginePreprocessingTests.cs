using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Engines.DocumentFile;
using WebVella.DocumentTemplates.Engines.DocumentFile.Utility;
using WebVella.DocumentTemplates.Tests.Models;
using WebVella.DocumentTemplates.Tests.Utils;
using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace WebVella.DocumentTemplates.Tests.Engines;

public class DocumentFileEnginePreprocessingTests : TestBase
{
    private static readonly object locker = new();

    public DocumentFileEnginePreprocessingTests() : base()
    {
    }

    #region << Paragraph >>

    [Fact]
    public void DocumentFile_Preprocessing1()
    {
        lock (locker)
        {
            //Given
            var utils = new TestUtils();
            var templateFile = "Template-Preprocessing-1.docx";
            var template = new WvDocumentFileTemplate
            {
                Template = new TestUtils().LoadFileAsStream(templateFile)
            };

            var templateErrors = template.Validate();
            Assert.Empty(templateErrors);

            var original = WordprocessingDocument.Open(template.Template!, true);
            //When
            WvDocumentFileEngineUtility.IsolateTemplateTags(original);
            original.Save();
            
            //Then
            var paragraphs = original.MainDocumentPart!.Document!.Body!
                .ChildElements.Where(x => x.GetType().FullName == typeof(Word.Paragraph).FullName).ToList();

            Assert.Single(paragraphs);
            var paragraph = paragraphs[0];
            var textElements = paragraph.Descendants<Word.Text>().ToList();
            Assert.Equal(6, textElements.Count);
            Assert.Equal("{{name(S=’,’)}}",textElements[0].InnerText);
            Assert.Equal("{{test}}",textElements[1].InnerText);
            Assert.Equal("{{<#}}",textElements[2].InnerText);
            Assert.Equal("tes",textElements[3].InnerText);
            Assert.Equal("t",textElements[4].InnerText);
            Assert.Equal("{{#>}}",textElements[5].InnerText);
            
            OpenXmlValidator validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
            var resultErrors =  validator.Validate(original).ToList();            
            Assert.Empty(resultErrors);
        }
    }

    [Fact]
    public void DocumentFile_Preprocessing2()
    {
        lock (locker)
        {
            //Given
            var utils = new TestUtils();
            var templateFile = "Template-Preprocessing-2.docx";
            var template = new WvDocumentFileTemplate
            {
                Template = new TestUtils().LoadFileAsStream(templateFile)
            };

            var templateErrors = template.Validate();
            Assert.Empty(templateErrors);

            var original = WordprocessingDocument.Open(template.Template!, true);
            //When
            WvDocumentFileEngineUtility.IsolateTemplateTags(original);
            original.Save();
            
            //Then
            var paragraphs = original.MainDocumentPart!.Document!.Body!
                .ChildElements.Where(x => x.GetType().FullName == typeof(Word.Paragraph).FullName).ToList();

            Assert.Equal(3,paragraphs.Count);
            var par1Texts = paragraphs[0].Descendants<Word.Text>().ToList();
            var par2Texts = paragraphs[1].Descendants<Word.Text>().ToList();
            var par3Texts = paragraphs[2].Descendants<Word.Text>().ToList();

            Assert.Single(par1Texts);
            Assert.Equal("{{<#}}",par1Texts[0].InnerText);
            
            Assert.Equal(4,par2Texts.Count);
            Assert.Equal("{{name(S=’,’)}}",par2Texts[0].InnerText);
            Assert.Equal("{{test}}",par2Texts[1].InnerText);
            Assert.Equal("tes",par2Texts[2].InnerText);        
            Assert.Equal("t",par2Texts[3].InnerText);        
            
            Assert.Single(par3Texts);
            Assert.Equal("{{#>}}",par3Texts[0].InnerText);            
            
            OpenXmlValidator validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
            var resultErrors =  validator.Validate(original).ToList();            
            Assert.Empty(resultErrors);            
        }
    }    
    
    [Fact]
    public void DocumentFile_Preprocessing3()
    {
        lock (locker)
        {
            //Given
            var utils = new TestUtils();
            var templateFile = "Template-Preprocessing-3.docx";
            var template = new WvDocumentFileTemplate
            {
                Template = new TestUtils().LoadFileAsStream(templateFile)
            };

            var templateErrors = template.Validate();
            Assert.Empty(templateErrors);

            var original = WordprocessingDocument.Open(template.Template!, true);
            //When
            WvDocumentFileEngineUtility.IsolateTemplateTags(original);
            original.Save();
            
            //Then
            var paragraphs = original.MainDocumentPart!.Document!.Body!
                .ChildElements.Where(x => x.GetType().FullName == typeof(Word.Paragraph).FullName).ToList();

            Assert.Single(paragraphs);
            var paragraph = paragraphs[0];
            var textElements = paragraph.Descendants<Word.Text>().ToList();
            Assert.Equal(3, textElements.Count);
            Assert.Equal("{{<#contact.[0]}}",textElements[0].InnerText);
            Assert.Equal("{{first_name[1]}}",textElements[1].InnerText);
            Assert.Equal("{{#>}}",textElements[2].InnerText);
            
            OpenXmlValidator validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
            var resultErrors =  validator.Validate(original).ToList();            
            Assert.Empty(resultErrors);
        }
    }    

    #endregion



}
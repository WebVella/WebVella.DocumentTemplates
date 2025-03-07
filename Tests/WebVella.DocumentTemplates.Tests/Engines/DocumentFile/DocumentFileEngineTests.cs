using WebVella.DocumentTemplates.Engines.DocumentFile;
using WebVella.DocumentTemplates.Tests.Models;
using WebVella.DocumentTemplates.Tests.Utils;

namespace WebVella.DocumentTemplates.Tests.Engines;
public class DocumentFileEngineTests : TestBase
{
	private static readonly object locker = new();
	public DocumentFileEngineTests() : base() { }

	#region << Paragraph >>
	[Fact]
	public void DocumentFile_Paragraph()
	{
		lock (locker)
		{
			//Given
			var templateFile = "Template-Paragraph1.docx";
			var template = new WvDocumentFileTemplate
			{
				Template = new TestUtils().LoadFileAsStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvDocumentFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			new TestUtils().SaveFileFromStream(result!.ResultItems[0]!.Result!,templateFile);
		}
	}

	[Fact]
	public void DocumentFile_Paragraph_2()
	{
		lock (locker)
		{
			//Given
			var templateFile = "Template-Paragraph2.docx";
			var template = new WvDocumentFileTemplate
			{
				Template = new TestUtils().LoadFileAsStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvDocumentFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			new TestUtils().SaveFileFromStream(result!.ResultItems[0]!.Result!,templateFile);
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
			var templateFile = "Template-Table1.docx";
			var template = new WvDocumentFileTemplate
			{
				Template = new TestUtils().LoadFileAsStream(templateFile)
			};
			var dataSource = SampleData;
			//When
			WvDocumentFileTemplateProcessResult? result = template.Process(dataSource);
			//Then
			new TestUtils().SaveFileFromStream(result!.ResultItems[0]!.Result!,templateFile);
		}
	}
	#endregion
}

using WebVella.DocumentTemplates.Engines.DocumentFile;
using WebVella.DocumentTemplates.Tests.Models;
using WebVella.DocumentTemplates.Tests.Utils;

namespace WebVella.DocumentTemplates.Tests.Engines;
public class DocumentFileEngineTests : TestBase
{
	private static readonly object locker = new();
	public DocumentFileEngineTests() : base() { }

	#region << General >>
	[Fact]
	public void DocumentFile_General()
	{
		lock (locker)
		{
			//Given
			var templateFile = "Template1.docx";
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

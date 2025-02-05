using System.Data;
using System.Text;
using WebVella.DocumentTemplates.Engines.Text;
using WebVella.DocumentTemplates.Extensions;
using WebVella.DocumentTemplates.Tests.Models;

namespace WebVella.DocumentTemplates.Tests.Engines;
public class TextFileEngineTests : TestBase
{
	private static readonly object locker = new();
	public TextFileEngineTests() : base() { }

	[Fact]
	public void Template1_Repeat()
	{
		var templateFile = "Template1.txt";
		var resultFile = "Template1Result.txt";
		var expectedResultText = Encoding.UTF8.GetString(LoadFile(resultFile)).RemoveZeroBitSpaceCharacters();
		lock (locker)
		{
			var template = new WvTextFileTemplate{ 
				Template = LoadFile(templateFile)
			};
			var result = template.Process(SampleData, encoding:Encoding.UTF8);
			Assert.NotNull(result);
			Assert.NotNull(result.Template);
			Assert.NotNull(result.Result);
			var resultString = Encoding.UTF8.GetString(result.Result);
			Assert.Equal(expectedResultText, resultString);
		}
	}

	[Fact]
	public void Template2_Repeat()
	{
		var templateFile = "Template2.txt";
		var resultFile = "Template2Result.txt";
		var expectedResultText = Encoding.UTF8.GetString(LoadFile(resultFile)).RemoveZeroBitSpaceCharacters();
		lock (locker)
		{
			var template = new WvTextFileTemplate{ 
				Template = LoadFile(templateFile)
			};
			var result = template.Process(SampleData, encoding:Encoding.UTF8);
			Assert.NotNull(result);
			Assert.NotNull(result.Template);
			Assert.NotNull(result.Result);
			var resultString = Encoding.UTF8.GetString(result.Result);
			Assert.Equal(expectedResultText, resultString);
		}
	}

	[Fact]
	public void Template3_SingleRow()
	{
		var templateFile = "Template3.txt";
		var resultFile = "Template3Result.txt";
		var expectedResultText = Encoding.UTF8.GetString(LoadFile(resultFile)).RemoveZeroBitSpaceCharacters();
		lock (locker)
		{
			var template = new WvTextFileTemplate{ 
				Template = LoadFile(templateFile)
			};
			var result = template.Process(SampleData, encoding:Encoding.UTF8);
			Assert.NotNull(result);
			Assert.NotNull(result.Template);
			Assert.NotNull(result.Result);
			var resultString = Encoding.UTF8.GetString(result.Result);
			Assert.Equal(expectedResultText, resultString);
		}
	}

	[Fact]
	public void Template4_AllIndexed()
	{
		var templateFile = "Template4.txt";
		var resultFile = "Template4Result.txt";
		var expectedResultText = Encoding.UTF8.GetString(LoadFile(resultFile)).RemoveZeroBitSpaceCharacters();
		lock (locker)
		{
			var template = new WvTextFileTemplate{ 
				Template = LoadFile(templateFile)
			};
			var result = template.Process(SampleData, encoding:Encoding.UTF8);
			Assert.NotNull(result);
			Assert.NotNull(result.Template);
			Assert.NotNull(result.Result);
			var resultString = Encoding.UTF8.GetString(result.Result);
			Assert.Equal(expectedResultText, resultString);
		}
	}

	[Fact]
	public void Template5_AllIndexedWithNotPresentIndexShouldNotFail()
	{
		var templateFile = "Template5.txt";
		var resultFile = "Template5Result.txt";
		var expectedResultText = Encoding.UTF8.GetString(LoadFile(resultFile)).RemoveZeroBitSpaceCharacters();
		lock (locker)
		{
			var template = new WvTextFileTemplate{ 
				Template = LoadFile(templateFile)
			};
			var result = template.Process(SampleData, encoding:Encoding.UTF8);
			Assert.NotNull(result);
			Assert.NotNull(result.Template);
			Assert.NotNull(result.Result);
			var resultString = Encoding.UTF8.GetString(result.Result);
			Assert.Equal(expectedResultText, resultString);
		}
	}

	[Fact]
	public void Template6_AllIndexedWithNotPresentIndexShouldNotFail()
	{
		var templateFile = "Template6.txt";
		var resultFile = "Template6Result.txt";
		var expectedResultText = Encoding.UTF8.GetString(LoadFile(resultFile)).RemoveZeroBitSpaceCharacters();
		lock (locker)
		{
			var template = new WvTextFileTemplate{ 
				Template = LoadFile(templateFile)
			};
			var result = template.Process(SampleData, encoding:Encoding.UTF8);
			Assert.NotNull(result);
			Assert.NotNull(result.Template);
			Assert.NotNull(result.Result);
			var resultString = Encoding.UTF8.GetString(result.Result);
			Assert.Equal(expectedResultText, resultString);
		}
	}
}

using System.Text;
using WebVella.DocumentTemplates.Engines.TextFile;
using WebVella.DocumentTemplates.Extensions;
using WebVella.DocumentTemplates.Tests.Models;
using WebVella.DocumentTemplates.Tests.Utils;

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
		var expectedResultText = Encoding.UTF8.GetString( new TestUtils().LoadFile(resultFile)).RemoveZeroBitSpaceCharacters();
		lock (locker)
		{
			var template = new WvTextFileTemplate
			{
				Template =  new TestUtils().LoadFileStream(templateFile)
			};
			var result = template.Process(SampleData, encoding: Encoding.UTF8);
			Assert.NotNull(result);
			Assert.NotNull(result.Template);
			Assert.NotNull(result.ResultItems);
			Assert.Single(result.ResultItems);
			Assert.NotNull(result.ResultItems[0].Result);
			var resultString = Encoding.UTF8.GetString(result.ResultItems[0].Result?.ToArray() ?? new byte[0]);
			Assert.Equal(expectedResultText, resultString);
		}
	}

	[Fact]
	public void Template2_Repeat()
	{
		var templateFile = "Template2.txt";
		var resultFile = "Template2Result.txt";
		var expectedResultText = Encoding.UTF8.GetString( new TestUtils().LoadFile(resultFile)).RemoveZeroBitSpaceCharacters();
		lock (locker)
		{
			var template = new WvTextFileTemplate
			{
				Template =  new TestUtils().LoadFileStream(templateFile)
			};
			var result = template.Process(SampleData, encoding: Encoding.UTF8);
			Assert.NotNull(result);
			Assert.NotNull(result.Template);
			Assert.NotNull(result.ResultItems);
			Assert.Single(result.ResultItems);
			Assert.NotNull(result.ResultItems[0].Result);
			var resultString = Encoding.UTF8.GetString(result.ResultItems[0].Result?.ToArray() ?? new byte[0]);
			Assert.Equal(expectedResultText, resultString);
		}
	}

	[Fact]
	public void Template3_SingleRow()
	{
		var templateFile = "Template3.txt";
		var resultFile = "Template3Result.txt";
		var expectedResultText = Encoding.UTF8.GetString( new TestUtils().LoadFile(resultFile)).RemoveZeroBitSpaceCharacters();
		lock (locker)
		{
			var template = new WvTextFileTemplate
			{
				Template =  new TestUtils().LoadFileStream(templateFile)
			};
			var result = template.Process(SampleData, encoding: Encoding.UTF8);
			Assert.NotNull(result);
			Assert.NotNull(result.Template);
			Assert.NotNull(result.ResultItems);
			Assert.Single(result.ResultItems);
			Assert.NotNull(result.ResultItems[0].Result);
			var resultString = Encoding.UTF8.GetString(result.ResultItems[0].Result?.ToArray() ?? new byte[0]);
			Assert.Equal(expectedResultText, resultString);
		}
	}

	[Fact]
	public void Template4_AllIndexed()
	{
		var templateFile = "Template4.txt";
		var resultFile = "Template4Result.txt";
		var expectedResultText = Encoding.UTF8.GetString( new TestUtils().LoadFile(resultFile)).RemoveZeroBitSpaceCharacters();
		lock (locker)
		{
			var template = new WvTextFileTemplate
			{
				Template =  new TestUtils().LoadFileStream(templateFile)
			};
			var result = template.Process(SampleData, encoding: Encoding.UTF8);
			Assert.NotNull(result);
			Assert.NotNull(result.Template);
			Assert.NotNull(result.ResultItems);
			Assert.Single(result.ResultItems);
			Assert.NotNull(result.ResultItems[0].Result);
			var resultString = Encoding.UTF8.GetString(result.ResultItems[0].Result?.ToArray() ?? new byte[0]);
			Assert.Equal(expectedResultText, resultString);
		}
	}

	[Fact]
	public void Template5_AllIndexedWithNotPresentIndexShouldNotFail()
	{
		var templateFile = "Template5.txt";
		var resultFile = "Template5Result.txt";
		var expectedResultText = Encoding.UTF8.GetString( new TestUtils().LoadFile(resultFile)).RemoveZeroBitSpaceCharacters();
		lock (locker)
		{
			var template = new WvTextFileTemplate
			{
				Template =  new TestUtils().LoadFileStream(templateFile)
			};
			var result = template.Process(SampleData, encoding: Encoding.UTF8);
			Assert.NotNull(result);
			Assert.NotNull(result.Template);
			Assert.NotNull(result.ResultItems);
			Assert.Single(result.ResultItems);
			Assert.NotNull(result.ResultItems[0].Result);
			var resultString = Encoding.UTF8.GetString(result.ResultItems[0].Result?.ToArray() ?? new byte[0]);
			Assert.Equal(expectedResultText, resultString);
		}
	}

	[Fact]
	public void Template6_AllIndexedWithNotPresentIndexShouldNotFail()
	{
		var templateFile = "Template6.txt";
		var resultFile = "Template6Result.txt";
		var expectedResultText = Encoding.UTF8.GetString( new TestUtils().LoadFile(resultFile)).RemoveZeroBitSpaceCharacters();
		lock (locker)
		{
			var template = new WvTextFileTemplate
			{
				Template =  new TestUtils().LoadFileStream(templateFile)
			};
			var result = template.Process(SampleData, encoding: Encoding.UTF8);
			Assert.NotNull(result);
			Assert.NotNull(result.Template);
			Assert.NotNull(result.ResultItems);
			Assert.Single(result.ResultItems);
			Assert.NotNull(result.ResultItems[0].Result);
			var resultString = Encoding.UTF8.GetString(result.ResultItems[0].Result?.ToArray() ?? new byte[0]);
			Assert.Equal(expectedResultText, resultString);
		}
	}

	[Fact]
	public void Template1_GroupBy()
	{
		var templateFile = "Template1.txt";

		lock (locker)
		{
			var template = new WvTextFileTemplate
			{
				Template =  new TestUtils().LoadFileStream(templateFile),
				GroupDataByColumns = new List<string> { "sku" }
			};
			var data = SampleData.CreateAsNew();
			data.Rows[1]["sku"] = data.Rows[0]["sku"];
			var result = template.Process(data, encoding: Encoding.UTF8);
			Assert.NotNull(result);
			Assert.NotNull(result.Template);
			Assert.NotNull(result.ResultItems);
			Assert.Equal(4, result.ResultItems.Count);
			Assert.NotNull(result.ResultItems[0].Result);
			var resultString = Encoding.UTF8.GetString(result.ResultItems[0].Result?.ToArray() ?? new byte[0]);
			Assert.Equal($"1{Environment.NewLine}2{Environment.NewLine}item1{Environment.NewLine}item2{Environment.NewLine}", resultString);
		}
	}

	[Fact]
	public void Error_File()
	{
		lock (locker)
		{
			//Given
			var templateFile = "Template1.xlsx";
			var template = new WvTextFileTemplate
			{
				Template =  new TestUtils().LoadFileStream(templateFile)
			};
			var dataSource = SampleData;
			WvTextFileTemplateProcessResult? result = null;
			//When
			var action = () => result = template.Process(dataSource);
			var ex = Record.Exception(action);
			Assert.Null(ex); //not a good way to check if it is plaintext
		}
	}
	
	[Fact]
	public void Inline1_Test()
	{
		var templateFile = "Inline1.txt";

		lock (locker)
		{
			var template = new WvTextFileTemplate
			{
				Template =  new TestUtils().LoadFileStream(templateFile)
			};
			var result = template.Process(SampleData, encoding: Encoding.UTF8);
			Assert.NotNull(result);
			Assert.NotNull(result.Template);
			Assert.NotNull(result.ResultItems);
			Assert.Single(result.ResultItems);
			Assert.NotNull(result.ResultItems[0].Result);
			var resultString = Encoding.UTF8.GetString(result.ResultItems[0].Result?.ToArray() ?? new byte[0]);
			Assert.Equal($"1 item12 item2", resultString);
			new TestUtils().SaveFileFromStream(result!.ResultItems[0]!.Result!, templateFile);
		}
	}	
	[Fact]
	public void Inline2_Test()
	{
		var templateFile = "Inline2.txt";

		lock (locker)
		{
			var template = new WvTextFileTemplate
			{
				Template =  new TestUtils().LoadFileStream(templateFile)
			};
			var result = template.Process(SampleData, encoding: Encoding.UTF8);
            Assert.NotNull(result);
			Assert.NotNull(result.Template);
			Assert.NotNull(result.ResultItems);
			// Assert.Equal(4, result.ResultItems.Count);
			// Assert.NotNull(result.ResultItems[0].Result);
			var resultString = Encoding.UTF8.GetString(result.ResultItems[0].Result?.ToArray() ?? new byte[0]);
			Assert.Equal($"1 item1{Environment.NewLine}2 item2{Environment.NewLine}", resultString);
			new TestUtils().SaveFileFromStream(result!.ResultItems[0]!.Result!, templateFile);
		}
	}
    [Fact]
    public void Inline3_Test()
    {
        var templateFile = "Inline3.txt";

        lock (locker)
        {
            var template = new WvTextFileTemplate
            {
                Template = new TestUtils().LoadFileStream(templateFile)
            };
            var result = template.Process(SampleData, encoding: Encoding.UTF8);
            Assert.NotNull(result);
            Assert.NotNull(result.Template);
            Assert.NotNull(result.ResultItems);
            Assert.Single(result.ResultItems);
            Assert.NotNull(result.ResultItems[0].Result);
            var resultString = Encoding.UTF8.GetString(result.ResultItems[0].Result?.ToArray() ?? new byte[0]);
            Assert.Equal($"first_name00", resultString);
            new TestUtils().SaveFileFromStream(result!.ResultItems[0]!.Result!, templateFile);
        }
    }
    [Fact]
    public void Inline4_Test()
    {
        var templateFile = "Inline4.txt";

        lock (locker)
        {
            var template = new WvTextFileTemplate
            {
                Template = new TestUtils().LoadFileStream(templateFile)
            };
            var result = template.Process(SampleData, encoding: Encoding.UTF8);
            Assert.NotNull(result);
            Assert.NotNull(result.Template);
            Assert.NotNull(result.ResultItems);
            Assert.Single(result.ResultItems);
            Assert.NotNull(result.ResultItems[0].Result);
            var resultString = Encoding.UTF8.GetString(result.ResultItems[0].Result?.ToArray() ?? new byte[0]);
            Assert.Equal($"first_name00", resultString);
            new TestUtils().SaveFileFromStream(result!.ResultItems[0]!.Result!, templateFile);
        }
    }
}

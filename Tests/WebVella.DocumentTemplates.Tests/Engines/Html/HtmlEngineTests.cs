using System.Data;
using System.Diagnostics;
using System.Text.Json;
using WebVella.DocumentTemplates.Engines.Html;
using WebVella.DocumentTemplates.Extensions;
using WebVella.DocumentTemplates.Tests.Models;

namespace WebVella.DocumentTemplates.Tests.Engines;
public class HtmlEngineTests : TestBase
{
	public HtmlEngineTests() : base() { }

	#region << Arguments >>
	[Fact]
	public void ShouldHaveDataSource()
	{
		//Given
		var template = new WvHtmlTemplate();
		//When
		WvHtmlTemplateProcessResult? result = null;
		var action = () => result = template.Process(null, DefaultCulture);
		//Then
		var ex = Record.Exception(action);
		Assert.NotNull(ex);
		Assert.IsType<ArgumentException>(ex);
		var argEx = (ArgumentException)ex;
		Assert.Equal("dataSource", argEx.ParamName);
		Assert.StartsWith("No datasource provided!", argEx.Message);
	}
	[Fact]
	public void ShouldProcessEmpty()
	{
		//Given
		var templateString = String.Empty;
		var template = new WvHtmlTemplate() { Template = templateString };
		var ds = new DataTable();
		//When
		WvHtmlTemplateProcessResult? result = template.Process(ds, DefaultCulture);
		//Then
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.Equal(templateString, result.ResultItems[0].Result);

	}
	#endregion
	#region << Plain text >>
	[Fact]
	public void Text_NoTag()
	{
		var template = new WvHtmlTemplate()
		{
			Template = "test"
		};
		var data = SampleData.CreateAsNew(new List<int> { 1 });
		WvHtmlTemplateProcessResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		Assert.Equal(template.Template, result.ResultItems[0].Result);
	}
	[Fact]
	public void Text_NoTag2()
	{
		var template = new WvHtmlTemplate()
		{
			Template = "test<br>test2"
		};
		var data = SampleData.CreateAsNew(new List<int> { 1 });
		WvHtmlTemplateProcessResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		Assert.Equal(template.Template, result.ResultItems[0].Result);
	}

	[Fact]
	public void Text_NoTag3()
	{
		var template = new WvHtmlTemplate()
		{
			Template = "test<br>{{name[0]}}"
		};
		var data = SampleData;
		WvHtmlTemplateProcessResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		Assert.Equal("test<br>item1", result.ResultItems[0].Result);
	}

	[Fact]
	public void Text_Tag8()
	{
		var template = new WvHtmlTemplate()
		{
			Template = "<p>Component: {{sku(F=H,S=', ')}} with ETA: <strong>{{name(F=H,S=', ')}}</strong></p>"
		};
		var data = SampleData.CreateAsNew(new List<int> { 0, 1, 2 });
		WvHtmlTemplateProcessResult? result = template.Process(data);
		Debug.WriteLine(JsonSerializer.Serialize(result.ResultItems[0]));
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		Assert.Equal("<p>Component: sku1, sku2, sku3 with ETA: <strong>item1, item2, item3</strong></p>", result.ResultItems[0].Result);
	}

	#endregion

	#region << Plan div >>
	[Fact]
	public void Text_Div1()
	{
		var template = new WvHtmlTemplate()
		{
			Template = "<div>test</div><div>{{name}}</div>"
		};
		var data = SampleData.CreateAsNew(new List<int> { 0, 1 });
		WvHtmlTemplateProcessResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		Assert.Equal("<div>test</div><div>item1</div><div>item2</div>", result.ResultItems[0].Result);
	}

	[Fact]
	public void Text_Div2()
	{
		var template = new WvHtmlTemplate()
		{
			Template = "<div>test</div><div><div>{{name}}</div></div>"
		};
		var data = SampleData.CreateAsNew(new List<int> { 0, 1 });
		WvHtmlTemplateProcessResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		Assert.Equal("<div>test</div><div><div>item1</div></div><div><div>item2</div></div>", result.ResultItems[0].Result);
	}
	#endregion

	#region << Plain Paragraph >>

	[Fact]
	public void Text_Paragraph1()
	{
		var template = new WvHtmlTemplate()
		{
			Template = "<p>test</p><p>{{name}}</p>"
		};
		var data = SampleData.CreateAsNew(new List<int> { 0, 1 });
		WvHtmlTemplateProcessResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		Assert.Equal("<p>test</p><p>item1</p><p>item2</p>", result.ResultItems[0].Result);
	}
	#endregion
}

using System.Data;
using WebVella.DocumentTemplates.Engines.Html;
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
		WvHtmlTemplateResult? result = null;
#pragma warning disable CS8625 // Cannot convert null literal to non-nullable reference type.
		var action = () => result = template.Process(null, DefaultCulture);
#pragma warning restore CS8625 // Cannot convert null literal to non-nullable reference type.
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
		WvHtmlTemplateResult? result = template.Process(ds, DefaultCulture);
		//Then
		Assert.NotNull(result);
		Assert.Equal(templateString, result.Result);

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
		var data = CreateDataTableFromExisting(SampleData, new List<int> { 1 });
		WvHtmlTemplateResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.False(String.IsNullOrWhiteSpace(result.Result));
		Assert.Equal(template.Template, result.Result);
	}
	[Fact]
	public void Text_NoTag2()
	{
		var template = new WvHtmlTemplate()
		{
			Template = "test<br>test2"
		};
		var data = CreateDataTableFromExisting(SampleData, new List<int> { 1 });
		WvHtmlTemplateResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.False(String.IsNullOrWhiteSpace(result.Result));
		Assert.Equal(template.Template, result.Result);
	}

	[Fact]
	public void Text_NoTag3()
	{
		var template = new WvHtmlTemplate()
		{
			Template = "test<br>{{name[0]}}"
		};
		var data = SampleData;
		WvHtmlTemplateResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.False(String.IsNullOrWhiteSpace(result.Result));
		Assert.Equal("test<br>item1", result.Result);
	}

	[Fact]
	public void Text_Tag8()
	{
		var template = new WvHtmlTemplate()
		{
			Template = "<p>Component: {{sku(F=H,S=', ')}} with ETA: <strong>{{name(F=H,S=', ')}}</strong></p>"
		};
		var data = CreateDataTableFromExisting(SampleData, new List<int> { 0,1,2 });
		WvHtmlTemplateResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.False(String.IsNullOrWhiteSpace(result.Result));
		Assert.Equal("<p>Component: sku1, sku2, sku3 with ETA: <strong>item1, item2, item3</strong></p>", result.Result);

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
		var data = CreateDataTableFromExisting(SampleData, new List<int> { 0,1 });
		WvHtmlTemplateResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.False(String.IsNullOrWhiteSpace(result.Result));
		Assert.Equal("<div>test</div><div>item1</div><div>item2</div>", result.Result);
	}

	[Fact]
	public void Text_Div2()
	{
		var template = new WvHtmlTemplate()
		{
			Template = "<div>test</div><div><div>{{name}}</div></div>"
		};
		var data = CreateDataTableFromExisting(SampleData, new List<int> { 0,1 });
		WvHtmlTemplateResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.False(String.IsNullOrWhiteSpace(result.Result));
		Assert.Equal("<div>test</div><div><div>item1</div></div><div><div>item2</div></div>", result.Result);
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
		var data = CreateDataTableFromExisting(SampleData, new List<int> { 0,1 });
		WvHtmlTemplateResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.False(String.IsNullOrWhiteSpace(result.Result));
		Assert.Equal("<p>test</p><p>item1</p><p>item2</p>", result.Result);
	}
	#endregion
}

using System.Data;
using WebVella.DocumentTemplates.Engines.Text;
using WebVella.DocumentTemplates.Tests.Models;

namespace WebVella.DocumentTemplates.Tests.Engines;
public class TextEngineTests : TestBase
{
	public TextEngineTests() : base() { }

	#region << Arguments >>
	[Fact]
	public void ShouldHaveDataSource()
	{
		//Given
		var template = new WvTextTemplate();
		//When
		WvTextTemplateResult? result = null;
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
		var template = new WvTextTemplate() { Template = templateString };
		var ds = new DataTable();
		//When
		WvTextTemplateResult? result = template.Process(ds, DefaultCulture);
		//Then
		Assert.NotNull(result);
		Assert.Equal(templateString, result.Result);

	}
	#endregion

	#region << Plain text >>
	[Fact]
	public void Text_NoTag()
	{
		var template = new WvTextTemplate()
		{
			Template = "test"
		};
		var data = CreateDataTableFromExisting(SampleData, new List<int> { 1 });
		WvTextTemplateResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.False(String.IsNullOrWhiteSpace(result.Result));
		var lines = GetLines(result.Result);
		Assert.Single(lines);
		Assert.Equal(template.Template, result.Result);
	}

	[Fact]
	public void Text_NoTag2()
	{
		var template = new WvTextTemplate()
		{
			Template = "test test2 is test {{"
		};
		var data = CreateDataTableFromExisting(SampleData, new List<int> { 1 });
		WvTextTemplateResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.False(String.IsNullOrWhiteSpace(result.Result));
		var lines = GetLines(result.Result);
		Assert.Single(lines);
		Assert.Equal(template.Template, result.Result);
	}

	[Fact]
	public void Text_Tag()
	{
		var template = new WvTextTemplate()
		{
			Template = "{{sku}}{{name}}"
		};
		var data = CreateDataTableFromExisting(SampleData, new List<int> { 1 });
		WvTextTemplateResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.False(String.IsNullOrWhiteSpace(result.Result));
		var lines = GetLines(result.Result);
		Assert.Single(lines);
		Assert.Equal($"{data.Rows[0]["sku"]}{data.Rows[0]["name"]}", result.Result);
	}

	[Fact]
	public void Text_Tag2()
	{
		var template = new WvTextTemplate()
		{
			Template = "{{sku}} test {{name}}"
		};
		var data = CreateDataTableFromExisting(SampleData, new List<int> { 1 });
		WvTextTemplateResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.False(String.IsNullOrWhiteSpace(result.Result));
		var lines = GetLines(result.Result);
		Assert.Single(lines);
		Assert.Equal($"{data.Rows[0]["sku"]} test {data.Rows[0]["name"]}", result.Result);
	}

	[Fact]
	public void Text_Tag3()
	{
		var template = new WvTextTemplate()
		{
			Template = "{{sku}} test {{name}}"
		};
		var data = CreateDataTableFromExisting(SampleData, new List<int> { 0, 1 });
		WvTextTemplateResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.False(String.IsNullOrWhiteSpace(result.Result));
		var lines = GetLines(result.Result);
		Assert.Single(lines);
		Assert.Equal($"{data.Rows[0]["sku"]} test {data.Rows[0]["name"]}{data.Rows[1]["sku"]} test {data.Rows[1]["name"]}", result.Result);
	}

	[Fact]
	public void Text_Tag4()
	{
		var template = new WvTextTemplate()
		{
			Template = "{{sku}} test {{name}}" + Environment.NewLine
		};
		var data = CreateDataTableFromExisting(SampleData, new List<int> { 0, 1 });
		WvTextTemplateResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.False(String.IsNullOrWhiteSpace(result.Result));
		var lines = GetLines(result.Result);
		Assert.Equal(2, lines.Count);
		Assert.Equal($"{data.Rows[0]["sku"]} test {data.Rows[0]["name"]}", lines[0]);
		Assert.Equal($"{data.Rows[1]["sku"]} test {data.Rows[1]["name"]}", lines[1]);
	}

	[Fact]
	public void Text_Tag5()
	{
		var template = new WvTextTemplate()
		{
			Template = "{{sku(F=H,S=',',B=\", \")}}"
		};
		var data = CreateDataTableFromExisting(SampleData, new List<int> { 0, 1, 2 });
		WvTextTemplateResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.False(String.IsNullOrWhiteSpace(result.Result));
		var lines = GetLines(result.Result);
		Assert.Single(lines);
		Assert.Equal($"{data.Rows[0]["sku"]},{data.Rows[1]["sku"]},{data.Rows[2]["sku"]}", lines[0]);
	}

	[Fact]
	public void Text_Tag6()
	{
		var template = new WvTextTemplate()
		{
			Template = "test {{sku(F=H,S=',')}} {{name(F=H,S=',')}}"
		};
		var data = CreateDataTableFromExisting(SampleData, new List<int> { 0, 1, 2 });
		WvTextTemplateResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.False(String.IsNullOrWhiteSpace(result.Result));
		var lines = GetLines(result.Result);
		Assert.Single(lines);
		Assert.Equal($"test {data.Rows[0]["sku"]},{data.Rows[1]["sku"]},{data.Rows[2]["sku"]} {data.Rows[0]["name"]},{data.Rows[1]["name"]},{data.Rows[2]["name"]}", lines[0]);
	}

	[Fact]
	public void Text_Tag7()
	{
		var template = new WvTextTemplate()
		{
			Template = "test {{sku(F=H,S=',')}} {{name(F=H,S=', ')}}"
		};
		var data = CreateDataTableFromExisting(SampleData, new List<int> { 0, 1, 2 });
		WvTextTemplateResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.False(String.IsNullOrWhiteSpace(result.Result));
		var lines = GetLines(result.Result);
		Assert.Single(lines);
		Assert.Equal($"test {data.Rows[0]["sku"]},{data.Rows[1]["sku"]},{data.Rows[2]["sku"]} {data.Rows[0]["name"]}, {data.Rows[1]["name"]}, {data.Rows[2]["name"]}", lines[0]);
	}

	[Fact]
	public void Text_Tag8()
	{
		var template = new WvTextTemplate()
		{
			Template = "Component: {{sku(F=H,S=', ')}} with ETA: {{name(F=H,S=', ')}}"
		};
		var data = CreateDataTableFromExisting(SampleData, new List<int> { 0, 1, 2 });
		WvTextTemplateResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.False(String.IsNullOrWhiteSpace(result.Result));
		var lines = GetLines(result.Result);
		Assert.Single(lines);
		Assert.Equal($"Component: {data.Rows[0]["sku"]}, {data.Rows[1]["sku"]}, {data.Rows[2]["sku"]} with ETA: {data.Rows[0]["name"]}, {data.Rows[1]["name"]}, {data.Rows[2]["name"]}", lines[0]);
	}
	#endregion
}

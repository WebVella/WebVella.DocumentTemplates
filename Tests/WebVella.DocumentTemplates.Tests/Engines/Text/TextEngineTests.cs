using System.Data;
using WebVella.DocumentTemplates.Engines.Text;
using WebVella.DocumentTemplates.Extensions;
using WebVella.DocumentTemplates.Tests.Models;
using WebVella.DocumentTemplates.Tests.Utils;

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
		WvTextTemplateProcessResult? result = null;
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
		var template = new WvTextTemplate() { Template = templateString };
		var ds = new DataTable();
		//When
		WvTextTemplateProcessResult? result = template.Process(ds, DefaultCulture);
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
		var template = new WvTextTemplate()
		{
			Template = "test"
		};
		var data = SampleData.CreateAsNew(new List<int> { 1 });
		WvTextTemplateProcessResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		var lines = new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
		Assert.Single(lines);
		Assert.Equal(template.Template, result.ResultItems[0].Result);
	}

	[Fact]
	public void Text_NoTag2()
	{
		var template = new WvTextTemplate()
		{
			Template = "test test2 is test {{"
		};
		var data = SampleData.CreateAsNew(new List<int> { 1 });
		WvTextTemplateProcessResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		var lines =  new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
		Assert.Single(lines);
		Assert.Equal(template.Template, result.ResultItems[0].Result);
	}

	[Fact]
	public void Text_Tag()
	{
		var template = new WvTextTemplate()
		{
			Template = "{{sku}}{{name}}"
		};
		var data = SampleData.CreateAsNew(new List<int> { 1 });
		WvTextTemplateProcessResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		var lines =  new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
		Assert.Single(lines);
		Assert.Equal($"{data.Rows[0]["sku"]}{data.Rows[0]["name"]}", result.ResultItems[0].Result);
	}
	[Fact]
	public void Text_Tag_OneIndex()
	{
		var template = new WvTextTemplate()
		{
			Template = "{{sku[0]}}"
		};
		WvTextTemplateProcessResult? result = template.Process(SampleData);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		var lines =  new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
		Assert.Single(lines);
		Assert.Equal($"sku1", result.ResultItems[0].Result);
	}	
	[Fact]
	public void Text_Tag_2Index()
	{
		var template = new WvTextTemplate()
		{
			Template = "{{sku[0,1]}}"
		};
		WvTextTemplateProcessResult? result = template.Process(SampleData);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		var lines =  new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
		Assert.Single(lines);
		Assert.Equal($"sku1sku2", result.ResultItems[0].Result);
	}		

	[Fact]
	public void Text_Tag2()
	{
		var template = new WvTextTemplate()
		{
			Template = "{{sku}} test {{name}}"
		};
		var data = SampleData.CreateAsNew(new List<int> { 1 });
		WvTextTemplateProcessResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		var lines =  new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
		Assert.Single(lines);
		Assert.Equal($"{data.Rows[0]["sku"]} test {data.Rows[0]["name"]}", result.ResultItems[0].Result);
	}

	[Fact]
	public void Text_Tag3()
	{
		var template = new WvTextTemplate()
		{
			Template = "{{sku}} test {{name}}"
		};
		var data = SampleData.CreateAsNew(new List<int> { 0, 1 });
		WvTextTemplateProcessResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		var lines =  new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
		Assert.Single(lines);
		Assert.Equal($"{data.Rows[0]["sku"]} test {data.Rows[0]["name"]}{data.Rows[1]["sku"]} test {data.Rows[1]["name"]}", result.ResultItems[0].Result);
	}

	[Fact]
	public void Text_Tag4()
	{
		var template = new WvTextTemplate()
		{
			Template = "{{sku}} test {{name}}" + Environment.NewLine
		};
		var data = SampleData.CreateAsNew(new List<int> { 0, 1 });
		WvTextTemplateProcessResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		var lines =  new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
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
		var data = SampleData.CreateAsNew(new List<int> { 0, 1, 2 });
		WvTextTemplateProcessResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		var lines =  new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
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
		var data = SampleData.CreateAsNew(new List<int> { 0, 1, 2 });
		WvTextTemplateProcessResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		var lines =  new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
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
		var data = SampleData.CreateAsNew(new List<int> { 0, 1, 2 });
		WvTextTemplateProcessResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		var lines =  new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
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
		var data = SampleData.CreateAsNew(new List<int> { 0, 1, 2 });
		WvTextTemplateProcessResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		var lines =  new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
		Assert.Single(lines);
		Assert.Equal($"Component: {data.Rows[0]["sku"]}, {data.Rows[1]["sku"]}, {data.Rows[2]["sku"]} with ETA: {data.Rows[0]["name"]}, {data.Rows[1]["name"]}, {data.Rows[2]["name"]}", lines[0]);
	}

	[Fact]
	public void Text_Tag9()
	{
		var template = new WvTextTemplate()
		{
			Template = "{{some text in the tag}}"
		};
		var data = SampleData;
		WvTextTemplateProcessResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		var lines =  new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
		Assert.Single(lines);
		Assert.Equal("{{some text in the tag}}", lines[0]);
	}


	[Fact]
	public void Text_Tag10()
	{
		var template = new WvTextTemplate()
		{
			Template = "{{notFoundColumnName}}"
		};
		var data = SampleData;
		WvTextTemplateProcessResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		var lines =  new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
		Assert.Single(lines);
		Assert.Equal("{{notFoundColumnName}}", lines[0]);
	}

	[Fact]
	public void Text_Tag11()
	{
		var template = new WvTextTemplate()
		{
			Template = "{{notFoundColumnName(F=H)}}"
		};
		var data = SampleData;
		WvTextTemplateProcessResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		var lines =  new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
		Assert.Single(lines);
		Assert.Equal("{{notFoundColumnName(F=H)}}", lines[0]);
	}

	[Fact]
	public void Text_Tag12()
	{
		var template = new WvTextTemplate()
		{
			Template = "text {{notFoundColumnName(F=H)}} text"
		};
		var data = SampleData;
		WvTextTemplateProcessResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		var lines =  new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
		Assert.Single(lines);
		Assert.Equal("text {{notFoundColumnName(F=H)}} text", lines[0]);
	}

	[Fact]
	public void Text_Tag13()
	{
		var template = new WvTextTemplate()
		{
			Template = "{{name(S=',')}} {{notFoundColumnName(F=H)}} text"
		};
		var data = SampleData.CreateAsNew(new List<int> { 0, 1});
		WvTextTemplateProcessResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		var lines =  new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
		Assert.Single(lines);
		Assert.Equal($"{data.Rows[0]["name"]},{data.Rows[1]["name"]} {{{{notFoundColumnName(F=H)}}}} text", lines[0]);
	}

	[Fact]
	public void Text_Tag_HorizontalIsForcedIfSeparator()
	{
		var template = new WvTextTemplate()
		{
			Template = "{{sku(S=',')}} test"
		};
		var data = SampleData.CreateAsNew(new List<int> { 0, 1, 2 });
		WvTextTemplateProcessResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		var lines =  new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
		Assert.Single(lines);
		Assert.Equal($"{data.Rows[0]["sku"]},{data.Rows[1]["sku"]},{data.Rows[2]["sku"]} test", lines[0]);
	}

	[Fact]
	public void Text_Tag_HorizontalIsForcedIfSeparatorNewLine()
	{
		var template = new WvTextTemplate()
		{
			Template = "{{sku(S='$rn')}}"
		};
		var data = SampleData.CreateAsNew(new List<int> { 0, 1, 2 });
		WvTextTemplateProcessResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		var lines =  new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
		Assert.Equal(3,lines.Count);

	}


	[Fact]
	public void Text_Tag_DataGroup()
	{
		var template = new WvTextTemplate()
		{
			Template = "{{sku}}{{name}}",
			GroupDataByColumns = new List<string> { "sku" }
		};
		var data = SampleData.CreateAsNew();
		data.Rows[1]["sku"] = data.Rows[0]["sku"];
		WvTextTemplateProcessResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.Equal(4, result.ResultItems.Count);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		var lines =  new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
		Assert.Single(lines);
		Assert.Equal($"{data.Rows[0]["sku"]}{data.Rows[0]["name"]}{data.Rows[0]["sku"]}{data.Rows[1]["name"]}", result.ResultItems[0].Result);
	}
	#endregion

	#region << Site Examples >>
	[Fact]
	public void Text_SiteExample1()
	{
		//Creating the DataTable
		DataTable dt = new DataTable();
		dt.Columns.Add("email", typeof(string));
		//Row 1
		var row = dt.NewRow();
		row["email"] = $"john@domain.com";
		dt.Rows.Add(row);
		//Row 2
		var row2 = dt.NewRow();
		row2["email"] = $"peter@domain.com";
		dt.Rows.Add(row2);

		//Creating the template
		WvTextTemplate template = new()
		{
			GroupDataByColumns = new List<string>(),
			Template = "{{email(S=\",\")}}"
		};

		//Execution
		WvTextTemplateProcessResult result = template.Process(dt);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		var lines =  new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
		Assert.Single(lines);
		Assert.Equal("john@domain.com,peter@domain.com", result.ResultItems[0].Result);
	}
	[Fact]
	public void Text_SiteExample2()
	{
		//Creating the DataTable
		DataTable dt = new DataTable();
		dt.Columns.Add("email", typeof(string));
		//Row 1
		var row = dt.NewRow();
		row["email"] = $"john@domain.com";
		dt.Rows.Add(row);
		//Row 2
		var row2 = dt.NewRow();
		row2["email"] = $"john@domain.com";
		dt.Rows.Add(row2);
		//Row 3
		var row3 = dt.NewRow();
		row3["email"] = $"peter@domain.com";
		dt.Rows.Add(row3);

		//Creating the template
		WvTextTemplate template = new()
		{
			GroupDataByColumns = new List<string>() { "email" },
			Template = "{{email(S=\",\")}}"
		};

		//Execution
		WvTextTemplateProcessResult result = template.Process(dt);
		Assert.NotNull(result);
		Assert.Equal(2, result.ResultItems.Count);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[1].Result));

		Assert.Equal("john@domain.com,john@domain.com", result.ResultItems[0].Result);
		Assert.Equal("peter@domain.com", result.ResultItems[1].Result);
	}
	#endregion
	
	#region << Inline Templates >>
	[Fact]
	public void Text_InlineTypeOneRecord()
	{
		var template = new WvTextTemplate()
		{
			Template = "{{<#}}{{position}} {{sku}} {{name}}{{#>}}"
		};
		var data = SampleData.CreateAsNew(new List<int> { 0 });
		WvTextTemplateProcessResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		var lines = new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
		Assert.Single(lines);
		Assert.Equal("1 sku1 item1", result.ResultItems[0].Result);
	}	
	[Fact]
	public void Text_InlineTypeTwoRecords()
	{
		var template = new WvTextTemplate()
		{
			Template = "{{<#}}{{position}} {{sku}} {{name}} {{#>}}"
		};
		var data = SampleData.CreateAsNew(new List<int> { 0, 1 });
		WvTextTemplateProcessResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		var lines = new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
		Assert.Single(lines);
		Assert.Equal("1 sku1 item1 2 sku2 item2 ", result.ResultItems[0].Result);
	}		
	[Fact]
	public void Text_InlineTypeTwoRecordsWithSeparator()
	{
		var template = new WvTextTemplate()
		{
			Template = "{{<#(S=',')}}{{position}} {{sku}} {{name}}{{#>}}"
		};
		var data = SampleData.CreateAsNew(new List<int> { 0, 1 });
		WvTextTemplateProcessResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		var lines = new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
		Assert.Single(lines);
		Assert.Equal("1 sku1 item1,2 sku2 item2", result.ResultItems[0].Result);
	}	
	
	[Fact]
	public void Text_InlineTypeTwoRecordsWithSeparatorInternalIsIgnored()
	{
		var template = new WvTextTemplate()
		{
			Template = "{{<#(S=',')}}{{<#}}{{position}}{{#>}} {{sku}} {{name}}{{#>}}"
		};
		var data = SampleData.CreateAsNew(new List<int> { 0, 1 });
		WvTextTemplateProcessResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		var lines = new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
		Assert.Single(lines);
		Assert.Equal("1 sku1 item1,2 sku2 item2", result.ResultItems[0].Result);
	}
	
	[Fact]
	public void Text_InlineTypeTwoRecordsWithSeparatorInternalIsIgnoredWithIndex()
	{
		var template = new WvTextTemplate()
		{
			Template = "{{<#[0,1](S=',')}}{{<#}}{{position}}{{#>}} {{sku}} {{name}}{{#>}}"
		};
		WvTextTemplateProcessResult? result = template.Process(SampleData);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		var lines = new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
		Assert.Single(lines);
		Assert.Equal("1 sku1 item1,2 sku2 item2", result.ResultItems[0].Result);
	}
    #endregion

    #region << List Data >>
    [Fact]
    public void Text_ListData1()
    {
        var template = new WvTextTemplate()
        {
            Template = "{{contact.first_name(S=',')}}"
        };
        var data = SampleData.CreateAsNew(new List<int> { 0 });
        WvTextTemplateProcessResult? result = template.Process(data);
        Assert.NotNull(result);
        Assert.Single(result.ResultItems);
        Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
        var lines = new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
        Assert.Single(lines);
        Assert.Equal("first_name0,first_name00", result.ResultItems[0].Result);
    }
    [Fact]
    public void Text_ListDataWithStartsWith()
    {
	    var template = new WvTextTemplate()
	    {
		    Template = "{{list1.(S=',')}}"
	    };
	    var data = SampleData.CreateAsNew(new List<int> { 0 });
	    WvTextTemplateProcessResult? result = template.Process(data);
	    Assert.NotNull(result);
	    Assert.Single(result.ResultItems);
	    Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
	    var lines = new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
	    Assert.Single(lines);
	    Assert.Equal("list1.test0,list1.test00,list1.test20,list1.test200", result.ResultItems[0].Result);
    }           
    [Fact]
    public void Text_ListDataWithStartsWith2()
    {
	    var template = new WvTextTemplate()
	    {
		    Template = "{{contact.(S=',')}}"
	    };
	    var data = SampleData.CreateAsNew(new List<int> { 0 });
	    WvTextTemplateProcessResult? result = template.Process(data);
	    Assert.NotNull(result);
	    Assert.Single(result.ResultItems);
	    Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
	    var lines = new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
	    Assert.Single(lines);
	    Assert.Equal("first_name0,first_name00,last_name0,last_name00", result.ResultItems[0].Result);
    }       
    [Fact]
    public void Text_ListDataWithStartsWith2MultiRow()
    {
	    var template = new WvTextTemplate()
	    {
		    Template = "{{contact.(S=',')}}"
	    };
	    var data = SampleData.CreateAsNew(new List<int> { 0,1 });
	    WvTextTemplateProcessResult? result = template.Process(data);
	    Assert.NotNull(result);
	    Assert.Single(result.ResultItems);
	    Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
	    var lines = new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
	    Assert.Single(lines);
	    Assert.Equal("first_name0,first_name00,last_name0,last_name00,first_name1,first_name11,last_name1,last_name11", result.ResultItems[0].Result);
    }          
    [Fact]
    public void Text_ListDataWithStartsWith2WithIndex()
    {
	    var template = new WvTextTemplate()
	    {
		    Template = "{{contact.[0](S=',')}}"
	    };
	    var data = SampleData.CreateAsNew(new List<int> { 0,1 });
	    WvTextTemplateProcessResult? result = template.Process(data);
	    Assert.NotNull(result);
	    Assert.Single(result.ResultItems);
	    Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
	    var lines = new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
	    Assert.Single(lines);
	    Assert.Equal("first_name0,first_name00,last_name0,last_name00", result.ResultItems[0].Result);
    }       
    [Fact]
    public void Text_InlineTemplateListDataWithStartsWithWithItemName()
    {
	    var template = new WvTextTemplate()
	    {
		    Template = "{{<#contact.}}{{first_name}} {{last_name}} {{#>}}"
	    };
	    var data = SampleData.CreateAsNew(new List<int> { 0,1 });
	    WvTextTemplateProcessResult? result = template.Process(data);
	    Assert.NotNull(result);
	    Assert.Single(result.ResultItems);
	    Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
	    var lines = new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
	    Assert.Single(lines);
	    Assert.Equal("first_name0 last_name0 first_name00 last_name00 first_name1 last_name1 first_name11 last_name11 ", result.ResultItems[0].Result);
    }        
    [Fact]
    public void Text_InlineTemplateListDataWithStartsWithWithItemNameOneRow()
    {
	    var template = new WvTextTemplate()
	    {
		    Template = "{{<#contact.}}{{first_name}} {{last_name}} {{#>}}"
	    };
	    var data = SampleData.CreateAsNew(new List<int> { 0 });
	    WvTextTemplateProcessResult? result = template.Process(data);
	    Assert.NotNull(result);
	    Assert.Single(result.ResultItems);
	    Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
	    var lines = new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
	    Assert.Single(lines);
	    Assert.Equal("first_name0 last_name0 first_name00 last_name00 ", result.ResultItems[0].Result);
    }    
    [Fact]
    public void Text_InlineTemplateListDataWithStartsWithWithItemNameWithIndex()
    {
	    var template = new WvTextTemplate()
	    {
		    Template = "{{<#contact.[0]}}{{first_name}} {{last_name}} {{#>}}"
	    };
	    var data = SampleData.CreateAsNew(new List<int> { 0, 1 });
	    WvTextTemplateProcessResult? result = template.Process(data);
	    Assert.NotNull(result);
	    Assert.Single(result.ResultItems);
	    Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
	    var lines = new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
	    Assert.Single(lines);
	    Assert.Equal("first_name0 last_name0 first_name00 last_name00 ", result.ResultItems[0].Result);
    }       
    [Fact]
    public void Text_InlineTemplateListDataWithStartsWithoutItemnName()
    {
	    var template = new WvTextTemplate()
	    {
		    Template = "{{<#[0]}}{{contact.first_name[0]}} {{contact.last_name[0]}} {{#>}}"
	    };
	    var data = SampleData.CreateAsNew(new List<int> { 0 });
	    WvTextTemplateProcessResult? result = template.Process(data);
	    Assert.NotNull(result);
	    Assert.Single(result.ResultItems);
	    Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
	    var lines = new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
	    Assert.Single(lines);
	    Assert.Equal("first_name0first_name00 last_name0last_name00 ", result.ResultItems[0].Result);
    }      
    
    
    [Fact]
    public void Text_InlineTemplateListDataWithExactColumnNoMatterThereIsStartWidth()
    {
	    var template = new WvTextTemplate()
	    {
		    Template = "{{<#list1[0]}}test: {{list1}} {{#>}}"
	    };
	    var data = SampleData.CreateAsNew(new List<int> { 0 });
	    WvTextTemplateProcessResult? result = template.Process(data);
	    Assert.NotNull(result);
	    Assert.Single(result.ResultItems);
	    Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
	    var lines = new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
	    Assert.Single(lines);
	    Assert.Equal("test: list10 test: list100 ", result.ResultItems[0].Result);
    }

    [Fact]
    public void Text_InlineTemplateWithIndexInTemplate()
    {
        var template = new WvTextTemplate()
        {
            Template = "{{<#list.[0]}}{{position[0]}}{{#>}}"
        };
        WvTextTemplateProcessResult? result = template.Process(SampleData);
        Assert.NotNull(result);
        Assert.Single(result.ResultItems);
        Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
        var lines = new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
        Assert.Single(lines);
        Assert.Equal("1", result.ResultItems[0].Result);
    }

    [Fact]
    public void Text_NormalListDataWithNoIndex()
    {
        var template = new WvTextTemplate()
        {
            Template = "{{list1[0]}}"
        };
        WvTextTemplateProcessResult? result = template.Process(SampleData);
        Assert.NotNull(result);
        Assert.Single(result.ResultItems);
        Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
        var lines = new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
        Assert.Single(lines);
        Assert.Equal("list10list100", result.ResultItems[0].Result);
    }

    [Fact]
    public void Text_NormalListDataWithTwoIndexes()
    {
        var template = new WvTextTemplate()
        {
            Template = "{{list1[0][1]}}"
        };
        WvTextTemplateProcessResult? result = template.Process(SampleData);
        Assert.NotNull(result);
        Assert.Single(result.ResultItems);
        Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
        var lines = new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
        Assert.Single(lines);
        Assert.Equal("list100", result.ResultItems[0].Result);
    }

    #endregion

    #region << Condition Tag >>
    [Fact]
	public void Text_ConditionOneRecordNoRules()
	{
		var template = new WvTextTemplate()
		{
			Template = "{{<?}}{{position}} {{sku}} {{name}}{{?>}}"
		};
		var data = SampleData.CreateAsNew(new List<int> { 0 });
		WvTextTemplateProcessResult? result = template.Process(data);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.False(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
		var lines = new TestUtils().GetLines(result.ResultItems[0].Result ?? String.Empty);
		Assert.Single(lines);
		Assert.Equal("1 sku1 item1", result.ResultItems[0].Result);
	}		
	[Fact]
	public void Text_ConditionOneRecordOneRules()
	{
		var template = new WvTextTemplate()
		{
			Template = "{{<?(position >= 200)}}{{position}} {{sku}} {{name}}{{?>}}"
		};
		WvTextTemplateProcessResult? result = template.Process(SampleData);
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.True(String.IsNullOrWhiteSpace(result.ResultItems[0].Result));
	}			
	#endregion


	
}

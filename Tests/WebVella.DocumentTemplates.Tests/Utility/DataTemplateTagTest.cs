using System.Data;
using System.Globalization;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Tests.Models;

namespace WebVella.DocumentTemplates.Tests.Utility;
public class DataTemplateTagTest : TestBase
{
	public DataTemplateTagTest() : base() { }

	#region << Parsing >>
	[Fact]
	public void NullTemplateShouldReturnNoTags()
	{
		//Given
		string? template = null;
		//When
		List<WvTemplateTag> result = WvTemplateUtility.GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Empty(result);
	}

	[Fact]
	public void EmptyTemplateShouldReturnNoTags()
	{
		//Given
		string template = string.Empty;
		//When
		List<WvTemplateTag> result = WvTemplateUtility.GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Empty(result);
	}

	[Fact]
	public void ExactTemplateShouldReturnOneDataTag()
	{
		//Given
		var columnName = "column_name";
		string template = "{{" + columnName + "}}";
		//When
		List<WvTemplateTag> result = WvTemplateUtility.GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(columnName, result[0].Name);
		Assert.Equal(WvTemplateTagType.Data, result[0].Type);
		Assert.NotNull(result[0].IndexList);
		Assert.Empty(result[0].IndexList);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Empty(result[0].ParamGroups);
	}

	[Fact]
	public void ExactTemplateShouldReturnOneDataTagAlwaysLowecase()
	{
		//Given
		var columnName = "ColumnName";
		string template = "{{" + columnName + "}}";
		//When
		List<WvTemplateTag> result = WvTemplateUtility.GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(columnName.ToLowerInvariant(), result[0].Name);
		Assert.Equal(WvTemplateTagType.Data, result[0].Type);
		Assert.NotNull(result[0].IndexList);
		Assert.Empty(result[0].IndexList);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Empty(result[0].ParamGroups);
	}

	[Fact]
	public void ExactTemplateShouldReturnOneDataTag_WithSpaces()
	{
		//Given
		var columnName = "columnName";
		string template = " {{" + columnName + "}} ";
		//When
		List<WvTemplateTag> result = WvTemplateUtility.GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template.Trim(), result[0].FullString);
		Assert.Equal(columnName.ToLowerInvariant(), result[0].Name);
		Assert.Equal(WvTemplateTagType.Data, result[0].Type);
		Assert.NotNull(result[0].IndexList);
		Assert.Empty(result[0].IndexList);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Empty(result[0].ParamGroups);
	}

	[Fact]
	public void ExactTemplateShouldReturnOneDataTag_WithSpaces2()
	{
		//Given
		var columnName = "column_name";
		string template = "{{ " + columnName + " }}";
		//When
		List<WvTemplateTag> result = WvTemplateUtility.GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(columnName, result[0].Name);
		Assert.Equal(WvTemplateTagType.Data, result[0].Type);
		Assert.NotNull(result[0].IndexList);
		Assert.Empty(result[0].IndexList);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Empty(result[0].ParamGroups);
	}

	[Fact]
	public void ExactTemplateShouldReturnOneDataTag_WithSpaces3()
	{
		//Given
		var columnName = "column_name";
		string template = "{{  " + columnName + "     }}";
		//When
		List<WvTemplateTag> result = WvTemplateUtility.GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(columnName, result[0].Name);
		Assert.Equal(WvTemplateTagType.Data, result[0].Type);
		Assert.NotNull(result[0].IndexList);
		Assert.Empty(result[0].IndexList);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Empty(result[0].ParamGroups);
	}

	[Fact]
	public void OnInTextTemplateShouldReturnOneDataTag()
	{
		//Given
		var columnName = "column_name";
		string template = "this is {{" + columnName + "}} a test";
		//When
		List<WvTemplateTag> result = WvTemplateUtility.GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal("{{" + columnName + "}}", result[0].FullString);
		Assert.Equal(columnName, result[0].Name);
		Assert.Equal(WvTemplateTagType.Data, result[0].Type);
		Assert.NotNull(result[0].IndexList);
		Assert.Empty(result[0].IndexList);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Empty(result[0].ParamGroups);
	}

	[Fact]
	public void OnInTextTemplateShouldReturnOneDataTagWithOneParam()
	{
		//Given
		var columnName = "column_name";
		string template = "this is {{" + columnName + "(\"test\")}} a test";
		//When
		List<WvTemplateTag> result = WvTemplateUtility.GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal("{{" + columnName + "(\"test\")}}", result[0].FullString);
		Assert.Equal(columnName, result[0].Name);
		Assert.Equal(WvTemplateTagType.Data, result[0].Type);
		Assert.NotNull(result[0].IndexList);
		Assert.Empty(result[0].IndexList);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Single(result[0].ParamGroups);
		Assert.NotNull(result[0].ParamGroups[0].Parameters);
		Assert.Single(result[0].ParamGroups[0].Parameters);
		Assert.Equal("test", result[0].ParamGroups[0].Parameters[0].ValueString);
	}

	[Fact]
	public void OnInTextTemplateShouldReturnOneDataTagWithOneNamedParam()
	{
		//Given
		var columnName = "column_name";
		var paramName = "param1";
		var paramValue = "test";
		string template = "{{" + columnName + "(" + paramName + "=\"" + paramValue + "\")}}";
		//When
		List<WvTemplateTag> result = WvTemplateUtility.GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(columnName, result[0].Name);
		Assert.Equal(WvTemplateTagType.Data, result[0].Type);
		Assert.NotNull(result[0].IndexList);
		Assert.Empty(result[0].IndexList);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Single(result[0].ParamGroups);
		Assert.NotNull(result[0].ParamGroups[0].Parameters);
		Assert.Single(result[0].ParamGroups[0].Parameters);
		Assert.Equal(paramName.ToLowerInvariant(), result[0].ParamGroups[0].Parameters[0].Name);
		Assert.Equal(paramValue, result[0].ParamGroups[0].Parameters[0].ValueString);
	}


	[Fact]
	public void OnInTextTemplateShouldReturnOneDataTagWithOneNamedParamWithSpaces()
	{
		//Given
		var columnName = "column_name";
		var paramName = "param1";
		var paramValue = "test";
		string template = "{{" + columnName + " ( " + paramName + " = \"" + paramValue + "\" )}}";
		//When
		List<WvTemplateTag> result = WvTemplateUtility.GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(columnName, result[0].Name);
		Assert.Equal(WvTemplateTagType.Data, result[0].Type);
		Assert.NotNull(result[0].IndexList);
		Assert.Empty(result[0].IndexList);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Single(result[0].ParamGroups);
		Assert.NotNull(result[0].ParamGroups[0].Parameters);
		Assert.Single(result[0].ParamGroups[0].Parameters);
		Assert.Equal(paramName, result[0].ParamGroups[0].Parameters[0].Name);
		Assert.Equal(paramValue, result[0].ParamGroups[0].Parameters[0].ValueString);
	}

	[Fact]
	public void OnInTextTemplateShouldReturnOneDataTagWithOneNamedParamShouldBeLowered()
	{
		//Given
		var columnName = "column_name";
		var paramName = "TestParam";
		var paramValue = "test";
		string template = "{{" + columnName + "(" + paramName + "=\"" + paramValue + "\")}}";
		//When
		List<WvTemplateTag> result = WvTemplateUtility.GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(columnName, result[0].Name);
		Assert.Equal(WvTemplateTagType.Data, result[0].Type);
		Assert.NotNull(result[0].IndexList);
		Assert.Empty(result[0].IndexList);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Single(result[0].ParamGroups);
		Assert.NotNull(result[0].ParamGroups[0].Parameters);
		Assert.Single(result[0].ParamGroups[0].Parameters);
		Assert.Equal(paramName.ToLowerInvariant(), result[0].ParamGroups[0].Parameters[0].Name);
		Assert.Equal(paramValue, result[0].ParamGroups[0].Parameters[0].ValueString);
	}

	[Fact]
	public void ExactTemplateShouldReturnMultipleDataTag()
	{
		//Given
		var columnName = "column_name";
		var columnName2 = "column_name2";
		string template = "{{" + columnName + "}}{{" + columnName2 + "}}";
		//When
		List<WvTemplateTag> result = WvTemplateUtility.GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Equal(2, result.Count);
		var column1Tag = result.FirstOrDefault(x => x.Name == columnName);
		var column2Tag = result.FirstOrDefault(x => x.Name == columnName2);
		Assert.NotNull(column1Tag);
		Assert.NotNull(column2Tag);

		Assert.Equal("{{" + columnName + "}}", column1Tag.FullString);
		Assert.Equal(columnName, column1Tag.Name);
		Assert.Equal(WvTemplateTagType.Data, column1Tag.Type);
		Assert.NotNull(column1Tag.IndexList);
		Assert.Empty(column1Tag.IndexList);
		Assert.NotNull(column1Tag.ParamGroups);
		Assert.Empty(column1Tag.ParamGroups);

		Assert.Equal("{{" + columnName2 + "}}", column2Tag.FullString);
		Assert.Equal(columnName2, column2Tag.Name);
		Assert.Equal(WvTemplateTagType.Data, column2Tag.Type);
		Assert.NotNull(column2Tag.IndexList);
		Assert.Empty(column2Tag.IndexList);
		Assert.NotNull(column2Tag.ParamGroups);
		Assert.Empty(column2Tag.ParamGroups);
	}

	[Fact]
	public void ExactTemplateShouldReturnMultipleDataTagWithParams()
	{
		//Given
		var columnName = "column_name";
		var columnNameParam1Name = "Param1";
		var columnNameParam1Value = "Param =1 Value";
		var columnNameParam2Name = "Param2";
		var columnNameParam2Value = "Param2Valu";
		var columnName2 = "column_name2";
		var columnName2Param1Name = "Param1";
		var columnName2Param1Value = "Param1Value";
		var columnName2Param2Name = "Param2";
		var columnName2Param2Value = "Param2Valu";
		string template = "{{" + columnName + "(" + columnNameParam1Name + "=\"" + columnNameParam1Value + "\", " + columnNameParam2Name + "=\"" + columnNameParam2Value + "\")}}" +
		"{{" + columnName2 + "(" + columnName2Param1Name + "=\"" + columnName2Param1Value + "\", " + columnName2Param2Name + "=\"" + columnName2Param2Value + "\")}}";
		//When
		List<WvTemplateTag> result = WvTemplateUtility.GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Equal(2, result.Count);
		var column1Tag = result.FirstOrDefault(x => x.Name == columnName);
		var column2Tag = result.FirstOrDefault(x => x.Name == columnName2);
		Assert.NotNull(column1Tag);
		Assert.NotNull(column2Tag);

		Assert.Equal("{{" + columnName + "(" + columnNameParam1Name + "=\"" + columnNameParam1Value + "\", " + columnNameParam2Name + "=\"" + columnNameParam2Value + "\")}}", column1Tag.FullString);
		Assert.Equal(columnName, column1Tag.Name);
		Assert.Equal(WvTemplateTagType.Data, column1Tag.Type);
		Assert.NotNull(column1Tag.IndexList);
		Assert.Empty(column1Tag.IndexList);
		Assert.NotNull(column1Tag.ParamGroups);
		Assert.Single(column1Tag.ParamGroups);
		Assert.NotNull(column1Tag.ParamGroups[0].Parameters);
		Assert.Equal(2, column1Tag.ParamGroups[0].Parameters.Count);
		var column1Param1 = column1Tag.ParamGroups[0].Parameters.FirstOrDefault(x => x.Name == columnNameParam1Name.ToLowerInvariant());
		var column1Param2 = column1Tag.ParamGroups[0].Parameters.FirstOrDefault(x => x.Name == columnNameParam2Name.ToLowerInvariant());
		Assert.NotNull(column1Param1);
		Assert.NotNull(column1Param2);
		Assert.Equal(columnNameParam1Value, column1Param1.ValueString);
		Assert.Equal(columnNameParam2Value, column1Param2.ValueString);

		Assert.Equal("{{" + columnName2 + "(" + columnName2Param1Name + "=\"" + columnName2Param1Value + "\", " + columnName2Param2Name + "=\"" + columnName2Param2Value + "\")}}", column2Tag.FullString);
		Assert.Equal(columnName2, column2Tag.Name);
		Assert.Equal(WvTemplateTagType.Data, column2Tag.Type);
		Assert.NotNull(column2Tag.IndexList);
		Assert.Empty(column2Tag.IndexList);
		Assert.NotNull(column2Tag.ParamGroups);
		Assert.Single(column2Tag.ParamGroups);
		Assert.NotNull(column2Tag.ParamGroups[0].Parameters);
		Assert.Equal(2, column2Tag.ParamGroups[0].Parameters.Count);
		var column2Param1 = column2Tag.ParamGroups[0].Parameters.FirstOrDefault(x => x.Name == columnName2Param1Name.ToLowerInvariant());
		var column2Param2 = column2Tag.ParamGroups[0].Parameters.FirstOrDefault(x => x.Name == columnName2Param2Name.ToLowerInvariant());
		Assert.NotNull(column2Param1);
		Assert.NotNull(column2Param2);
		Assert.Equal(columnName2Param1Value, column2Param1.ValueString);
		Assert.Equal(columnName2Param2Value, column2Param2.ValueString);
	}

	[Fact]
	public void ExactTemplateShouldReturnMultipleDataTagIntext()
	{
		//Given
		var columnName = "column_name";
		var columnName2 = "column_name2";
		string template = "test is with {{" + columnName + "}} and longer text {{" + columnName2 + "}} everywhere";
		//When
		List<WvTemplateTag> result = WvTemplateUtility.GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Equal(2, result.Count);
		var column1Tag = result.FirstOrDefault(x => x.Name == columnName);
		var column2Tag = result.FirstOrDefault(x => x.Name == columnName2);
		Assert.NotNull(column1Tag);
		Assert.NotNull(column2Tag);

		Assert.Equal("{{" + columnName + "}}", column1Tag.FullString);
		Assert.Equal(columnName, column1Tag.Name);
		Assert.Equal(WvTemplateTagType.Data, column1Tag.Type);
		Assert.Empty(column1Tag.ParamGroups);

		Assert.Equal("{{" + columnName2 + "}}", column2Tag.FullString);
		Assert.Equal(columnName2, column2Tag.Name);
		Assert.Equal(WvTemplateTagType.Data, column2Tag.Type);
		Assert.Empty(column2Tag.ParamGroups);
	}

	[Fact]
	public void ExactTemplateShouldReturnEmptyGroups()
	{
		//Given
		var columnName = "column_name";
		string template = "{{" + columnName + "()(\"test\")(test=1)}}";
		//When
		List<WvTemplateTag> result = WvTemplateUtility.GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		var column1Tag = result.FirstOrDefault(x => x.Name == columnName);
		Assert.NotNull(column1Tag);
		Assert.Equal("{{" + columnName + "()(\"test\")(test=1)}}", column1Tag.FullString);
		Assert.Equal(columnName, column1Tag.Name);
		Assert.Equal(WvTemplateTagType.Data, column1Tag.Type);
		Assert.NotNull(column1Tag.ParamGroups);
		Assert.Equal(3, column1Tag.ParamGroups.Count);

		Assert.Equal("()", column1Tag.ParamGroups[0].FullString);
		Assert.Empty(column1Tag.ParamGroups[0].Parameters);

		Assert.Equal("(\"test\")", column1Tag.ParamGroups[1].FullString);
		Assert.Single(column1Tag.ParamGroups[1].Parameters);


		Assert.Equal("(test=1)", column1Tag.ParamGroups[2].FullString);
		Assert.Single(column1Tag.ParamGroups[2].Parameters);
	}

	[Fact]
	public void ExactTemplateShouldReturnOneTagWithMultipleParameterGroups()
	{
		//Given
		var columnName = "column_name";
		var columnNameParam1Name = "Param1";
		var columnNameParam1Value = "Param =1 Value";
		var columnNameParam2Name = "Param2";
		var columnNameParam2Value = "Param2Valu";
		string template = "{{" + columnName +
		"(" + columnNameParam1Name + "=\"" + columnNameParam1Value + "\")" +
		"(" + columnNameParam2Name + "=\"" + columnNameParam2Value + "\")" +
		"}}";
		//When
		List<WvTemplateTag> result = WvTemplateUtility.GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		var column1Tag = result.FirstOrDefault(x => x.Name == columnName);
		Assert.NotNull(column1Tag);

		Assert.Equal("{{column_name(Param1=\"Param =1 Value\")(Param2=\"Param2Valu\")}}", column1Tag.FullString);
		Assert.Equal(columnName, column1Tag.Name);
		Assert.Equal(WvTemplateTagType.Data, column1Tag.Type);
		Assert.Equal(2, column1Tag.ParamGroups.Count);

		Assert.Single(column1Tag.ParamGroups[0].Parameters);
		Assert.Single(column1Tag.ParamGroups[1].Parameters);

		Assert.Equal(columnNameParam1Name.ToLowerInvariant(), column1Tag.ParamGroups[0].Parameters[0].Name);
		Assert.Equal(columnNameParam1Value, column1Tag.ParamGroups[0].Parameters[0].ValueString);

		Assert.Equal(columnNameParam2Name.ToLowerInvariant(), column1Tag.ParamGroups[1].Parameters[0].Name);
		Assert.Equal(columnNameParam2Value, column1Tag.ParamGroups[1].Parameters[0].ValueString);
	}

	[Fact]
	public void ExactTemplateWithIndexingShouldReturnOneDataTagAndIndex()
	{
		//Given
		var columnName = "column_name";
		int columnIndex = 5;
		string template = "{{" + columnName + "[" + columnIndex + "]}}";
		//When
		List<WvTemplateTag> result = WvTemplateUtility.GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(columnName, result[0].Name);
		Assert.Equal(WvTemplateTagType.Data, result[0].Type);
		Assert.NotNull(result[0].IndexList);
		Assert.Single(result[0].IndexList);
		Assert.Equal(columnIndex, result[0].IndexList[0]);
		Assert.Empty(result[0].ParamGroups);
	}

	[Fact]
	public void ExactTemplateWithIndexingShouldReturnOneDataTagAndIndexZERO()
	{
		//Given
		var columnName = "column_name";
		int columnIndex = 5;
		string template = "{{" + columnName + "[][" + columnIndex + "]}}";
		//When
		List<WvTemplateTag> result = WvTemplateUtility.GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(columnName, result[0].Name);
		Assert.Equal(WvTemplateTagType.Data, result[0].Type);
		Assert.NotNull(result[0].IndexList);
		Assert.Equal(2, result[0].IndexList.Count);
		Assert.Equal(0, result[0].IndexList[0]);
		Assert.Equal(columnIndex, result[0].IndexList[1]);
		Assert.Empty(result[0].ParamGroups);
	}

	[Fact]
	public void ExactTemplateWithMultiIndexingShouldReturnOneDataTagAndMultiIndex()
	{
		//Given
		var columnName = "column_name";
		int columnIndex = 5;
		int columnIndex2 = 2;
		string template = "{{" + columnName + "[" + columnIndex + "][" + columnIndex2 + "]}}";
		//When
		List<WvTemplateTag> result = WvTemplateUtility.GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(columnName, result[0].Name);
		Assert.Equal(WvTemplateTagType.Data, result[0].Type);
		Assert.NotNull(result[0].IndexList);
		Assert.Equal(2, result[0].IndexList.Count);
		Assert.Equal(columnIndex, result[0].IndexList[0]);
		Assert.Equal(columnIndex2, result[0].IndexList[1]);
		Assert.Empty(result[0].ParamGroups);
	}

	[Fact]
	public void OnInTextTemplateShouldReturnOneDataTagWitTwoNamedGroupParams()
	{
		//Given
		var columnName = "column_name";
		var paramName = "param1";
		var paramValue = "test";
		var paramName2 = "param2";
		var paramValue2 = "test2";
		string template = "{{" + columnName + "(" + paramName + "='" + paramValue + "')" + "(" + paramName2 + "='" + paramValue2 + "')}}";
		//When
		List<WvTemplateTag> result = WvTemplateUtility.GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(columnName, result[0].Name);
		Assert.Equal(WvTemplateTagType.Data, result[0].Type);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Equal(2, result[0].ParamGroups.Count);
		Assert.Single(result[0].ParamGroups[0].Parameters);
		Assert.Equal(paramName.ToLowerInvariant(), result[0].ParamGroups[0].Parameters[0].Name);
		Assert.Equal(paramValue, result[0].ParamGroups[0].Parameters[0].ValueString);

		Assert.Single(result[0].ParamGroups[1].Parameters);
		Assert.Equal(paramName2.ToLowerInvariant(), result[0].ParamGroups[1].Parameters[0].Name);
		Assert.Equal(paramValue2, result[0].ParamGroups[1].Parameters[0].ValueString);
	}

	[Fact]
	public void OnInTextTemplateShouldReturnOneDataTagWitGroupWithParams()
	{
		//Given
		var columnName = "column_name";
		var paramName = "param1";
		var paramValue = "test";
		var paramName2 = "param2";
		var paramValue2 = "test2";
		string template = "{{" + columnName + "(" + paramName + "='" + paramValue + "'," + paramName2 + "='" + paramValue2 + "')}}";
		//When
		List<WvTemplateTag> result = WvTemplateUtility.GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(columnName, result[0].Name);
		Assert.Equal(WvTemplateTagType.Data, result[0].Type);

		Assert.NotNull(result[0].ParamGroups);
		Assert.Single(result[0].ParamGroups);
		Assert.Equal(2, result[0].ParamGroups[0].Parameters.Count);

		Assert.Equal(paramName.ToLowerInvariant(), result[0].ParamGroups[0].Parameters[0].Name);
		Assert.Equal(paramValue, result[0].ParamGroups[0].Parameters[0].ValueString);

		Assert.Equal(paramName2.ToLowerInvariant(), result[0].ParamGroups[0].Parameters[1].Name);
		Assert.Equal(paramValue2, result[0].ParamGroups[0].Parameters[1].ValueString);
	}

	[Fact]
	public void OnInTextTemplateShouldReturnOneDataTagWitGroupWithParamsWithQuotes()
	{
		//Given
		var columnName = "column_name";
		var paramName = "param1";
		var paramValue = "test";
		var paramName2 = "param2";
		var paramValue2 = "test2";
		string template = "{{" + columnName + "(" + paramName + "=\"" + paramValue + "\"," + paramName2 + "=\"" + paramValue2 + "\")}}";
		//When
		List<WvTemplateTag> result = WvTemplateUtility.GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(columnName, result[0].Name);
		Assert.Equal(WvTemplateTagType.Data, result[0].Type);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Single(result[0].ParamGroups);
		Assert.Equal(2, result[0].ParamGroups[0].Parameters.Count);

		Assert.Equal(paramName.ToLowerInvariant(), result[0].ParamGroups[0].Parameters[0].Name);
		Assert.Equal(paramValue, result[0].ParamGroups[0].Parameters[0].ValueString);

		Assert.Equal(paramName2.ToLowerInvariant(), result[0].ParamGroups[0].Parameters[1].Name);
		Assert.Equal(paramValue2, result[0].ParamGroups[0].Parameters[1].ValueString);
	}

	[Fact]
	public void OnInTextTemplateShouldReturnOneDataTagWitGroupWithParamsWithQuotes2()
	{
		//Given
		var columnName = "column_name";
		var paramName = "F";
		var paramValue = "H";
		var paramName2 = "S";
		var paramValue2 = ", ";
		var paramName3 = "B";
		var paramValue3 = ", ";
		//var text = "{{sku(F=H,S=',',B=\", \")}}";
		string template = "{{" + columnName + "(" + paramName + "=" + paramValue + "," + paramName2 + "='" + paramValue2 + "'," + paramName3 + "=\"" + paramValue3 + "\")}}";
		//When
		List<WvTemplateTag> result = WvTemplateUtility.GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(columnName, result[0].Name);
		Assert.Equal(WvTemplateTagType.Data, result[0].Type);

		Assert.NotNull(result[0].ParamGroups);
		Assert.Single(result[0].ParamGroups);
		Assert.Equal(3, result[0].ParamGroups[0].Parameters.Count);

		Assert.Equal(paramName.ToLowerInvariant(), result[0].ParamGroups[0].Parameters[0].Name);
		Assert.Equal(paramValue, result[0].ParamGroups[0].Parameters[0].ValueString);

		Assert.Equal(paramName2.ToLowerInvariant(), result[0].ParamGroups[0].Parameters[1].Name);
		Assert.Equal(paramValue2, result[0].ParamGroups[0].Parameters[1].ValueString);

		Assert.Equal(paramName3.ToLowerInvariant(), result[0].ParamGroups[0].Parameters[2].Name);
		Assert.Equal(paramValue3, result[0].ParamGroups[0].Parameters[2].ValueString);
	}

	[Fact]
	public void ExactTemplateWithFunctionSupported()
	{
		//Given
		var function = "SUM";
		string template = "{{=" + function + "(A1:B1)}}";
		//When
		List<WvTemplateTag> result = WvTemplateUtility.GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(function.ToLowerInvariant(), result[0].Name);
		Assert.Equal(WvTemplateTagType.Function, result[0].Type);
		Assert.Equal(function.ToLowerInvariant(), result[0].FunctionName);
		Assert.Single(result[0].ParamGroups);
		Assert.Equal(2, result[0].ParamGroups[0].Parameters.Count);
		Assert.Equal("A1", result[0].ParamGroups[0].Parameters[0].ValueString);
		Assert.Equal("B1", result[0].ParamGroups[0].Parameters[1].ValueString);

	}

	[Fact]
	public void ExactTemplateWithFunctionNotSupported()
	{
		//Given
		var function = "SUM123";
		string template = "{{=" + function + "(A1:B1)}}";
		//When
		List<WvTemplateTag> result = WvTemplateUtility.GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(function.ToLowerInvariant(), result[0].Name);
		Assert.Equal(WvTemplateTagType.Function, result[0].Type);
		Assert.True(String.IsNullOrWhiteSpace(result[0].FunctionName));
		Assert.Single(result[0].ParamGroups);
		Assert.Equal(2, result[0].ParamGroups[0].Parameters.Count);
		Assert.Equal("A1", result[0].ParamGroups[0].Parameters[0].ValueString);
		Assert.Equal("B1", result[0].ParamGroups[0].Parameters[1].ValueString);

	}


	[Fact]
	public void ExactTemplateWithExcelFunctionSupported()
	{
		//Given
		var function = "SUM";
		string template = "{{==" + function + "(A1:B1)}}";
		//When
		List<WvTemplateTag> result = WvTemplateUtility.GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(function.ToLowerInvariant(), result[0].Name);
		Assert.Equal(WvTemplateTagType.ExcelFunction, result[0].Type);
		Assert.Equal(function.ToLowerInvariant(), result[0].FunctionName);
		Assert.Single(result[0].ParamGroups);
		Assert.Equal(2, result[0].ParamGroups[0].Parameters.Count);
		Assert.Equal("A1", result[0].ParamGroups[0].Parameters[0].ValueString);
		Assert.Equal("B1", result[0].ParamGroups[0].Parameters[1].ValueString);

	}

	[Fact]
	public void ExactTemplateWithExcelFunctionNotSupported()
	{
		//Given
		var function = "SUM123";
		string template = "{{==" + function + "(A1:B1)}}";
		//When
		List<WvTemplateTag> result = WvTemplateUtility.GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(function.ToLowerInvariant(), result[0].Name);
		Assert.Equal(WvTemplateTagType.ExcelFunction, result[0].Type);
		Assert.True(String.IsNullOrWhiteSpace(result[0].FunctionName));
		Assert.Single(result[0].ParamGroups);
		Assert.Equal(2, result[0].ParamGroups[0].Parameters.Count);
		Assert.Equal("A1", result[0].ParamGroups[0].Parameters[0].ValueString);
		Assert.Equal("B1", result[0].ParamGroups[0].Parameters[1].ValueString);

	}
	#endregion

	#region << Processing >>
	[Fact]
	public void TemplateProcessShouldReturnNoResultsIfEmpty()
	{
		//Given
		string template = "";
		DataTable ds = SampleData;
		//When
		WvTemplateTagResultList result = WvTemplateUtility.ProcessTemplateTag(template, ds, DefaultCulture);
		//Then
		Assert.NotNull(result);
		Assert.Single(result.Values);
		Assert.True(String.IsNullOrWhiteSpace((string?)result.Values[0]));
	}

	[Fact]
	public void TemplateProcessShouldReturnNoResultsIfNoTag()
	{
		//Given
		string template = "sometext test";
		DataTable ds = SampleData;
		//When
		WvTemplateTagResultList result = WvTemplateUtility.ProcessTemplateTag(template, ds, DefaultCulture);
		//Then
		Assert.NotNull(result);
		Assert.Single(result.Values);
		Assert.Equal(template, result.Values[0]);
	}

	[Fact]
	public void TemplateProcessShouldReturnResultsIfTagCannotBeProcessed()
	{
		//Given
		string template = "{{}}";
		DataTable ds = SampleData;
		//When
		WvTemplateTagResultList result = WvTemplateUtility.ProcessTemplateTag(template, ds, DefaultCulture);
		//Then
		Assert.NotNull(result);
		Assert.Single(result.Values);
		Assert.Equal(template, result.Values[0]);
	}

	[Fact]
	public void TemplateProcessShouldReturnResultsIfTagCanBeProcessed()
	{
		//Given
		string template = "{{name}}";
		DataTable ds = SampleData;
		//When
		WvTemplateTagResultList result = WvTemplateUtility.ProcessTemplateTag(template, ds, DefaultCulture);
		//Then
		Assert.NotNull(result);
		Assert.Equal(5, result.Values.Count);
		Assert.Equal(ds.Rows[0]["name"]?.ToString(), result.Values[0]);
		Assert.Equal(ds.Rows[1]["name"]?.ToString(), result.Values[1]);
		Assert.Equal(ds.Rows[2]["name"]?.ToString(), result.Values[2]);
		Assert.Equal(ds.Rows[3]["name"]?.ToString(), result.Values[3]);
		Assert.Equal(ds.Rows[4]["name"]?.ToString(), result.Values[4]);
	}

	[Fact]
	public void TemplateProcessShouldReturnResultsIfTagCanBeProcessedMulti()
	{
		//Given
		string template = "{{position}}.{{name}}";
		DataTable ds = SampleData;
		//When
		WvTemplateTagResultList result = WvTemplateUtility.ProcessTemplateTag(template, ds, DefaultCulture);
		//Then
		Assert.NotNull(result);
		Assert.Equal(5, result.Values.Count);
		Assert.Equal(ds.Rows[0]["position"] + "." + ds.Rows[0]["name"], result.Values[0]);
		Assert.Equal(ds.Rows[1]["position"] + "." + ds.Rows[1]["name"], result.Values[1]);
		Assert.Equal(ds.Rows[2]["position"] + "." + ds.Rows[2]["name"], result.Values[2]);
		Assert.Equal(ds.Rows[3]["position"] + "." + ds.Rows[3]["name"], result.Values[3]);
		Assert.Equal(ds.Rows[4]["position"] + "." + ds.Rows[4]["name"], result.Values[4]);
	}

	[Fact]
	public void TemplateProcessShouldReturnResultsIfTagCanBeProcessedMultiFixedIndex()
	{
		//Given
		string template = "{{position}}.{{name[0]}}";
		DataTable ds = SampleData;
		//When
		WvTemplateTagResultList result = WvTemplateUtility.ProcessTemplateTag(template, ds, DefaultCulture);
		//Then
		Assert.NotNull(result);
		Assert.Equal(5, result.Values.Count);
		Assert.Equal(ds.Rows[0]["position"] + "." + ds.Rows[0]["name"], result.Values[0]);
		Assert.Equal(ds.Rows[1]["position"] + "." + ds.Rows[0]["name"], result.Values[1]);
		Assert.Equal(ds.Rows[2]["position"] + "." + ds.Rows[0]["name"], result.Values[2]);
		Assert.Equal(ds.Rows[3]["position"] + "." + ds.Rows[0]["name"], result.Values[3]);
		Assert.Equal(ds.Rows[4]["position"] + "." + ds.Rows[0]["name"], result.Values[4]);
	}

	[Fact]
	public void TemplateProcessShouldReturnResultsIfTagCanBeProcessedMultiFixedIndexSingle()
	{
		//Given
		string template = "{{name[0]}}";
		DataTable ds = SampleData;
		//When
		WvTemplateTagResultList result = WvTemplateUtility.ProcessTemplateTag(template, ds, DefaultCulture);
		//Then
		Assert.NotNull(result);
		Assert.Single(result.Values);
		Assert.Equal((string)ds.Rows[0]["name"], result.Values[0]);
	}

	[Fact]
	public void ProcessValueShouldBeCorrect()
	{
		//Given
		var columnTypeNames = new List<string>{
			"short","int","long","number","date","datetime",
			"shorttext","text","guid"
		};
		var culture = new CultureInfo("en-US");
		var currentCulture = Thread.CurrentThread.CurrentCulture;
		var currentUICulture = Thread.CurrentThread.CurrentUICulture;
		try
		{
			Thread.CurrentThread.CurrentCulture = culture;
			Thread.CurrentThread.CurrentCulture = culture;
			foreach (var columnName in columnTypeNames)
			{
				string template = "{{" + columnName + "[0]}}";
				DataTable ds = TypedData;
				//When
				WvTemplateTagResultList result = WvTemplateUtility.ProcessTemplateTag(template, ds, culture);
				//Then
				Assert.NotNull(result);
				Assert.Single(result.Values);
				Assert.Equal(result.Values[0], ds.Rows[0][columnName]);
			}
		}
		finally
		{
			Thread.CurrentThread.CurrentCulture = currentCulture;
			Thread.CurrentThread.CurrentUICulture = currentUICulture;
		}
	}

	#endregion
}

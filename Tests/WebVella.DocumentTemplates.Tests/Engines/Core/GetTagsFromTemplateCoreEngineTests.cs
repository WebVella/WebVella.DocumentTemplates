using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Engines.SpreadsheetFile.Utility;
using WebVella.DocumentTemplates.Tests.Models;

namespace WebVella.DocumentTemplates.Tests.Engines;
public partial class GetTagsFromTemplateCoreEngineTests : TestBase
{
	private static readonly object locker = new();
	
	#region <<General>>	
	public GetTagsFromTemplateCoreEngineTests() : base() { }

	[Fact]
	public void GetApplicableRangeForFunctionTag_Test1()
	{
		lock (locker)
		{
			//Given
			var tags = new WvTemplateUtility().GetTagsFromTemplate("{{=SUM(A1)}}");
			Assert.NotNull(tags);
			Assert.Single(tags);
			//When
			var ranges = tags[0].GetApplicableRangeForFunctionTag();
			//Then
			Assert.NotNull(ranges);
			Assert.Single(ranges);
			var range = ranges[0];
			Assert.Equal(1, range.FirstRow);
			Assert.Equal(1, range.FirstColumn);
			Assert.Equal(1, range.LastRow);
			Assert.Equal(1, range.LastColumn);
		}
	}

	[Fact]
	public void GetApplicableRangeForFunctionTag_Test2()
	{
		lock (locker)
		{
			//Given
			var tags = new WvTemplateUtility().GetTagsFromTemplate("{{=SUM(A1:A1)}}");
			Assert.NotNull(tags);
			Assert.Single(tags);
			//When
			var ranges = tags[0].GetApplicableRangeForFunctionTag();
			//Then
			Assert.NotNull(ranges);
			Assert.Single(ranges);
			var range = ranges[0];
			Assert.Equal(1, range.FirstRow);
			Assert.Equal(1, range.FirstColumn);
			Assert.Equal(1, range.LastRow);
			Assert.Equal(1, range.LastColumn);
		}
	}

	[Fact]
	public void GetApplicableRangeForFunctionTag_Test3()
	{
		lock (locker)
		{
			//Given
			var tags = new WvTemplateUtility().GetTagsFromTemplate("{{=SUM(B1:C3)}}");
			Assert.NotNull(tags);
			Assert.Single(tags);
			//When
			var ranges = tags[0].GetApplicableRangeForFunctionTag();
			//Then
			Assert.NotNull(ranges);
			Assert.Single(ranges);
			var range = ranges[0];
			Assert.Equal(1, range.FirstRow);
			Assert.Equal(2, range.FirstColumn);
			Assert.Equal(3, range.LastRow);
			Assert.Equal(3, range.LastColumn);
		}
	}

	[Fact]
	public void NullTemplateShouldReturnNoTags()
	{
		//Given
		string? template = null;
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
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
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
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
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(columnName, result[0].Operator);
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
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(columnName.ToLowerInvariant(), result[0].Operator);
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
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template.Trim(), result[0].FullString);
		Assert.Equal(columnName.ToLowerInvariant(), result[0].Operator);
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
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(columnName, result[0].Operator);
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
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(columnName, result[0].Operator);
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
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal("{{" + columnName + "}}", result[0].FullString);
		Assert.Equal(columnName, result[0].Operator);
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
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal("{{" + columnName + "(\"test\")}}", result[0].FullString);
		Assert.Equal(columnName, result[0].Operator);
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
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(columnName, result[0].Operator);
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
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(columnName, result[0].Operator);
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
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(columnName, result[0].Operator);
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
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Equal(2, result.Count);
		var column1Tag = result.FirstOrDefault(x => x.Operator == columnName);
		var column2Tag = result.FirstOrDefault(x => x.Operator == columnName2);
		Assert.NotNull(column1Tag);
		Assert.NotNull(column2Tag);

		Assert.Equal("{{" + columnName + "}}", column1Tag.FullString);
		Assert.Equal(columnName, column1Tag.Operator);
		Assert.Equal(WvTemplateTagType.Data, column1Tag.Type);
		Assert.NotNull(column1Tag.IndexList);
		Assert.Empty(column1Tag.IndexList);
		Assert.NotNull(column1Tag.ParamGroups);
		Assert.Empty(column1Tag.ParamGroups);

		Assert.Equal("{{" + columnName2 + "}}", column2Tag.FullString);
		Assert.Equal(columnName2, column2Tag.Operator);
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
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Equal(2, result.Count);
		var column1Tag = result.FirstOrDefault(x => x.Operator == columnName);
		var column2Tag = result.FirstOrDefault(x => x.Operator == columnName2);
		Assert.NotNull(column1Tag);
		Assert.NotNull(column2Tag);

		Assert.Equal("{{" + columnName + "(" + columnNameParam1Name + "=\"" + columnNameParam1Value + "\", " + columnNameParam2Name + "=\"" + columnNameParam2Value + "\")}}", column1Tag.FullString);
		Assert.Equal(columnName, column1Tag.Operator);
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
		Assert.Equal(columnName2, column2Tag.Operator);
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
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Equal(2, result.Count);
		var column1Tag = result.FirstOrDefault(x => x.Operator == columnName);
		var column2Tag = result.FirstOrDefault(x => x.Operator == columnName2);
		Assert.NotNull(column1Tag);
		Assert.NotNull(column2Tag);

		Assert.Equal("{{" + columnName + "}}", column1Tag.FullString);
		Assert.Equal(columnName, column1Tag.Operator);
		Assert.Equal(WvTemplateTagType.Data, column1Tag.Type);
		Assert.Empty(column1Tag.ParamGroups);

		Assert.Equal("{{" + columnName2 + "}}", column2Tag.FullString);
		Assert.Equal(columnName2, column2Tag.Operator);
		Assert.Equal(WvTemplateTagType.Data, column2Tag.Type);
		Assert.Empty(column2Tag.ParamGroups);
	}
	
	[Fact]
	public void ExactTemplateShouldWithIndex()
	{
		//Given
		string template = "{{column_name[0]}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.NotNull(result[0]);

		Assert.Equal(WvTemplateTagType.Data, result[0].Type);
		Assert.Equal("column_name", result[0].ItemName);
		Assert.Empty(result[0].ParamGroups);
		Assert.Single(result[0].IndexList);
		Assert.Equal(0,result[0].IndexList[0]);
	}		

	[Fact]
	public void ExactTemplateShouldWith2Index()
	{
		//Given
		string template = "{{column_name[0,1]}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.NotNull(result[0]);

		Assert.Equal(WvTemplateTagType.Data, result[0].Type);
		Assert.Equal("column_name", result[0].ItemName);
		Assert.Empty(result[0].ParamGroups);
		Assert.Equal(2,result[0].IndexList.Count);
		Assert.Equal(0,result[0].IndexList[0]);
		Assert.Equal(1,result[0].IndexList[1]);
	}			
	
	[Fact]
	public void ExactTemplateShouldWith2Index2()
	{
		//Given
		string template = "{{column_name[0,1][2]}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.NotNull(result[0]);

		Assert.Equal(WvTemplateTagType.Data, result[0].Type);
		Assert.Equal("column_name", result[0].ItemName);
		Assert.Empty(result[0].ParamGroups);
		Assert.Equal(3,result[0].IndexList.Count);
		Assert.Equal(0,result[0].IndexList[0]);
		Assert.Equal(1,result[0].IndexList[1]);
		Assert.Equal(2,result[0].IndexList[2]);
	}		
	
	[Fact]
	public void ExactTemplateShouldWithDot()
	{
		//Given
		string template = "{{column_prefix.(S=',')}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.NotNull(result[0]);

		Assert.Equal(WvTemplateTagType.Data, result[0].Type);
		Assert.Equal("column_prefix.", result[0].ItemName);
		Assert.Single(result[0].ParamGroups);
	}	
	
	[Fact]
	public void ExactTemplateShouldReturnEmptyGroups()
	{
		//Given
		var columnName = "column_name";
		string template = "{{" + columnName + "()(\"test\")(test=1)}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		var column1Tag = result.FirstOrDefault(x => x.Operator == columnName);
		Assert.NotNull(column1Tag);
		Assert.Equal("{{" + columnName + "()(\"test\")(test=1)}}", column1Tag.FullString);
		Assert.Equal(columnName, column1Tag.Operator);
		Assert.Equal(WvTemplateTagType.Data, column1Tag.Type);
		Assert.NotNull(column1Tag.ParamGroups);
		Assert.Equal(2, column1Tag.ParamGroups.Count);

		Assert.Equal("(\"test\")", column1Tag.ParamGroups[0].FullString);
		Assert.Single(column1Tag.ParamGroups[0].Parameters);

		Assert.Equal("(test=1)", column1Tag.ParamGroups[1].FullString);
		Assert.Single(column1Tag.ParamGroups[1].Parameters);
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
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		var column1Tag = result.FirstOrDefault(x => x.Operator == columnName);
		Assert.NotNull(column1Tag);

		Assert.Equal("{{column_name(Param1=\"Param =1 Value\")(Param2=\"Param2Valu\")}}", column1Tag.FullString);
		Assert.Equal(columnName, column1Tag.Operator);
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
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(columnName, result[0].Operator);
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
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(columnName, result[0].Operator);
		Assert.Equal(WvTemplateTagType.Data, result[0].Type);
		Assert.NotNull(result[0].IndexList);
		Assert.Single(result[0].IndexList);
		Assert.Equal(columnIndex, result[0].IndexList[0]);
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
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(columnName, result[0].Operator);
		Assert.Equal(WvTemplateTagType.Data, result[0].Type);
		Assert.NotNull(result[0].IndexList);
		Assert.Equal(2, result[0].IndexList.Count);
		Assert.Equal(columnIndex, result[0].IndexList[0]);
		Assert.Equal(columnIndex2, result[0].IndexList[1]);
		Assert.Empty(result[0].ParamGroups);
	}

	[Fact]
	public void ExactTemplateWithMultiIndexingShouldReturnOneDataTagAndMultiIndex2()
	{
		//Given
		var columnName = "column_name";
		int columnIndex = 5;
		int columnIndex2 = 2;
		string template = "{{" + columnName + "[" + columnIndex + "," + columnIndex2 + "]}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(columnName, result[0].Operator);
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
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(columnName, result[0].Operator);
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
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(columnName, result[0].Operator);
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
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(columnName, result[0].Operator);
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
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(columnName, result[0].Operator);
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
	public void OnInTextTemplateShouldReturnOneDataTagWitGroupWithParamsWithQuotes3()
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
		string template = "{{" + columnName + "(" + paramName + "=" + paramValue + "," + paramName2 + "=\"" + paramValue2 + "\"," + paramName3 + "=\"" + paramValue3 + "\")}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(columnName, result[0].Operator);
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
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(function.ToLowerInvariant(), result[0].Operator);
		Assert.Equal(WvTemplateTagType.Function, result[0].Type);
		Assert.Equal(function.ToLowerInvariant(), result[0].ItemName);
		Assert.Single(result[0].ParamGroups);
		Assert.Single(result[0].ParamGroups[0].Parameters);
		Assert.Equal("A1:B1", result[0].ParamGroups[0].Parameters[0].ValueString);
	}

	[Fact]
	public void ExactTemplateWithSpreadsheetFunctionSupported()
	{
		//Given
		var function = "SUM";
		string template = "{{==" + function + "(A1:B1)}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(function.ToLowerInvariant(), result[0].Operator);
		Assert.Equal(WvTemplateTagType.SpreadsheetFunction, result[0].Type);
		Assert.Equal(function.ToLowerInvariant(), result[0].ItemName);
		Assert.Single(result[0].ParamGroups);
		Assert.Single(result[0].ParamGroups[0].Parameters);
		Assert.Equal("A1:B1", result[0].ParamGroups[0].Parameters[0].ValueString);
	}

	[Fact]
	public void ExactTemplateShouldReturnOneSpreadsheetFunctionTag()
	{
		//Given
		var functionName = "functionName";
		string template = "{{==" + functionName + "()}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(functionName.ToLowerInvariant(), result[0].Operator);
		Assert.Equal(WvTemplateTagType.SpreadsheetFunction, result[0].Type);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Empty(result[0].ParamGroups);
	}

	[Fact]
	public void ExactTemplateShouldReturnOneSpreadsheetFunctionTagWithMissingParams()
	{
		//Given
		var functionName = "functionName";
		string template = "{{==" + functionName + "}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(functionName.ToLowerInvariant(), result[0].Operator);
		Assert.Equal(WvTemplateTagType.SpreadsheetFunction, result[0].Type);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Empty(result[0].ParamGroups);
	}

	[Fact]
	public void ExactTemplateShouldReturnOneFunctionTag()
	{
		//Given
		var functionName = "functionName";
		string template = "{{=" + functionName + "()}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(functionName.ToLowerInvariant(), result[0].Operator);
		Assert.Equal(WvTemplateTagType.Function, result[0].Type);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Empty(result[0].ParamGroups);
	}

	[Fact]
	public void ExactTemplateShouldReturnOneFunctionTagWithMissingParams()
	{
		//Given
		var functionName = "functionName";
		string template = "{{=" + functionName + "}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(functionName.ToLowerInvariant(), result[0].Operator);
		Assert.Equal(WvTemplateTagType.Function, result[0].Type);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Empty(result[0].ParamGroups);

	}

	[Fact]
	public void ExactTemplateShouldReturnOneFunctionTagWithStaticRanges()
	{
		//Given
		var functionName = "sum";
		string template = "{{=" + functionName + "($A$1:B1)}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(functionName.ToLowerInvariant(), result[0].Operator);
		Assert.Equal(WvTemplateTagType.Function, result[0].Type);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Single(result[0].ParamGroups);
		Assert.Single(result[0].ParamGroups[0].Parameters);
		Assert.Equal("$A$1:B1", result[0].ParamGroups[0].Parameters[0].ValueString);

	}

	[Fact]
	public void ExactTemplateShouldReturnOneFunctionTagWithStaticRanges2()
	{
		//Given
		var functionName = "sum";
		string template = "{{=" + functionName + "($A$1:$B$1)}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(functionName.ToLowerInvariant(), result[0].Operator);
		Assert.Equal(WvTemplateTagType.Function, result[0].Type);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Single(result[0].ParamGroups);
		Assert.Single(result[0].ParamGroups[0].Parameters);
		Assert.Equal("$A$1:$B$1", result[0].ParamGroups[0].Parameters[0].ValueString);
	}

	[Fact]
	public void ExactTemplateShouldReturnOneFunctionTagWithStaticRanges3()
	{
		//Given
		var functionName = "sum";
		string template = "{{=" + functionName + "($A$1:$B$1,$A3:C$3)}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(functionName.ToLowerInvariant(), result[0].Operator);
		Assert.Equal(WvTemplateTagType.Function, result[0].Type);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Single(result[0].ParamGroups);
		Assert.Equal(2, result[0].ParamGroups[0].Parameters.Count);
		Assert.Equal("$A$1:$B$1", result[0].ParamGroups[0].Parameters[0].ValueString);
		Assert.Equal("$A3:C$3", result[0].ParamGroups[0].Parameters[1].ValueString);
	}
	
	[Fact]
	public void ExactTemplateMultipleSame()
	{
		//Given
		var columnName = "column_name";
		string template = "{{" + columnName + "(S=',')(S=';')[0][1]}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(columnName.ToLowerInvariant(), result[0].ItemName);
		Assert.Equal(WvTemplateTagType.Data, result[0].Type);
		Assert.NotNull(result[0].ParamGroups);
		Assert.NotNull(result[0].IndexList);
		Assert.Equal(2,result[0].ParamGroups.Count);
		Assert.Equal(2,result[0].IndexList.Count);

	}	
	#endregion
	
	#region <<Inline template>>
	[Fact]
	public void ExactTemplateShouldReturnOneInlineStartTag()
	{
		//Given
		string template = "{{<#}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(WvTemplateTagType.InlineStart, result[0].Type);
		Assert.Equal(String.Empty, result[0].ItemName);
		Assert.Equal("<#", result[0].Operator);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Empty(result[0].ParamGroups);
	}	
	
	[Fact]
	public void ExactTemplateShouldReturnOneInlineEndTag()
	{
		//Given
		string template = "{{#>}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(WvTemplateTagType.InlineEnd, result[0].Type);
		Assert.Equal(String.Empty, result[0].ItemName);
		Assert.Equal("#>", result[0].Operator);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Empty(result[0].ParamGroups);
	}		
	[Fact]
	public void ExactTemplateShouldReturnTwoInlineTags()
	{
		//Given
		string template = "{{<#}}{{#>}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Equal(2,result.Count);
		Assert.Equal(WvTemplateTagType.InlineStart, result[0].Type);
		Assert.Equal(String.Empty, result[0].ItemName);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Empty(result[0].ParamGroups);
		Assert.Equal(WvTemplateTagType.InlineEnd, result[1].Type);
		Assert.Equal(String.Empty, result[1].ItemName);
		Assert.NotNull(result[1].ParamGroups);
		Assert.Empty(result[1].ParamGroups);		
	}	
	
	[Fact]
	public void ExactTemplateShouldReturnOneStartTagWithParam()
	{
		//Given
		string template = "{{<#(F=V)}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(WvTemplateTagType.InlineStart, result[0].Type);
		Assert.Equal(String.Empty, result[0].ItemName);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Single(result[0].ParamGroups);
		Assert.Single(result[0].ParamGroups[0].Parameters);
		Assert.IsType<WvTemplateTagDataFlowParameterProcessor>(result[0].ParamGroups[0].Parameters[0]);
	}		
	[Fact]
	public void ExactTemplateShouldReturnOneEndTagWithParam()
	{
		//Given
		string template = "{{#>(F=V)}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(WvTemplateTagType.InlineEnd, result[0].Type);
		Assert.Equal(String.Empty, result[0].ItemName);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Single(result[0].ParamGroups);
		Assert.Single(result[0].ParamGroups[0].Parameters);
		Assert.IsType<WvTemplateTagDataFlowParameterProcessor>(result[0].ParamGroups[0].Parameters[0]);
	}		
	
	[Fact]
	public void ExactTemplateShouldReturnOneEndTagWithIndex()
	{
		//Given
		string template = "{{<#[0](F=V)}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(WvTemplateTagType.InlineStart, result[0].Type);
		Assert.Equal(String.Empty, result[0].ItemName);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Single(result[0].ParamGroups);
		Assert.Single(result[0].ParamGroups[0].Parameters);
		Assert.NotNull(result[0].IndexList);
		Assert.Single(result[0].IndexList);
		Assert.Equal(0,result[0].IndexList[0]);
		Assert.IsType<WvTemplateTagDataFlowParameterProcessor>(result[0].ParamGroups[0].Parameters[0]);
	}			
	[Fact]
	public void ExactTemplateShouldReturnOneEndTagWithIndex2()
	{
		//Given
		string template = "{{<#[0,1]}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		//Assert.Equal(WvTemplateTagType.InlineStart, result[0].Type);
		Assert.Equal(String.Empty, result[0].ItemName);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Empty(result[0].ParamGroups);
		Assert.NotNull(result[0].IndexList);
		Assert.Equal(2,result[0].IndexList.Count);
		Assert.Equal(0,result[0].IndexList[0]);
		Assert.Equal(1,result[0].IndexList[1]);

	}		
	
	[Fact]
	public void ExactTemplateShouldReturnTheColumNameToo()
	{
		//Given
		string template = "{{<#column_name[0,1]}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(WvTemplateTagType.InlineStart, result[0].Type);
		Assert.Equal("<#", result[0].Operator);
		Assert.Equal("column_name", result[0].ItemName);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Empty(result[0].ParamGroups);
		Assert.NotNull(result[0].IndexList);
		Assert.Equal(2,result[0].IndexList.Count);
		Assert.Equal(0,result[0].IndexList[0]);
		Assert.Equal(1,result[0].IndexList[1]);
	}	
	
	#endregion
	
	#region <<Conditional>>
	[Fact]
	public void ExactTemplateShouldReturnOneConditionalStartTag()
	{
		//Given
		string template = "{{<?}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(WvTemplateTagType.ConditionStart, result[0].Type);
		Assert.Equal("<?", result[0].Operator);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Empty(result[0].ParamGroups);
	}	
	
	[Fact]
	public void ExactTemplateShouldReturnOneConditionalEndTag()
	{
		//Given
		string template = "{{?>}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(WvTemplateTagType.ConditionEnd, result[0].Type);
		Assert.Equal("?>", result[0].Operator);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Empty(result[0].ParamGroups);
	}		
	[Fact]
	public void ExactTemplateShouldReturnTwoConditionalTags()
	{
		//Given
		string template = "{{<?}}{{?>}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Equal(2,result.Count);
		Assert.Equal(WvTemplateTagType.ConditionStart, result[0].Type);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Empty(result[0].ParamGroups);
		Assert.Equal(WvTemplateTagType.ConditionEnd, result[1].Type);
		Assert.NotNull(result[1].ParamGroups);
		Assert.Empty(result[1].ParamGroups);		
	}	
	
	[Fact]
	public void ExactTemplateShouldReturnConditionalOneStartTagWithParam()
	{
		//Given
		string template = "{{<?(column_name = 'test')}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(WvTemplateTagType.ConditionStart, result[0].Type);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Single(result[0].ParamGroups);
		Assert.Single(result[0].ParamGroups[0].Parameters);
		Assert.IsType<WvTemplateTagUnknownParameterProcessor>(result[0].ParamGroups[0].Parameters[0]);
	}		
	[Fact]
	public void ExactTemplateShouldReturnConditionalOneEndTagWithParam()
	{
		//Given
		string template = "{{?>(column_name = 'test')}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(WvTemplateTagType.ConditionEnd, result[0].Type);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Single(result[0].ParamGroups);
		Assert.Single(result[0].ParamGroups[0].Parameters);
		Assert.IsType<WvTemplateTagUnknownParameterProcessor>(result[0].ParamGroups[0].Parameters[0]);
	}		
	
	[Fact]
	public void ExactTemplateShouldReturnConditionalOneEndTagWithIndex()
	{
		//Given
		string template = "{{<?[0](column_name >= 'test')}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(WvTemplateTagType.ConditionStart, result[0].Type);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Single(result[0].ParamGroups);
		Assert.Single(result[0].ParamGroups[0].Parameters);
		Assert.NotNull(result[0].IndexList);
		Assert.Single(result[0].IndexList);
		Assert.Equal(0,result[0].IndexList[0]);
		Assert.IsType<WvTemplateTagUnknownParameterProcessor>(result[0].ParamGroups[0].Parameters[0]);
	}			
	[Fact]
	public void ExactTemplateShouldReturnConditionalOneEndTagWithIndex3()
	{
		//Given
		string template = "{{<?[0](column_name !*= 'test')}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(WvTemplateTagType.ConditionStart, result[0].Type);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Single(result[0].ParamGroups);
		Assert.Single(result[0].ParamGroups[0].Parameters);
		Assert.NotNull(result[0].IndexList);
		Assert.Single(result[0].IndexList);
		Assert.Equal(0,result[0].IndexList[0]);
		Assert.IsType<WvTemplateTagUnknownParameterProcessor>(result[0].ParamGroups[0].Parameters[0]);
	}			
	[Fact]
	public void ExactTemplateShouldReturnConditionalOneEndTagWithIndex2()
	{
		//Given
		string template = "{{<?[0,1]}}";
		//When
		List<WvTemplateTag> result = new WvTemplateUtility().GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		//Assert.Equal(WvTemplateTagType.ConditionStart, result[0].Type);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Empty(result[0].ParamGroups);
		Assert.NotNull(result[0].IndexList);
		Assert.Equal(2,result[0].IndexList.Count);
		Assert.Equal(0,result[0].IndexList[0]);
		Assert.Equal(1,result[0].IndexList[1]);

	}			
	#endregion
}

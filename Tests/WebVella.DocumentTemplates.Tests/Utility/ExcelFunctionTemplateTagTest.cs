using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Tests.Models;

namespace WebVella.DocumentTemplates.Tests.Utility;
public class ExcelFunctionTemplateTagTest : TestBase
{
	public ExcelFunctionTemplateTagTest() : base() { }

	[Fact]
	public void ExactTemplateShouldReturnOneExcelFunctionTag()
	{
		//Given
		var functionName = "functionName";
		string template = "{{==" + functionName + "()}}";
		//When
		List<WvTemplateTag> result = WvTemplateUtility.GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(functionName.ToLowerInvariant(), result[0].Name);
		Assert.Equal(WvTemplateTagType.ExcelFunction, result[0].Type);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Single(result[0].ParamGroups);
	}

	[Fact]
	public void ExactTemplateShouldReturnOneExcelFunctionTagWithMissingParams()
	{
		//Given
		var functionName = "functionName";
		string template = "{{==" + functionName + "}}";
		//When
		List<WvTemplateTag> result = WvTemplateUtility.GetTagsFromTemplate(template);
		//Then
		Assert.NotNull(result);
		Assert.Single(result);
		Assert.Equal(template, result[0].FullString);
		Assert.Equal(functionName.ToLowerInvariant(), result[0].Name);
		Assert.Equal(WvTemplateTagType.ExcelFunction, result[0].Type);
		Assert.NotNull(result[0].ParamGroups);
		Assert.Empty(result[0].ParamGroups);
	}
}

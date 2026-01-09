using System.Data;
using System.Globalization;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Extensions;
using WebVella.DocumentTemplates.Tests.Models;

namespace WebVella.DocumentTemplates.Tests.Engines;
public partial class ExtractTagParametersFromGroupTests : TestBase
{
	public ExtractTagParametersFromGroupTests() : base() { }

	[Fact]
	public void ConditionEmpty()
	{
		//Given
		string paramGroup = "";
		var tagType = WvTemplateTagType.ConditionStart;
		//When
		IWvTemplateTagParameterProcessorBase? result = new WvTemplateUtility().ExtractTagParameterFromDefinition(paramGroup, tagType);
		//Then
		Assert.Null(result);
	}
	
	[Fact]
	public void ConditionEqual()
	{
		//Given
		string columnName = "column_name";
		var operatorType = WvTemplateTagParamOperatorType.Equals;
		string value = "1";
		string paramGroup = $"{columnName}{operatorType.ToDescriptionString()}{value}";
		var tagType = WvTemplateTagType.ConditionStart;
		//When
		IWvTemplateTagParameterProcessorBase? result = new WvTemplateUtility().ExtractTagParameterFromDefinition(paramGroup, tagType);
		//Then
		Assert.NotNull(result);
		Assert.IsType<WvTemplateTagUnknownParameterProcessor>(result);
		Assert.Equal(operatorType,result.OperatorType);
		Assert.Equal(columnName,result.Name);
		Assert.Equal(value,result.ValueString);
	}	
	
	[Fact]
	public void ConditionEqualSpaces()
	{
		//Given
		string columnName = "column_name";
		var operatorType = WvTemplateTagParamOperatorType.Equals;
		string value = "1";
		string paramGroup = $"{columnName} {operatorType.ToDescriptionString()} {value}";
		var tagType = WvTemplateTagType.ConditionStart;
		//When
		IWvTemplateTagParameterProcessorBase? result = new WvTemplateUtility().ExtractTagParameterFromDefinition(paramGroup, tagType);
		//Then
		Assert.NotNull(result);
		Assert.IsType<WvTemplateTagUnknownParameterProcessor>(result);
		Assert.Equal(operatorType,result.OperatorType);
		Assert.Equal(columnName,result.Name);
		Assert.Equal(value,result.ValueString);
	}		
	
	[Fact]
	public void ConditionEqualNotEqual()
	{
		//Given
		string columnName = "column_name";
		var operatorType = WvTemplateTagParamOperatorType.NotEquals;
		string value = "1";
		string paramGroup = $"{columnName} {operatorType.ToDescriptionString()} {value}";
		var tagType = WvTemplateTagType.ConditionStart;
		//When
		IWvTemplateTagParameterProcessorBase? result = new WvTemplateUtility().ExtractTagParameterFromDefinition(paramGroup, tagType);
		//Then
		Assert.NotNull(result);
		Assert.IsType<WvTemplateTagUnknownParameterProcessor>(result);
		Assert.Equal(operatorType,result.OperatorType);
		Assert.Equal(columnName,result.Name);
		Assert.Equal(value,result.ValueString);
	}		
	
	[Fact]
	public void ConditionEqualGreater()
	{
		//Given
		string columnName = "column_name";
		var operatorType = WvTemplateTagParamOperatorType.GreaterThan;
		string value = "1";
		string paramGroup = $"{columnName} {operatorType.ToDescriptionString()} {value}";
		var tagType = WvTemplateTagType.ConditionStart;
		//When
		IWvTemplateTagParameterProcessorBase? result = new WvTemplateUtility().ExtractTagParameterFromDefinition(paramGroup, tagType);
		//Then
		Assert.NotNull(result);
		Assert.IsType<WvTemplateTagUnknownParameterProcessor>(result);
		Assert.Equal(operatorType,result.OperatorType);
		Assert.Equal(columnName,result.Name);
		Assert.Equal(value,result.ValueString);
	}		
	
	[Fact]
	public void ConditionEqualGreaterOrEqual()
	{
		//Given
		string columnName = "column_name";
		var operatorType = WvTemplateTagParamOperatorType.GreaterOrEqualThan;
		string value = "1";
		string paramGroup = $"{columnName} {operatorType.ToDescriptionString()} {value}";
		var tagType = WvTemplateTagType.ConditionStart;
		//When
		IWvTemplateTagParameterProcessorBase? result = new WvTemplateUtility().ExtractTagParameterFromDefinition(paramGroup, tagType);
		//Then
		Assert.NotNull(result);
		Assert.IsType<WvTemplateTagUnknownParameterProcessor>(result);
		Assert.Equal(operatorType,result.OperatorType);
		Assert.Equal(columnName,result.Name);
		Assert.Equal(value,result.ValueString);
	}		
	[Fact]
	public void ConditionEqualLess()
	{
		//Given
		string columnName = "column_name";
		var operatorType = WvTemplateTagParamOperatorType.LessThan;
		string value = "1";
		string paramGroup = $"{columnName} {operatorType.ToDescriptionString()} {value}";
		var tagType = WvTemplateTagType.ConditionStart;
		//When
		IWvTemplateTagParameterProcessorBase? result = new WvTemplateUtility().ExtractTagParameterFromDefinition(paramGroup, tagType);
		//Then
		Assert.NotNull(result);
		Assert.IsType<WvTemplateTagUnknownParameterProcessor>(result);
		Assert.Equal(operatorType,result.OperatorType);
		Assert.Equal(columnName,result.Name);
		Assert.Equal(value,result.ValueString);
	}			
	
	[Fact]
	public void ConditionEqualLessOrEqual()
	{
		//Given
		string columnName = "column_name";
		var operatorType = WvTemplateTagParamOperatorType.LessOrEqualThan;
		string value = "1";
		string paramGroup = $"{columnName} {operatorType.ToDescriptionString()} {value}";
		var tagType = WvTemplateTagType.ConditionStart;
		//When
		IWvTemplateTagParameterProcessorBase? result = new WvTemplateUtility().ExtractTagParameterFromDefinition(paramGroup, tagType);
		//Then
		Assert.NotNull(result);
		Assert.IsType<WvTemplateTagUnknownParameterProcessor>(result);
		Assert.Equal(operatorType,result.OperatorType);
		Assert.Equal(columnName,result.Name);
		Assert.Equal(value,result.ValueString);
	}	
	
	
	
	
	[Fact]
	public void ConditionEqualContains()
	{
		//Given
		string columnName = "column_name";
		var operatorType = WvTemplateTagParamOperatorType.Contains;
		string value = "1";
		string paramGroup = $"{columnName} {operatorType.ToDescriptionString()} {value}";
		var tagType = WvTemplateTagType.ConditionStart;
		//When
		IWvTemplateTagParameterProcessorBase? result = new WvTemplateUtility().ExtractTagParameterFromDefinition(paramGroup, tagType);
		//Then
		Assert.NotNull(result);
		Assert.IsType<WvTemplateTagUnknownParameterProcessor>(result);
		Assert.Equal(operatorType,result.OperatorType);
		Assert.Equal(columnName,result.Name);
		Assert.Equal(value,result.ValueString);
	}		
	
	[Fact]
	public void ConditionEqualNotContains()
	{
		//Given
		string columnName = "column_name";
		var operatorType = WvTemplateTagParamOperatorType.NotContains;
		string value = "1";
		string paramGroup = $"{columnName} {operatorType.ToDescriptionString()} {value}";
		var tagType = WvTemplateTagType.ConditionStart;
		//When
		IWvTemplateTagParameterProcessorBase? result = new WvTemplateUtility().ExtractTagParameterFromDefinition(paramGroup, tagType);
		//Then
		Assert.NotNull(result);
		Assert.IsType<WvTemplateTagUnknownParameterProcessor>(result);
		Assert.Equal(operatorType,result.OperatorType);
		Assert.Equal(columnName,result.Name);
		Assert.Equal(value,result.ValueString);
	}		

	
	[Fact]
	public void ConditionEqualStartsWith()
	{
		//Given
		string columnName = "column_name";
		var operatorType = WvTemplateTagParamOperatorType.StartsWith;
		string value = "1";
		string paramGroup = $"{columnName} {operatorType.ToDescriptionString()} {value}";
		var tagType = WvTemplateTagType.ConditionStart;
		//When
		IWvTemplateTagParameterProcessorBase? result = new WvTemplateUtility().ExtractTagParameterFromDefinition(paramGroup, tagType);
		//Then
		Assert.NotNull(result);
		Assert.IsType<WvTemplateTagUnknownParameterProcessor>(result);
		Assert.Equal(operatorType,result.OperatorType);
		Assert.Equal(columnName,result.Name);
		Assert.Equal(value,result.ValueString);
	}		
	
	[Fact]
	public void ConditionEqualNotStartsWith()
	{
		//Given
		string columnName = "column_name";
		var operatorType = WvTemplateTagParamOperatorType.NotStartsWith;
		string value = "1";
		string paramGroup = $"{columnName} {operatorType.ToDescriptionString()} {value}";
		var tagType = WvTemplateTagType.ConditionStart;
		//When
		IWvTemplateTagParameterProcessorBase? result = new WvTemplateUtility().ExtractTagParameterFromDefinition(paramGroup, tagType);
		//Then
		Assert.NotNull(result);
		Assert.IsType<WvTemplateTagUnknownParameterProcessor>(result);
		Assert.Equal(operatorType,result.OperatorType);
		Assert.Equal(columnName,result.Name);
		Assert.Equal(value,result.ValueString);
	}			
	[Fact]
	public void ConditionEqualEndsWith()
	{
		//Given
		string columnName = "column_name";
		var operatorType = WvTemplateTagParamOperatorType.EndsWith;
		string value = "1";
		string paramGroup = $"{columnName} {operatorType.ToDescriptionString()} {value}";
		var tagType = WvTemplateTagType.ConditionStart;
		//When
		IWvTemplateTagParameterProcessorBase? result = new WvTemplateUtility().ExtractTagParameterFromDefinition(paramGroup, tagType);
		//Then
		Assert.NotNull(result);
		Assert.IsType<WvTemplateTagUnknownParameterProcessor>(result);
		Assert.Equal(operatorType,result.OperatorType);
		Assert.Equal(columnName,result.Name);
		Assert.Equal(value,result.ValueString);
	}		
	[Fact]
	public void ConditionEqualNotEndsWith()
	{
		//Given
		string columnName = "column_name";
		var operatorType = WvTemplateTagParamOperatorType.NotEndsWith;
		string value = "1";
		string paramGroup = $"{columnName} {operatorType.ToDescriptionString()} {value}";
		var tagType = WvTemplateTagType.ConditionStart;
		//When
		IWvTemplateTagParameterProcessorBase? result = new WvTemplateUtility().ExtractTagParameterFromDefinition(paramGroup, tagType);
		//Then
		Assert.NotNull(result);
		Assert.IsType<WvTemplateTagUnknownParameterProcessor>(result);
		Assert.Equal(operatorType,result.OperatorType);
		Assert.Equal(columnName,result.Name);
		Assert.Equal(value,result.ValueString);
	}		
}

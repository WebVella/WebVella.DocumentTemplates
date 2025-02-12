using System.Text.RegularExpressions;

namespace WebVella.DocumentTemplates.Core.Utility;
public static partial class WvTemplateUtility
{
public static IWvTemplateTagParameterBase? ExtractTagParameterFromDefinition(string parameterDefinition, WvTemplateTagType tagType)
	{
		if (String.IsNullOrWhiteSpace(parameterDefinition)) return null;
		string? paramName = null;
		string? paramValue = null;
		var firstEqualSignIndex = parameterDefinition.IndexOf('=');
		//if not a named parameter
		if (firstEqualSignIndex == -1)
		{
			paramValue = parameterDefinition;
		}
		else
		{
			paramName = parameterDefinition.Substring(0, firstEqualSignIndex);
			paramValue = parameterDefinition.Substring(firstEqualSignIndex + 1); //Remove the =
		}

		paramName = paramName?.Trim()?.ToLowerInvariant(); //names are always lowered
		paramValue = paramValue?.Trim();

		//Check if it is a string value
		if (
			!String.IsNullOrWhiteSpace(paramValue)
			&& ((paramValue.StartsWith("\"") && paramValue.EndsWith("\"")) || (paramValue.StartsWith("'") && paramValue.EndsWith("'")))
		)
		{
			paramValue = paramValue.Remove(0, 1);
			paramValue = paramValue.Substring(0, paramValue.Length - 1);
		}

		if (tagType == WvTemplateTagType.Data)
		{
			switch(paramName){ 
				case "f":
					return new WvTemplateTagDataFlowParameter(paramValue);
				case "s":
					return new WvTemplateTagSeparatorParameter(paramValue);
				default: 
					return new WvTemplateTagUnknownParameter(paramName, paramValue);;
			}
		}
		else if (tagType == WvTemplateTagType.Function)
		{
		}
		else if (tagType == WvTemplateTagType.ExcelFunction)
		{
		}
		return new WvTemplateTagUnknownParameter(paramName, paramValue);
	}
}

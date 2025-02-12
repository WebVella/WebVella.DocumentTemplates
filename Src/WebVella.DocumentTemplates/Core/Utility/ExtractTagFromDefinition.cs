using System.Text.RegularExpressions;

namespace WebVella.DocumentTemplates.Core.Utility;
public partial class WvTemplateUtility
{
	private static List<string> _supportedFunctions = new List<string> { "sum" };
	private static List<string> _supportedExcelFunctions = new List<string> { "sum" };
	public static WvTemplateTag? ExtractTagFromDefinition(string? tagDefinition)
	{
		if (String.IsNullOrWhiteSpace(tagDefinition)
		|| !tagDefinition.StartsWith("{{")
		|| !tagDefinition.EndsWith("}}")) return null;
		WvTemplateTag result = new();
		//Remove leading and end {{}}
		var processedDefinition = tagDefinition.Remove(0, 2);
		processedDefinition = processedDefinition.Substring(0, processedDefinition.Length - 2);
		processedDefinition = processedDefinition?.Trim();
		if (String.IsNullOrWhiteSpace(processedDefinition)) return null;
		result.FullString = tagDefinition;
		//Excel function
		if (processedDefinition.StartsWith("=="))
		{
			result.Type = WvTemplateTagType.ExcelFunction;
			processedDefinition = processedDefinition.Remove(0, 2);
		}
		//Function
		else if (processedDefinition.StartsWith("="))
		{
			result.Type = WvTemplateTagType.Function;
			processedDefinition = processedDefinition.Remove(0, 1);
		}
		//Data
		else
		{
			result.Type = WvTemplateTagType.Data;

		}

		//PARAMETERS
		foreach (Match matchGroup in Regex.Matches(processedDefinition, @"(\([^()]*\))"))
		{
			//Replace the first occurance of the group string
			var regex = new Regex(Regex.Escape(matchGroup.Value));
			processedDefinition = regex.Replace(processedDefinition, "", 1);

			var group = new WvTemplateTagParamGroup
			{
				FullString = matchGroup.Value,
				Parameters = ExtractTagParametersFromGroup(matchGroup.Value, result.Type)
			};
			result.ParamGroups.Add(group);
		}
		processedDefinition = processedDefinition?.Trim();
		//if (String.IsNullOrWhiteSpace(processedDefinition)) return null;

		//INDEX
		foreach (Match matchGroup in Regex.Matches((processedDefinition ?? String.Empty), @"(\[[^[]]*\]|\[\s*\])"))
		{
			//Replace the first occurance of the group string
			var regex = new Regex(Regex.Escape(matchGroup.Value));
			processedDefinition = regex.Replace((processedDefinition ?? String.Empty), "", 1);

			var tagIndex = ExtractTagIndexFromGroup(matchGroup.Value);
			if (tagIndex is null || tagIndex < 0)
			{
				//interpred [] or any [invalid] as 0
				result.IndexList.Add(0);
			}
			else
			{
				result.IndexList.Add(tagIndex.Value);
			}
		}

		processedDefinition = processedDefinition?.Trim();
		//if (String.IsNullOrWhiteSpace(processedDefinition)) return null;
		result.Name = (processedDefinition ?? String.Empty).ToLowerInvariant();

		if (result.Type == WvTemplateTagType.Function)
		{
			if (!String.IsNullOrWhiteSpace(result.Name) && _supportedFunctions.Contains(result.Name))
			{
				result.FunctionName = result.Name;
			}
		}
		if (result.Type == WvTemplateTagType.ExcelFunction)
		{
			if (!String.IsNullOrWhiteSpace(result.Name) && _supportedExcelFunctions.Contains(result.Name))
			{
				result.FunctionName = result.Name;
			}
		}
		return result;

	}
}

using System.Text.RegularExpressions;

namespace WebVella.DocumentTemplates.Core.Utility;
public partial class WvTemplateUtility
{
	public WvTemplateTag? ExtractTagFromDefinition(string? tagDefinition)
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
		//spreadsheet function
		if (processedDefinition.StartsWith("=="))
		{
			result.Type = WvTemplateTagType.SpreadsheetFunction;
			processedDefinition = processedDefinition.Remove(0, 2);
		}
		//Function
		else if (processedDefinition.StartsWith("="))
		{
			result.Type = WvTemplateTagType.Function;
			processedDefinition = processedDefinition.Remove(0, 1);
		}
		//Conditional start
		else if (processedDefinition.StartsWith("<?"))
		{
			result.Type = WvTemplateTagType.ConditionStart;
		}
		//Conditional end
		else if (processedDefinition.StartsWith("?>"))
		{
			result.Type = WvTemplateTagType.ConditionEnd;
		}			
        //Inline start
        else if (processedDefinition.StartsWith("<#"))
        {
            result.Type = WvTemplateTagType.InlineStart;
        }
        //Inline end
        else if (processedDefinition.StartsWith("#>"))
        {
	        result.Type = WvTemplateTagType.InlineEnd;
        }		
	
        //Data
        else
		{
			result.Type = WvTemplateTagType.Data;

		}

		//PARAMETERS
		//remove empty groups
		processedDefinition = processedDefinition.Replace("()", String.Empty);
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
		//Remove empty indices
		processedDefinition = processedDefinition.Replace("[]", String.Empty);
		//OLD  @"(\[[^[]]*\]|\[\s*\])";
		//Pattern matches multiple groups as "[2]","[1,2]","[1][2]"
		var indexMatchPattern = @"\[\s*\d+(?:\s*,\s*\d+)*\s*\]";
		foreach (Match matchGroup in Regex.Matches((processedDefinition ?? String.Empty), indexMatchPattern))
		{
			//Replace the first occurance of the group string
			var regex = new Regex(Regex.Escape(matchGroup.Value));
			processedDefinition = regex.Replace((processedDefinition ?? String.Empty), "", 1);

			var indices = ExtractTagIndexFromGroup(matchGroup.Value);
			if (indices.Count == 0)
			{
				//skip improper indexes
			}
			else
			{
				result.IndexList.AddRange(indices);
			}
		}

		processedDefinition = processedDefinition?.Trim();
		//if (String.IsNullOrWhiteSpace(processedDefinition)) return null;
		result.Name = (processedDefinition ?? String.Empty).ToLowerInvariant();

		if (result.Type == WvTemplateTagType.Function)
		{
			if (!String.IsNullOrWhiteSpace(result.Name))
			{
				result.FunctionName = result.Name;
			}
		}
		if (result.Type == WvTemplateTagType.SpreadsheetFunction)
		{
			if (!String.IsNullOrWhiteSpace(result.Name))
			{
				result.FunctionName = result.Name;
			}
		}
		return result;

	}
}

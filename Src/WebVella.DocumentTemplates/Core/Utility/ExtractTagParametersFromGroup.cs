using System.Text.RegularExpressions;

namespace WebVella.DocumentTemplates.Core.Utility;
public partial class WvTemplateUtility
{
	public List<IWvTemplateTagParameterProcessorBase> ExtractTagParametersFromGroup(string parameterGroup, WvTemplateTagType tagType)
	{
		var result = new List<IWvTemplateTagParameterProcessorBase>();
		if (String.IsNullOrWhiteSpace(parameterGroup)
		|| !parameterGroup.StartsWith("(")
		|| !parameterGroup.EndsWith(")")) return result;
		//Remove leading and end ()
		var processedParameterGroup = parameterGroup.Remove(0, 1);
		processedParameterGroup = processedParameterGroup.Substring(0, processedParameterGroup.Length - 1);
		processedParameterGroup = processedParameterGroup?.Trim();
		if (String.IsNullOrWhiteSpace(processedParameterGroup)) return result;

		//string pattern = @"\w+(?:\s*=\s*('[^']*'|\"".*?\""|[^,]*))?";
		string pattern = @"[\w:$]+(?:\s*=\s*('[^']*'|\"".*?\""|[^,]*))?";
		Regex regex = new Regex(pattern);
		MatchCollection matches = regex.Matches(processedParameterGroup);
		foreach (Match match in matches)
		{
			var parameter = ExtractTagParameterFromDefinition(match.Value, tagType);
			if (parameter is not null) result.Add(parameter);
		}


		return result;
	}
}

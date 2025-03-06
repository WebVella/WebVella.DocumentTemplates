using System.Text.RegularExpressions;
using WebVella.DocumentTemplates.Core.Services;

namespace WebVella.DocumentTemplates.Core.Utility;
public partial class WvTemplateUtility
{
	public IWvTemplateTagParameterProcessorBase? ExtractTagParameterFromDefinition(string parameterDefinition, WvTemplateTagType tagType)
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
			&& ((paramValue.StartsWith("\"") && paramValue.EndsWith("\"")) 
				|| (paramValue.StartsWith("'") && paramValue.EndsWith("'"))
				|| (paramValue.StartsWith("'") && paramValue.EndsWith("'"))
				|| (paramValue.StartsWith("’") && paramValue.EndsWith("’"))
				|| (paramValue.StartsWith("”") && paramValue.EndsWith("”"))
				
				)
		)
		{
			paramValue = paramValue.Remove(0, 1);
			paramValue = paramValue.Substring(0, paramValue.Length - 1);
		}

		if (!String.IsNullOrWhiteSpace(paramName))
		{
			var paramProcessorType = new WvCoreMetaService().GetRegisteredTagParameterProcessorByName(paramName);
			if (paramProcessorType is null) return new WvTemplateTagUnknownParameterProcessor(paramName, paramValue);

			return Activator.CreateInstance(paramProcessorType, paramValue) as IWvTemplateTagParameterProcessorBase;
		}
		return new WvTemplateTagUnknownParameterProcessor(paramName, paramValue);
	}
}

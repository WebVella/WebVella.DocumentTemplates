using System.Text.RegularExpressions;

namespace WebVella.DocumentTemplates.Core.Utility;
public partial class WvTemplateUtility
{
	public List<WvTemplateTag> GetTagsFromTemplate(string? text)
	{
		var result = new List<WvTemplateTag>();
		if(String.IsNullOrEmpty(text)) return result;
		foreach (Match match in Regex.Matches(text, @"(\{\{[^{}]*\}\})"))
		{
			var tag = ExtractTagFromDefinition(match.Value);
			if (tag is not null) result.Add(tag);
		}
		return result;
	}
}

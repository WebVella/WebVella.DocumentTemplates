using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;

namespace WebVella.DocumentTemplates.Core.Utility;
public partial class WvTemplateUtility
{
	public static WvTemplateTagResultList ProcessTemplateTag(string? template, DataTable dataSource, CultureInfo culture)
	{
		var result = new WvTemplateTagResultList();
		result.Tags = GetTagsFromTemplate(template);
		//if there are no tags - return one with the template
		if (result.Tags.Count == 0)
		{
			result.Values.Add(template ?? String.Empty);
			return result;
		}
		//if all tags are index - return one with processed template
		if (!result.IsMultiValueTemplate)
		{
			var resultValue = GenerateTemplateTagResult(template, result.Tags, dataSource, null, culture);
			if (resultValue is not null && resultValue.Value is not null)
			{
				result.Values.Add(resultValue.Value);
			}
			return result;
		}

		for (int i = 0; i < dataSource.Rows.Count; i++)
		{
			var resultValue = GenerateTemplateTagResult(template, result.Tags, dataSource, i, culture);
			if (resultValue is not null && resultValue.Value is not null)
			{
				result.Values.Add(resultValue.Value);
			}
		}
		return result;
	}

}

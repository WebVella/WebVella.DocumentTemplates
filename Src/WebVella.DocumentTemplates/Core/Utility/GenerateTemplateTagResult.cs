using System.Data;
using System.Globalization;

namespace WebVella.DocumentTemplates.Core.Utility;
public partial class WvTemplateUtility
{
	public WvTemplateTagResult GenerateTemplateTagResult(string? template, List<WvTemplateTag> tags, DataTable dataSource, int? rowIndex, CultureInfo culture)
	{
		var result = new WvTemplateTagResult();
		if (String.IsNullOrEmpty(template)) return result;
		result.ValueString = template;
		if (String.IsNullOrWhiteSpace(template)) return result;
		result.Tags = tags;
		foreach (var tag in result.Tags)
		{
			(result.ValueString, result.Value) = ProcessTagInTemplate(result.ValueString, result.Value, tag, dataSource, rowIndex, culture);
		}

		return result;
	}


}

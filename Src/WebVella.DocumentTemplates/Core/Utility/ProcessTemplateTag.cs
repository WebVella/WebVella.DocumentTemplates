using System.Data;
using System.Globalization;
using System.Text;

namespace WebVella.DocumentTemplates.Core.Utility;
public partial class WvTemplateUtility
{

	public WvTemplateTagResultList ProcessTemplateTag(
		string? template,
		DataTable dataSource,
		CultureInfo culture)
	{
	
		var result = new WvTemplateTagResultList();
		result.Tags = GetTagsFromTemplate(template);
		//if there are no tags - return one with the template
		if (result.Tags.Count == 0)
		{
			result.Values.Add(template ?? String.Empty);
			return result;
		}

		if (result.Tags.Any(x => x.Type == WvTemplateTagType.ConditionStart) &&
		    result.Tags.Any(x => x.Type == WvTemplateTagType.ConditionEnd))
		{
			var firstStartTag = result.Tags.First(x => x.Type == WvTemplateTagType.ConditionStart);
			var lastEndTag = result.Tags.Last(x => x.Type == WvTemplateTagType.ConditionEnd);
			var firstTagIndex = template!.IndexOf(firstStartTag.FullString!, StringComparison.Ordinal);
			var lastTagIndex = template!.LastIndexOf(lastEndTag.FullString!, StringComparison.Ordinal);
			string templateWithWrappers = template!;
			templateWithWrappers = templateWithWrappers.Substring(firstTagIndex);
			templateWithWrappers = templateWithWrappers.Substring(0, (lastTagIndex + lastEndTag.FullString!.Length));

			string templateWithoutWrappers = templateWithWrappers;
			templateWithoutWrappers = templateWithoutWrappers.Substring(firstStartTag.FullString!.Length);
			templateWithoutWrappers =
				templateWithoutWrappers.Substring(0, templateWithoutWrappers.Length - lastEndTag.FullString!.Length);

			string cleanedTemplate = templateWithoutWrappers;
			//clean all other inline tags
			foreach (var tag in result.Tags)
			{
				if (tag.Type != WvTemplateTagType.InlineStart && tag.Type != WvTemplateTagType.InlineEnd)
					continue;
				cleanedTemplate = cleanedTemplate.Replace(tag.FullString!, "");
			}

			if (firstStartTag.ParamGroups.Count > 0)
			{
				template = template.Replace(templateWithWrappers,String.Empty);
			}
			else
			{
				template = template.Replace(templateWithWrappers,cleanedTemplate);
			}
			result.Tags = GetTagsFromTemplate(template);
		}

		//Process inline templates 
		if (result.Tags.Any(x => x.Type == WvTemplateTagType.InlineStart) &&
		    result.Tags.Any(x => x.Type == WvTemplateTagType.InlineEnd))
		{
			var firstStartTag = result.Tags.First(x => x.Type == WvTemplateTagType.InlineStart);
			var lastEndTag = result.Tags.Last(x => x.Type == WvTemplateTagType.InlineEnd);
			var firstTagIndex = template!.IndexOf(firstStartTag.FullString!, StringComparison.Ordinal);
			var lastTagIndex = template!.LastIndexOf(lastEndTag.FullString!, StringComparison.Ordinal);
			
			string templateWithWrappers = template!;
			templateWithWrappers = templateWithWrappers.Substring(firstTagIndex);
			templateWithWrappers = templateWithWrappers.Substring(0, (lastTagIndex + lastEndTag.FullString!.Length));

			string templateWithoutWrappers = templateWithWrappers;
			templateWithoutWrappers = templateWithoutWrappers.Substring(firstStartTag.FullString!.Length);
			templateWithoutWrappers = templateWithoutWrappers.Substring(0,templateWithoutWrappers.Length - lastEndTag.FullString!.Length);
			
			string cleanedTemplate = templateWithoutWrappers;
			//clean all other inline tags
			foreach (var tag in result.Tags)
			{
				if(tag.Type != WvTemplateTagType.InlineStart && tag.Type != WvTemplateTagType.InlineEnd)
					continue;
				cleanedTemplate = cleanedTemplate.Replace(tag.FullString!, "");
			}
			var inlineTemplateResult = ProcessTemplateTag(cleanedTemplate,dataSource,culture);
			if (inlineTemplateResult.Values.Count == 1)
			{
				if(inlineTemplateResult.Values[0] is string)
					template = template.Replace(templateWithWrappers,(string)inlineTemplateResult.Values[0]);
			}
			else if(inlineTemplateResult.Values.All(x=> x is string))
			{
				var separator = firstStartTag.FlowSeparator ?? "";
				template = template.Replace(templateWithWrappers,String.Join(separator,inlineTemplateResult.Values.Select(x=> (string)x)));
			}
		}

		//if all tags are index - return one with processed template
		if (result.ShouldGenerateOneResult(dataSource))
		{
			var resultValue = GenerateTemplateTagResult(template, result.Tags, dataSource, null, culture);
			if (resultValue is not null && resultValue.Value is not null)
			{
				result.Values.Add(resultValue.Value);
			}
			return result;
		}
		var resultValues = new List<object>();
		for (int i = 0; i < dataSource.Rows.Count; i++)
		{
			var resultValue = GenerateTemplateTagResult(template, result.Tags, dataSource, i, culture);
			if (resultValue is not null && resultValue.Value is not null)
			{
				resultValues.Add(resultValue.Value);
			}
			else
			{
				resultValues.Add(String.Empty);
			}

		}
		if (resultValues.Count == 0)
		{
			result.Values = new();
			result.ExpandCount = 0;
			return result;
		}

		//If all the results are the same as the template return only one
		bool allValuesMatchTemplate = true;
		foreach (var rstValue in resultValues)
		{
			if (rstValue is not string || rstValue.ToString() != template)
			{
				allValuesMatchTemplate = false;
				break;
			}
		}
		if (allValuesMatchTemplate) result.Values.Add(resultValues[0]);
		else result.Values.AddRange(resultValues);
		result.ExpandCount = result.Values.Count;
		return result;
	}

	public object? GetTemplateValue(
		string? template,
		int dataRowPosition,
		DataTable dataSource,
		CultureInfo culture)
	{
		if (String.IsNullOrWhiteSpace(template)) return null;
		var tags = GetTagsFromTemplate(template);
		//if there are no tags - return one with the template
		if (tags.Count == 0) return template;

		string? valueString = template;
		object? value = null;
		if (dataRowPosition < 1) dataRowPosition = 1;
		if (tags.Count == 1 && tags[0].FullString == template)
		{
			(valueString, value) = ProcessTagInTemplate(valueString, value, tags[0], dataSource, dataRowPosition - 1, culture);
			return value;
		}

		foreach (var tag in tags)
		{
			(valueString, value) = ProcessTagInTemplate(valueString, value, tag, dataSource, dataRowPosition - 1, culture);
		}
		if (value is string) return valueString;

		return value;
	}

}

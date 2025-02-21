using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Engines.ExcelFile.Models;

namespace WebVella.DocumentTemplates.Engines.ExcelFile.Utility;
public static partial class WvTemplateTagExtensions
{
	public static List<WvExcelRange> GetApplicableRangeForFunctionTag(this WvTemplateTag tag)
	{
		var result = new List<WvExcelRange>();
		if (tag.Type == WvTemplateTagType.Data) return result;
		if (String.IsNullOrWhiteSpace(tag.FunctionName)) return result;
		if (tag.ParamGroups.Count == 0 || tag.ParamGroups[0].Parameters.Count == 0)
			return result;

		if (tag.Type == WvTemplateTagType.Function)
		{
			if (tag.FunctionName == "sum")
			{
				foreach (var parameter in tag.ParamGroups[0].Parameters)
				{
					if(String.IsNullOrWhiteSpace(parameter.ValueString)) continue;
					var range = new WvExcelRangeHelpers().GetRangeFromString(parameter.ValueString);
					if (range is null) return result;
					result.Add(range);
				}
			}
		}
		else if (tag.Type == WvTemplateTagType.ExcelFunction)
		{
			if (tag.FunctionName == "sum")
			{
				foreach (var parameter in tag.ParamGroups[0].Parameters)
				{
					if(String.IsNullOrWhiteSpace(parameter.ValueString)) continue;
					var range = new WvExcelRangeHelpers().GetRangeFromString(parameter.ValueString);
					if (range is null) return result;
					result.Add(range);
				}
			}
		}
		return result;

	}
}

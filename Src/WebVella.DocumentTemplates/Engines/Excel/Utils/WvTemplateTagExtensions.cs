using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Engines.Excel.Models;

namespace WebVella.DocumentTemplates.Engines.Excel.Utility;
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
				if (String.IsNullOrWhiteSpace(tag.ParamGroups[0].FullString)) return result;

				var address = tag.ParamGroups[0].FullString!.Replace("(", "").Replace(")", "");
				var range = WvExcelRangeHelpers.GetRangeFromString(address);
				if (range is null) return result;
				result.Add(range);
			}
		}
		else if (tag.Type == WvTemplateTagType.ExcelFunction)
		{
			if (tag.FunctionName == "sum")
			{
				if (String.IsNullOrWhiteSpace(tag.ParamGroups[0].FullString)) return result;

				var address = tag.ParamGroups[0].FullString!.Replace("(", "").Replace(")", "");
				var range = WvExcelRangeHelpers.GetRangeFromString(address);
				if (range is null) return result;
				result.Add(range);
			}
		}
		return result;

	}
}

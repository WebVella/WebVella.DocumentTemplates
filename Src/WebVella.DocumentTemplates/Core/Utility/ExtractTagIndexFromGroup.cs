using System.Text.RegularExpressions;

namespace WebVella.DocumentTemplates.Core.Utility;
public partial class WvTemplateUtility
{
	public static int? ExtractTagIndexFromGroup(string indexGroup)
	{
		if (String.IsNullOrWhiteSpace(indexGroup)
		|| !indexGroup.StartsWith("[")
		|| !indexGroup.EndsWith("]")) return null;

		//Remove leading and end []
		var processedIndexGroup = indexGroup.Remove(0, 1);
		processedIndexGroup = processedIndexGroup.Substring(0, processedIndexGroup.Length - 1);
		processedIndexGroup = processedIndexGroup?.Trim();
		if (String.IsNullOrWhiteSpace(processedIndexGroup)) return null;

		if (int.TryParse(processedIndexGroup, out int outInt))
			return outInt;

		return null;
	}
}

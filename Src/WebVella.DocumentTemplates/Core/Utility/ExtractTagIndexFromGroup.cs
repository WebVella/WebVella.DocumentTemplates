using System.Text.RegularExpressions;

namespace WebVella.DocumentTemplates.Core.Utility;
public partial class WvTemplateUtility
{
	private List<int> ExtractTagIndexFromGroup(string indexGroup)
	{
		if (String.IsNullOrWhiteSpace(indexGroup)
		|| !indexGroup.StartsWith("[")
		|| !indexGroup.EndsWith("]")) return new List<int>();

		//Remove leading and end []
		var processedIndexGroup = indexGroup.Remove(0, 1);
		processedIndexGroup = processedIndexGroup.Substring(0, processedIndexGroup.Length - 1);
		processedIndexGroup = processedIndexGroup?.Trim();
		if (String.IsNullOrWhiteSpace(processedIndexGroup)) return new List<int>();

		var result = new List<int>();
		var indexArray = processedIndexGroup.Split(',');
		foreach (var index in indexArray)
		{
			if(String.IsNullOrWhiteSpace(index)) continue;
			
			if (int.TryParse(index.Trim(), out int outInt))
				result.Add(outInt);		
		}

		return result;
	}
}

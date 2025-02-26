namespace WebVella.DocumentTemplates.Core;
public class WvTemplateTagResultList
{
	//What string needs to be replaced in the place of the tag data
	public List<WvTemplateTag> Tags { get; set; } = new();
	public bool AllDataTags
	{
		get
		{
			return Tags.All(x => x.Type == WvTemplateTagType.Data);
		}
	}
	public bool AllTagsAreIndexedFunctionOrWithSeparator
	{
		get
		{
			//All tags should have indexes or have no index by separator defined
			foreach (WvTemplateTag tag in Tags) {
				if(tag.IndexList.Count > 0) continue;
				if(tag.Type == WvTemplateTagType.Function) continue;
				if(tag.Type == WvTemplateTagType.SpreadsheetFunction) continue;
				if(tag.ParamGroups.Any(g => g.Parameters.Any(x=> x.Type.FullName == typeof(WvTemplateTagSeparatorParameterProcessor).FullName)))
					continue;

				return false;
			}
			return true;
			
		}
	}
	public List<object> Values { get; set; } = new();
	public int ExpandCount { get; set; } = 1;
}

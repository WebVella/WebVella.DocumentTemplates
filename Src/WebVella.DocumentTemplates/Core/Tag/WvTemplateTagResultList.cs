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
	public bool IsMultiValueTemplate
	{
		get
		{
			return Tags.Any(x => x.IndexList.Count() == 0 && x.Flow != WvTemplateTagDataFlow.Horizontal);
		}
	}
	public List<object> Values { get; set; } = new();
}

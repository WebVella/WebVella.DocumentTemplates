namespace WebVella.DocumentTemplates.Core;
public class WvTemplateTagResult
{
	//What string needs to be replaced in the place of the tag data
	public List<WvTemplateTag> Tags { get; set; } = new();
	public string? ValueString { get; set; }
	public object? Value { get; set; }
}

namespace WebVella.DocumentTemplates.Core;
public class WvTemplateTagParamGroup
{
	//What string needs to be replaced in the place of the tag data
	public string? FullString { get; set; }
	public List<IWvTemplateTagParameterProcessorBase> Parameters { get; set; } = new List<IWvTemplateTagParameterProcessorBase>();
}

namespace WebVella.DocumentTemplates.Core;
public abstract class WvTemplateProcessContextBase
{
	public Guid Id { get; set; }
	public WvTemplateTagResultList TagProcessResult { get; set; } = new();
	public List<WvTemplateTag> Tags { get; set; } = new List<WvTemplateTag>();
	public HashSet<Guid> Dependencies { get; set; } = new();
	public HashSet<Guid> Dependants { get; set; } = new();
	//for optimization purpose - when all tags are a data type their values are set during placement
	public bool IsDataSet { get; set; } = false;
}

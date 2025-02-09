namespace WebVella.DocumentTemplates.Core;
public abstract class WvTemplateProcessResultItemBase
{
	public virtual object? Result { get; set; }
	public List<WvTemplateProcessContextBase> Contexts { get; set; } = new();
	//To find how many attempts were made for a context to be processed
	public HashSet<Guid> ProcessedContexts { get; set; } = new();
	public Dictionary<Guid,int> ContextProcessLog { get; set; } = new();	
}

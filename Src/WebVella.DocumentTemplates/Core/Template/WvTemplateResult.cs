namespace WebVella.DocumentTemplates.Core;
public class WvTemplateResult
{
	public virtual object? Template { get; set; }
	public virtual object? Result { get; set; }
	public virtual IEnumerable<object>? Contexts { get; set; }
	public HashSet<Guid> ProcessedContexts { get; set; } = new();
	//To find how many attempts were made for a context to be processed
	public Dictionary<Guid,int> ContextProcessLog { get; set; } = new();	
}

using System.Data;

namespace WebVella.DocumentTemplates.Core;
public abstract class WvTemplateProcessResultItemBase
{
	public virtual object? Result { get; set; }
	public virtual List<WvTemplateProcessContextBase> Contexts { get; set; } = new();
	public bool HasError => Contexts.Any(x => x.HasError);
	public DataTable? DataTable { get; set; } = null;
}

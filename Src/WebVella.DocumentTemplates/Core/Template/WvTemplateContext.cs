using System.Data;
using System.Globalization;

namespace WebVella.DocumentTemplates.Core;
public class WvTemplateContext
{
	public Guid Id { get; set; }
	public WvTemplateTagResultList TagProcessResult { get; set; } = new();
	public IEnumerable<WvTemplateTag> Tags { get; set; } = Enumerable.Empty<WvTemplateTag>();
	public HashSet<Guid> Dependencies { get; set; } = new();
	public HashSet<Guid> Dependants { get; set; } = new();
	//for optimization purpose - when all tags are a data type their values are set during placement
	public bool IsDataSet { get; set; } = false;
}

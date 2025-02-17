namespace WebVella.DocumentTemplates.Core;
public class WvTemplateTagUnknownParameterProcessor : IWvTemplateTagParameterProcessorBase
{
	public Type Type { get => this.GetType(); }
	public string Name { get; set; } = "uknown";
	public string? ValueString { get; set; }
	public string? Value { get; set; }
	public int Priority { get; set; } = 1000;
	public WvTemplateTagUnknownParameterProcessor() { }
	public WvTemplateTagUnknownParameterProcessor(string? name, string? valueString)
	{
		Name = String.IsNullOrWhiteSpace(name) ? "uknown" : name;
		ValueString = valueString;
		Value = valueString;
	}
}
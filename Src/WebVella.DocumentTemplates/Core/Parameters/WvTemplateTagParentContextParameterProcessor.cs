namespace WebVella.DocumentTemplates.Core;
public class WvTemplateTagParentContextParameterProcessor : IWvTemplateTagParameterProcessorBase
{
	public Type Type { get => this.GetType(); }
	public string Name { get; set; } = "pc";
	public WvTemplateTagParamOperatorType OperatorType { get; set; } = WvTemplateTagParamOperatorType.Unknown;
	public string? ValueString { get; set; }
	public string? Value { get; set; }
	public int Priority { get; set; } = 1000;
	public WvTemplateTagParentContextParameterProcessor() { }
	public WvTemplateTagParentContextParameterProcessor(string? valueString)
	{
		ValueString = valueString;
		Value = valueString;
	}
}
namespace WebVella.DocumentTemplates.Core;
public class WvTemplateTagSeparatorParameterProcessor : IWvTemplateTagParameterProcessorBase
{
	public Type Type { get => this.GetType(); }
	public string Name { get; set; } = "s";
	public WvTemplateTagParamOperatorType OperatorType { get; set; } = WvTemplateTagParamOperatorType.Unknown;
	public string? ValueString { get; set; }
	public string? Value { get; set; }
	public int Priority { get; set; } = 1000;
	public WvTemplateTagSeparatorParameterProcessor() { }
	public WvTemplateTagSeparatorParameterProcessor(string? valueString)
	{
		ValueString = valueString;
		Value = valueString;
	}
}
namespace WebVella.DocumentTemplates.Core;
public class WvTemplateTagUnknownParameterProcessor : IWvTemplateTagParameterProcessorBase
{
	public Type Type { get => this.GetType(); }
	public string Name { get; set; } = "unknown";
	public WvTemplateTagParamOperatorType OperatorType { get; set; } = WvTemplateTagParamOperatorType.Unknown;
	public string? ValueString { get; set; }
	public string? Value { get; set; }
	public int Priority { get; set; } = 1000;
	public WvTemplateTagUnknownParameterProcessor() { }
	public WvTemplateTagUnknownParameterProcessor(string? name, string? valueString, WvTemplateTagParamOperatorType operatorTypeEnum)
	{
		Name = String.IsNullOrWhiteSpace(name) ? "unknown" : name;
		ValueString = valueString;
		Value = valueString;
		OperatorType = operatorTypeEnum;
	}
}
using WebVella.DocumentTemplates.Extensions;

namespace WebVella.DocumentTemplates.Core;
public class WvTemplateTagDataFlowParameterProcessor : IWvTemplateTagParameterProcessorBase
{
	public Type Type { get => this.GetType(); }
	public string Name { get; set; } = "f";
	public WvTemplateTagParamOperatorType OperatorType { get; set; } = WvTemplateTagParamOperatorType.Unknown;
	public string? ValueString { get; set; } = null;
	public WvTemplateTagDataFlow Value { get; set; }
	public int Priority { get; set; } = 1000;
	public WvTemplateTagDataFlowParameterProcessor() { }
	public WvTemplateTagDataFlowParameterProcessor(string? valueString)
	{
		ValueString = valueString;
		Value = StringToValue(valueString);
	}
	public string ValueToString(WvTemplateTagDataFlow? value)
		=> value is null ? WvTemplateTagDataFlow.Vertical.ToDescriptionString() : value.Value.ToDescriptionString();
	public WvTemplateTagDataFlow StringToValue(string? valueString)
	{
		if(String.IsNullOrWhiteSpace(valueString)) return WvTemplateTagDataFlow.Vertical;

		foreach (var item in Enum.GetValues<WvTemplateTagDataFlow>())
		{
			if (item.ToDescriptionString().ToLowerInvariant() == valueString.ToLowerInvariant())
				return item;
		}
		return WvTemplateTagDataFlow.Vertical;
	}
}
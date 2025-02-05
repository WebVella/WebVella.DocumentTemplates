using WebVella.DocumentTemplates.Extensions;

namespace WebVella.DocumentTemplates.Core;
public class WvTemplateTagDataFlowParameter : IWvTemplateTagParameterBase
{
	public Type Type { get => this.GetType(); }
	public string Name { get; set; } = "f";
	public string? ValueString { get; set; } = null;
	public WvTemplateTagType TagType { get => WvTemplateTagType.Data; }
	public WvTemplateTagDataFlow Value { get; set; }
	public WvTemplateTagDataFlowParameter() { }
	public WvTemplateTagDataFlowParameter(string? valueString)
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
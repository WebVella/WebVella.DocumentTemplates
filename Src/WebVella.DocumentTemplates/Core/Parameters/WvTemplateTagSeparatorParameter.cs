namespace WebVella.DocumentTemplates.Core;
public class WvTemplateTagSeparatorParameter : IWvTemplateTagParameterBase
{
	public Type Type { get => this.GetType(); }
	public string Name { get; set; } = "s";
	public string? ValueString { get; set; }
	public WvTemplateTagType TagType { get => WvTemplateTagType.Data; }
	public string? Value { get; set; }

	public WvTemplateTagSeparatorParameter() { }
	public WvTemplateTagSeparatorParameter(string? valueString)
	{
		ValueString = valueString;
		Value = valueString;
	}
}
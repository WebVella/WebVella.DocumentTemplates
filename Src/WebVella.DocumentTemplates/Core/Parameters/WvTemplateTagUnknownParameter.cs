namespace WebVella.DocumentTemplates.Core;
public class WvTemplateTagUnknownParameter : IWvTemplateTagParameterBase
{
	public Type Type { get => this.GetType(); }
	public string Name { get; set; } = "uknown";
	public string? ValueString { get; set; }
	public WvTemplateTagType TagType { get => WvTemplateTagType.Data; }
	public string? Value { get; set; }
	public WvTemplateTagUnknownParameter() { }
	public WvTemplateTagUnknownParameter(string? name, string? valueString)
	{
		Name = String.IsNullOrWhiteSpace(name) ? "uknown" : name;
		ValueString = valueString;
		Value = valueString;
	}
}
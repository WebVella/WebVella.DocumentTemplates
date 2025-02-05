namespace WebVella.DocumentTemplates.Core;
public interface IWvTemplateTagParameterBase
{
	Type Type { get; }
	string Name { get; set; }
	string? ValueString { get; set; }
	WvTemplateTagType TagType { get; }
}

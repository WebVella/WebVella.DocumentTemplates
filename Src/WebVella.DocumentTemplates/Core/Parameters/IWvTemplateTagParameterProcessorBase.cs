﻿namespace WebVella.DocumentTemplates.Core;
public interface IWvTemplateTagParameterProcessorBase
{
	Type Type { get; }
	string Name { get; set; }
	string? ValueString { get; set; }
	int Priority { get; set; }
}

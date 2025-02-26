using WebVella.DocumentTemplates.Extensions;

namespace WebVella.DocumentTemplates.Core;
public class WvTemplateTag
{
	//What string needs to be replaced in the place of the tag data
	public string? FullString { get; set; }
	public string? Name { get; set; }
	public WvTemplateTagType Type { get; set; } = WvTemplateTagType.Data;
	//as in the sheet name there cannot be used [] and in all cases that there is a list of data matched
	//in the methods the first one is always 0 by default
	public List<int> IndexList { get; set; } = new List<int>();
	public WvTemplateTagDataFlow? Flow
	{
		get
		{
			WvTemplateTagDataFlow? flow = null;
			foreach (var paramGroup in ParamGroups)
			{
				if (flow is not null) break;
				foreach (var param in paramGroup.Parameters)
				{
					if (param.Type.InheritsClass(typeof(WvTemplateTagDataFlowParameterProcessor)))
					{
						if (((WvTemplateTagDataFlowParameterProcessor)param).Value == WvTemplateTagDataFlow.Horizontal)
						{
							flow = WvTemplateTagDataFlow.Horizontal;
							break;
						}
						else if (((WvTemplateTagDataFlowParameterProcessor)param).Value == WvTemplateTagDataFlow.Vertical)
						{
							flow = WvTemplateTagDataFlow.Vertical;
							break;
						}
					}
				}
			}

			if (!String.IsNullOrEmpty(FlowSeparator))
			{
				flow = WvTemplateTagDataFlow.Horizontal;
			}
			return flow;
		}
	}
	public string? FlowSeparator
	{
		get
		{
			foreach (var paramGroup in ParamGroups)
			{
				foreach (var param in paramGroup.Parameters)
				{
					if (param.Type.InheritsClass(typeof(WvTemplateTagSeparatorParameterProcessor)))
					{
						if(((WvTemplateTagSeparatorParameterProcessor)param).ValueString == "$rn"){ 
							return Environment.NewLine;
						}
						return ((WvTemplateTagSeparatorParameterProcessor)param).ValueString;
					}
				}
			}
			return null;
		}
	}
	public virtual List<WvTemplateTagParamGroup> ParamGroups { get; set; } = new List<WvTemplateTagParamGroup>();
	//Will be filled in if the function is supported by the system
	public string? FunctionName { get; set; } = null;
}

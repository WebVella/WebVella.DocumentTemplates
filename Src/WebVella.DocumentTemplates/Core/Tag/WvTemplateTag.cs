using WebVella.DocumentTemplates.Extensions;

namespace WebVella.DocumentTemplates.Core;
public class WvTemplateTag
{
	//What string needs to be replaced in the place of the tag data
	public string? FullString { get; set; }
	public string? Operator { get; set; }
	public WvTemplateTagType Type { get; set; } = WvTemplateTagType.Data;
	//as in the sheet name there cannot be used [] and in all cases that there is a list of data matched
	//in the methods the first one is always 0 by default
	public List<WvTemplateTagIndexGroup> IndexGroups { get; set; } = new();
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

			if (flow is null && FlowSeparatorList.Count > 0)
			{
				flow = WvTemplateTagDataFlow.Horizontal;
			}
			return flow;
		}
	}
	
	public string FlowSeparator {
		get
		{
			if (FlowSeparatorList.Count == 0)
				return String.Empty;

			return FlowSeparatorList[0];
		}
	}
	public List<string> FlowSeparatorList
	{
		get
		{
			var separators = new List<string>();
			foreach (var paramGroup in ParamGroups)
			{
				foreach (var param in paramGroup.Parameters)
				{
					if (param.Type.InheritsClass(typeof(WvTemplateTagSeparatorParameterProcessor)))
					{
						var parsedParam = (WvTemplateTagSeparatorParameterProcessor)param;
						if (parsedParam.ValueString == "$rn")
						{
							separators.Add(Environment.NewLine);
							continue;
						}
						separators.Add(parsedParam.ValueString ?? String.Empty);
					}

					if (param.Type.InheritsClass(typeof(WvTemplateTagDataFlowParameterProcessor)))
					{
						var parsedParam = (WvTemplateTagDataFlowParameterProcessor)param;
						if (parsedParam.Value == WvTemplateTagDataFlow.Vertical)
						{
							separators.Add(Environment.NewLine);
							continue;
						}
					}
				}
			}
			return separators;
		}
	}
	public virtual List<WvTemplateTagParamGroup> ParamGroups { get; set; } = new List<WvTemplateTagParamGroup>();
	//Will be filled in if the function is supported by the system
	public string? ItemName { get; set; } = null;
}

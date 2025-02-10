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
	public WvTemplateTagDataFlow Flow
	{
		get
		{
			if(!String.IsNullOrEmpty(FlowSeparator)){ 
				return WvTemplateTagDataFlow.Horizontal;
			}

			foreach (var paramGroup in ParamGroups)
			{
				foreach (var param in paramGroup.Parameters)
				{
					if (param.Type.InheritsClass(typeof(WvTemplateTagDataFlowParameter))
						&& ((WvTemplateTagDataFlowParameter)param).Value == WvTemplateTagDataFlow.Horizontal)
					{
						return WvTemplateTagDataFlow.Horizontal;
					}
				}
			}

			return WvTemplateTagDataFlow.Vertical;
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
					if (param.Type.InheritsClass(typeof(WvTemplateTagSeparatorParameter)))
					{
						return ((WvTemplateTagSeparatorParameter)param).ValueString;
					}
				}
			}
			return null;
		}
	}
	public virtual List<WvTemplateTagParamGroup> ParamGroups { get; set; } = new List<WvTemplateTagParamGroup>();

}

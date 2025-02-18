using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using System.Drawing;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Engines.Excel.Utility;
namespace WebVella.DocumentTemplates.Engines.Excel;
public class WvExcelFileTemplateContext
{
	public Guid Id { get; set; }
	//Row is null for embeds like Pictures as they are not part of the process flow
	public WvExcelFileTemplateRow? Row { get; set; } = null;
	public IXLWorksheet? Worksheet { get; set; }
	public WvTemplateTagDataFlow? ForcedFlow { get; set; } = null;
	public WvTemplateTagDataFlow Flow
	{
		get
		{
			if (ForcedFlow.HasValue) return ForcedFlow.Value;
			if (ParentContext is not null)
			{
				WvTemplateTagDataFlow? parentFlow = ParentContext.Flow;
				if (parentFlow.HasValue) return parentFlow.Value;
			}
			return WvTemplateTagDataFlow.Vertical;
		}
	}
	public WvExcelFileTemplateContextType Type { get; set; } = WvExcelFileTemplateContextType.CellRange;
	public IXLRange? Range { get; set; }
	public IXLPicture? Picture { get; set; }
	//The ItemRanges are index in the order of the datasource rows
	public List<IXLRange> ItemRanges { get; set; } = new();
	public bool IsNullContextForced { get; set; } = false;
	public WvExcelFileTemplateContext? ForcedContext { get; set; }
	public WvExcelFileTemplateContext? TopContext { get; set; }
	public WvExcelFileTemplateContext? LeftContext { get; set; }
	public WvExcelFileTemplateContext? ParentContext
	{
		get
		{
			if (IsNullContextForced) return null;
			if (ForcedContext is not null) return ForcedContext;
			WvTemplateTagDataFlow flow = ForcedFlow ?? WvTemplateTagDataFlow.Vertical;

			if (flow == WvTemplateTagDataFlow.Vertical && LeftContext is not null) return LeftContext;
			if (flow == WvTemplateTagDataFlow.Horizontal && TopContext is not null) return TopContext;
			//first column cells even verticle does not have left context - attache to the top
			if (flow == WvTemplateTagDataFlow.Vertical && TopContext is not null) return TopContext;
			//first row cells even verticle does not have top context - attach to the left
			if (flow == WvTemplateTagDataFlow.Horizontal && LeftContext is not null) return LeftContext;

			return null;
		}
	}
	public HashSet<Guid> ContextDependencies { get; set; } = new();
}

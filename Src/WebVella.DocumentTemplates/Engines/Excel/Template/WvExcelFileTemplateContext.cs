using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Excel;
public class WvExcelFileTemplateContext
{
	public Guid Id { get; set; }
	public IXLWorksheet? Worksheet { get; set; }
	public WvTemplateTagDataFlow Flow { get; set; } = WvTemplateTagDataFlow.Vertical;
	public WvExcelFileTemplateContextType Type { get; set; } = WvExcelFileTemplateContextType.CellRange;
	public IXLRange? Range { get; set; }
	public IXLPicture? Picture { get; set; }
	public WvExcelFileTemplateContext? TopContext { get; set; }
	public WvExcelFileTemplateContext? LeftContext { get; set; }
	public WvExcelFileTemplateContext? ParentContext
	{
		get
		{
			if (Flow == WvTemplateTagDataFlow.Vertical && LeftContext is not null) return LeftContext;
			if (Flow == WvTemplateTagDataFlow.Horizontal && TopContext is not null) return TopContext;

			//first column cells even verticle does not have left context - attache to the top
			if (Flow == WvTemplateTagDataFlow.Vertical && TopContext is not null) return TopContext;

			//first row cells even verticle does not have top context - attach to the left
			if (Flow == WvTemplateTagDataFlow.Horizontal && LeftContext is not null) return LeftContext;
			return null;
		}
	}
	public HashSet<Guid> ContextDependencies { get; set; } = new();
}

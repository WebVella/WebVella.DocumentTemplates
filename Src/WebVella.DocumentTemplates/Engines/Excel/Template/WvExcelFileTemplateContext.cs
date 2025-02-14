using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Excel;
public class WvExcelFileTemplateContext
{
	public Guid Id { get; set; }
	//Row is null for embeds like Pictures as they are not part of the process flow
	public WvExcelFileTemplateRow? Row { get; set; } = null; 
	public IXLWorksheet? Worksheet { get; set; }
	public WvTemplateTagDataFlow Flow { get; set; } = WvTemplateTagDataFlow.Vertical;
	public WvExcelFileTemplateContextType Type { get; set; } = WvExcelFileTemplateContextType.CellRange;
	public IXLRange? Range { get; set; }
	public IXLPicture? Picture { get; set; }
	public HashSet<Guid> ContextDependencies { get; set; } = new();
}

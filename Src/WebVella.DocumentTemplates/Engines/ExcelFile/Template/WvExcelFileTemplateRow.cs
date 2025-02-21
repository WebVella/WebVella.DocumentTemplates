using ClosedXML.Excel;
namespace WebVella.DocumentTemplates.Engines.ExcelFile;
public class WvExcelFileTemplateRow
{
	public int Position { get; set; }
	public IXLWorksheet? Worksheet { get; set; }
	public List<Guid> Contexts { get; set; } = new();
}

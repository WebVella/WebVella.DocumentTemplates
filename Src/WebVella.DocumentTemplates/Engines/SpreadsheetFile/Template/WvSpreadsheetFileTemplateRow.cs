using ClosedXML.Excel;
namespace WebVella.DocumentTemplates.Engines.SpreadsheetFile;
public class WvSpreadsheetFileTemplateRow
{
	public int Position { get; set; }
	public IXLWorksheet? Worksheet { get; set; }
	public List<Guid> Contexts { get; set; } = new();
}

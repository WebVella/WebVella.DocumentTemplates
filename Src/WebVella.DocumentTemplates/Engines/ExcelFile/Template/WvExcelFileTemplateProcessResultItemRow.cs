using ClosedXML.Excel;
namespace WebVella.DocumentTemplates.Engines.ExcelFile;
public class WvExcelFileTemplateProcessResultItemRow
{
	public int TemplateFirstRow { get; set; }
	public int ResultFirstRow { get; set; } = 1;
	public IXLWorksheet? Worksheet { get; set; }
	public List<Guid> Contexts { get; set; } = new();
	//Coordinates are relative to the template Row not the Result worksheet
	//So the grid always starts with row and column 1
	public List<WvExcelFileTemplateProcessResultItemRowGridItem> RelativeGrid { get; set; } = new();
	//optimization
	public int RelativeGridMaxRow { get; set; } = 0;
	public int RelativeGridMaxColumn { get; set; } = 0;
	public HashSet<string> RelativeGridAddressHs { get; set; } = new();

}

public class WvExcelFileTemplateProcessResultItemRowGridItem
{
	public int Row { get; set; }
	public int Column { get; set; }
	public Guid ContextId { get; set; }
}

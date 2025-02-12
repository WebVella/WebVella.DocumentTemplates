namespace WebVella.DocumentTemplates.Engines.Excel.Models;
public class WvExcelRange
{
	public string? Worksheet { get; set; }
	public int FirstRow { get; set; }
	public int FirstColumn { get; set; }
	public int LastRow { get; set; }
	public int LastColumn { get; set; }
}

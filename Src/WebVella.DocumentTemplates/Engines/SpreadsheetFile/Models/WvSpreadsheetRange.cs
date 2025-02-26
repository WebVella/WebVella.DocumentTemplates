namespace WebVella.DocumentTemplates.Engines.SpreadsheetFile.Models;
public record WvSpreadsheetRange
{
	public string? Worksheet { get; set; }
	public int FirstRow { get; set; }
	public int FirstColumn { get; set; }
	public int LastRow { get; set; }
	public int LastColumn { get; set; }
	public bool FirstRowLocked { get; set; } = false;
	public bool FirstColumnLocked { get; set; } = false;
	public bool LastRowLocked { get; set; } = false;
	public bool LastColumnLocked { get; set; } = false;
}

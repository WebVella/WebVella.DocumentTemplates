namespace WebVella.DocumentTemplates.Engines.SpreadsheetFile;
public class WvSpreadsheetFileTemplateContextRangeAddress
{
	public int FirstRow { get; set; }
	public int FirstColumn { get; set; }
	public int LastRow { get; set; }
	public int LastColumn { get; set; }

	public WvSpreadsheetFileTemplateContextRangeAddress() { }
	public WvSpreadsheetFileTemplateContextRangeAddress(int firstRow, int firstColumn, int lastRow, int lastColumn)
	{
		FirstRow = firstRow; FirstColumn = firstColumn;
		LastRow = lastRow; LastColumn = lastColumn;
	}
}
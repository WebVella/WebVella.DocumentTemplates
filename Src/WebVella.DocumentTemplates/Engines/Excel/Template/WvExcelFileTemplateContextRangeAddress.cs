namespace WebVella.DocumentTemplates.Engines.Excel;
public class WvExcelFileTemplateContextRangeAddress
{
	public int FirstRow { get; set; }
	public int FirstColumn { get; set; }
	public int LastRow { get; set; }
	public int LastColumn { get; set; }

	public WvExcelFileTemplateContextRangeAddress() { }
	public WvExcelFileTemplateContextRangeAddress(int firstRow, int firstColumn, int lastRow, int lastColumn)
	{
		FirstRow = firstRow; FirstColumn = firstColumn;
		LastRow = lastRow; LastColumn = lastColumn;
	}
}
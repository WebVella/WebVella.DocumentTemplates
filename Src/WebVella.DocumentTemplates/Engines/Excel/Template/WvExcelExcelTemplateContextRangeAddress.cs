namespace WebVella.DocumentTemplates.Engines.Excel;
public class WvExcelExcelTemplateContextRangeAddress
{
	public int FirstRow { get; set; }
	public int FirstColumn { get; set; }
	public int LastRow { get; set; }
	public int LastColumn { get; set; }

	public WvExcelExcelTemplateContextRangeAddress() { }
	public WvExcelExcelTemplateContextRangeAddress(int firstRow, int firstColumn, int lastRow, int lastColumn)
	{
		FirstRow = firstRow; FirstColumn = firstColumn;
		LastRow = lastRow; LastColumn = lastColumn;
	}
}
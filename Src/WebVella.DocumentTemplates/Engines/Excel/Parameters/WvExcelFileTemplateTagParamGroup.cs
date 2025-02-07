namespace WebVella.DocumentTemplates.Engines.Excel;
public class WvExcelFileTemplateTagParamGroup
{
	//What string needs to be replaced in the place of the tag data
	public string? FullString { get; set; }
	public IEnumerable<IWvExcelFileTemplateTagParameter> Parameters { get; set; } = Enumerable.Empty<IWvExcelFileTemplateTagParameter>();
}

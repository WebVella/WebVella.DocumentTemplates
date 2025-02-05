namespace WebVella.DocumentTemplates.Engines.Excel;
public class WvExcelTemplateTagParamGroup
{
	//What string needs to be replaced in the place of the tag data
	public string? FullString { get; set; }
	public IEnumerable<IWvExcelTemplateTagParameter> Parameters { get; set; } = Enumerable.Empty<IWvExcelTemplateTagParameter>();
}

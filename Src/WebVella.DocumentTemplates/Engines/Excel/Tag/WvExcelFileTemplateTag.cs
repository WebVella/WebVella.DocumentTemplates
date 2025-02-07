using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Excel;
public class WvExcelFileTemplateTag : WvTemplateTag
{
	public new IEnumerable<WvExcelFileTemplateTagParamGroup> ParamGroups { get; set; } = Enumerable.Empty<WvExcelFileTemplateTagParamGroup>();
}
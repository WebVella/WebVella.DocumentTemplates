using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Excel;
public class WvExcelTemplateTag : WvTemplateTag
{
	public new IEnumerable<WvExcelTemplateTagParamGroup> ParamGroups { get; set; } = Enumerable.Empty<WvExcelTemplateTagParamGroup>();
}
using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Excel;
public class WvExcelTemplateTagSeparatorParameter : WvTemplateTagSeparatorParameter, IWvExcelTemplateTagParameter
{
	public HashSet<Guid> GetDependencies(WvExcelFileTemplateResult result, WvExcelExcelTemplateContext context,
		WvTemplateTag tag, WvTemplateTagParamGroup parameterGroup, IWvExcelTemplateTagParameter paramter)
	 => new HashSet<Guid>();
}
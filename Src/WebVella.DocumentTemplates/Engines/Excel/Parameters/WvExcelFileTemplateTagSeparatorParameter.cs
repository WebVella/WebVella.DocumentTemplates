using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Excel;
public class WvExcelFileTemplateTagSeparatorParameter : WvTemplateTagSeparatorParameter, IWvExcelFileTemplateTagParameter
{
	public HashSet<Guid> GetDependencies(WvExcelFileTemplateResult result, WvExcelFileTemplateContext context,
		WvTemplateTag tag, WvTemplateTagParamGroup parameterGroup, IWvExcelFileTemplateTagParameter paramter)
	 => new HashSet<Guid>();
}
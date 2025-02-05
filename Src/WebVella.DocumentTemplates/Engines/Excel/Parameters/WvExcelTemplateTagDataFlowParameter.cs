using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Excel;
public class WvExcelTemplateTagDataFlowParameter : WvTemplateTagDataFlowParameter, IWvExcelTemplateTagParameter
{
	public HashSet<Guid> GetDependencies(WvExcelFileTemplateResult result, WvExcelExcelTemplateContext context,
		WvTemplateTag tag, WvTemplateTagParamGroup parameterGroup, IWvExcelTemplateTagParameter paramter)
	 => new HashSet<Guid>();
}
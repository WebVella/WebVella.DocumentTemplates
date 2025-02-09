using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Excel;
public class WvExcelFileTemplateTagDataFlowParameter : WvTemplateTagDataFlowParameter, IWvExcelFileTemplateTagParameter
{
	public HashSet<Guid> GetDependencies(WvExcelFileTemplateProcessResultItem resultItem, WvExcelFileTemplateProcessContext context,
		WvTemplateTag tag, WvTemplateTagParamGroup parameterGroup, IWvExcelFileTemplateTagParameter paramter)
	 => new HashSet<Guid>();
}
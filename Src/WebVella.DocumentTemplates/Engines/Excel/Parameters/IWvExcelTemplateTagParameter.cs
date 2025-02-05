using WebVella.DocumentTemplates.Core;

namespace WebVella.DocumentTemplates.Engines.Excel;
public interface IWvExcelTemplateTagParameter
{
	HashSet<Guid> GetDependencies(WvExcelFileTemplateResult result,
		WvExcelExcelTemplateContext context,
		WvTemplateTag tag,
		WvTemplateTagParamGroup parameterGroup,
		IWvExcelTemplateTagParameter parameter);
}

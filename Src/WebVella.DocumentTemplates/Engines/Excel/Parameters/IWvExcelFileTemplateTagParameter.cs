using WebVella.DocumentTemplates.Core;

namespace WebVella.DocumentTemplates.Engines.Excel;
public interface IWvExcelFileTemplateTagParameter
{
	HashSet<Guid> GetDependencies(WvExcelFileTemplateResult result,
		WvExcelFileTemplateContext context,
		WvTemplateTag tag,
		WvTemplateTagParamGroup parameterGroup,
		IWvExcelFileTemplateTagParameter parameter);
}

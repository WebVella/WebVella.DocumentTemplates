using WebVella.DocumentTemplates.Core;

namespace WebVella.DocumentTemplates.Engines.Excel;
public interface IWvExcelFileTemplateTagParameter
{
	HashSet<Guid> GetDependencies(WvExcelFileTemplateProcessResultItem resultItem,
		WvExcelFileTemplateProcessContext context,
		WvTemplateTag tag,
		WvTemplateTagParamGroup parameterGroup,
		IWvExcelFileTemplateTagParameter parameter);
}

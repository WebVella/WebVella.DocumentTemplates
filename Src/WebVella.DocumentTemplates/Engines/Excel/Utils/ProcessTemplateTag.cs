using ClosedXML.Excel;
using System.Data;

namespace WebVella.DocumentTemplates.Engines.Excel.Utility;
public static partial class WvExcelFileEngineUtility
{
	public static IXLCell? GetResultRangeTopLeftForTemplateRangeTopLeft(
		List<WvExcelFileTemplateProcessContext>? contextList,
		IXLWorksheet worksheet,
		IXLCell templateRangeTopLeft)
	{
		if (contextList == null || contextList.Count == 0) return null;
		var resultContexts = contextList.Where(x => x.TemplateWorksheet is not null && x.TemplateWorksheet.Position == worksheet.Position).ToList();
		return null;
	}


}

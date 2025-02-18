using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using System.Data;
using System.Globalization;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Extensions;

namespace WebVella.DocumentTemplates.Engines.Excel.Utility;
public partial class WvExcelFileEngineUtility
{
	/// <summary>
	/// Generates dependencies between the contexts based on formula fields ranges
	/// </summary>
	/// <param name="result"></param>
	public void ProcessExcelTemplateCalculateDependencies(WvExcelFileTemplateProcessResult result)
	{
		foreach (var context in result.TemplateContexts)
		{
			context.CalculateDependencies(result.TemplateContexts);
		}
	}
}

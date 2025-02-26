using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using System.Data;
using System.Globalization;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Extensions;

namespace WebVella.DocumentTemplates.Engines.SpreadsheetFile.Utility;
public partial class WvSpreadsheetFileEngineUtility
{
	/// <summary>
	/// Generates dependencies between the contexts based on formula fields ranges
	/// </summary>
	/// <param name="result"></param>
	public void ProcessSpreadsheetTemplateCalculateDependencies(WvSpreadsheetFileTemplateProcessResult result)
	{
		foreach (var context in result.TemplateContexts)
		{
			context.CalculateDependencies(result.TemplateContexts);
		}
	}
}

using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using System.Data;
using System.Globalization;
using System.Text;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Engines.ExcelFile.Utility;
using WebVella.DocumentTemplates.Extensions;
namespace WebVella.DocumentTemplates.Engines.ExcelFile;
public class WvExcelFileTemplate : WvTemplateBase
{
	public XLWorkbook? Template { get; set; }
	public WvExcelFileTemplateProcessResult Process(DataTable? dataSource, CultureInfo? culture = null)
	{
		if (culture == null) culture = new CultureInfo("en-US");
		if (dataSource is null) throw new ArgumentException("No datasource provided!", nameof(dataSource));
		var result = new WvExcelFileTemplateProcessResult()
		{
			Template = Template,
			GroupByDataColumns = GroupDataByColumns,
			ResultItems = new()
		};
		if (Template is null)
		{
			result.ResultItems.Add(new WvExcelFileTemplateProcessResultItem
			{
				Result = null
			});
			return result;
		};
		new WvExcelFileEngineUtility().ProcessExcelTemplateInitTemplateContexts(result);
		new WvExcelFileEngineUtility().ProcessExcelTemplateCalculateDependencies(result);
		var templateContextDict = result.TemplateContexts.ToDictionary(x => x.Id);
		var datasourceGroups = dataSource.GroupBy(GroupDataByColumns);
		foreach (var grouptedDs in datasourceGroups)
		{
			var resultItem = new WvExcelFileTemplateProcessResultItem
			{
				Result = new XLWorkbook(),
				NumberOfDataTableRows = grouptedDs.Rows.Count
			};
			var templateContextsDict = result.TemplateContexts.ToDictionary(x => x.Id);
			new WvExcelFileEngineUtility().ProcessExcelTemplateGenerateResultContexts(
				result: result,
				resultItem: resultItem,
				dataSource: grouptedDs,
				culture: culture,
				templateContextDict:templateContextDict);
			result.ResultItems.Add(resultItem);
		}
		return result;
	}

}
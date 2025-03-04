using ClosedXML.Excel;
using System.Data;
using System.Globalization;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Engines.SpreadsheetFile.Utility;
using WebVella.DocumentTemplates.Extensions;
namespace WebVella.DocumentTemplates.Engines.SpreadsheetFile;
public class WvSpreadsheetFileTemplate : WvTemplateBase
{
	public MemoryStream? Template { get; set; }
	public WvSpreadsheetFileTemplateProcessResult Process(DataTable? dataSource, CultureInfo? culture = null)
	{
		if (culture == null) culture = new CultureInfo("en-US");
		if (dataSource is null) throw new ArgumentException("No datasource provided!", nameof(dataSource));
		var result = new WvSpreadsheetFileTemplateProcessResult()
		{
			Template = Template,
			GroupByDataColumns = GroupDataByColumns,
			ResultItems = new()
		};
		if (Template is null)
		{
			result.ResultItems.Add(new WvSpreadsheetFileTemplateProcessResultItem
			{
				Workbook = null
			});
			return result;
		};
		result.Workbook = new XLWorkbook(Template);

		new WvSpreadsheetFileEngineUtility().ProcessSpreadsheetTemplateInitTemplateContexts(result);
		new WvSpreadsheetFileEngineUtility().ProcessSpreadsheetTemplateCalculateDependencies(result);
		var templateContextDict = result.TemplateContexts.ToDictionary(x => x.Id);
		var datasourceGroups = dataSource.GroupBy(GroupDataByColumns);
		foreach (var grouptedDs in datasourceGroups)
		{
			var resultItem = new WvSpreadsheetFileTemplateProcessResultItem
			{
				Workbook = new XLWorkbook(),
				DataTable = grouptedDs
			};
			var templateContextsDict = result.TemplateContexts.ToDictionary(x => x.Id);
			new WvSpreadsheetFileEngineUtility().ProcessSpreadsheetTemplateGenerateResultContexts(
				result: result,
				resultItem: resultItem,
				dataSource: grouptedDs,
				culture: culture,
				templateContextDict: templateContextDict);

			
			resultItem.Result = null;
			if(resultItem.Workbook is not null){ 
			var ms = new MemoryStream();
				resultItem.Workbook.SaveAs(ms);
				resultItem.Result = ms;
			}
			result.ResultItems.Add(resultItem);
		}
		return result;
	}

}
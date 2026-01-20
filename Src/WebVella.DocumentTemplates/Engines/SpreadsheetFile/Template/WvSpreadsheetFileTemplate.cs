using ClosedXML.Excel;
using System.Data;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
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
		}
		;
		if (!Template.CanRead)
		{
			throw new Exception("The template memory stream is closed");
		}
		try
		{
			result.Workbook = new XLWorkbook(Template);
		}
		catch
		{
			throw new Exception("The provided template memory stream cannot be opened as XLWorkbook. Wrong file?");
		}
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
			if (resultItem.Workbook is not null)
			{
				var ms = new MemoryStream();
				resultItem.Workbook.SaveAs(ms);
				resultItem.Result = ms;
			}
			result.ResultItems.Add(resultItem);
		}
		return result;
	}
	public List<ValidationErrorInfo> Validate(FileFormatVersions version = FileFormatVersions.Microsoft365)
	{
		if (Template is null) return new List<ValidationErrorInfo>();
		using SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(Template, false);
		OpenXmlValidator validator = new OpenXmlValidator(version);
		return validator.Validate(excelDoc).ToList();
	}	

}
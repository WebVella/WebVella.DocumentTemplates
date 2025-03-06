using DocumentFormat.OpenXml.Packaging;
using System.Data;
using System.Globalization;
using System.Text;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Engines.DocumentFile.Utility;
using WebVella.DocumentTemplates.Extensions;
namespace WebVella.DocumentTemplates.Engines.DocumentFile;
public class WvDocumentFileTemplate : WvTemplateBase
{
	public MemoryStream? Template { get; set; }
	public WvDocumentFileTemplateProcessResult Process(DataTable? dataSource, CultureInfo? culture = null)
	{
		if (culture == null) culture = new CultureInfo("en-US");
		if (dataSource is null) throw new ArgumentException("No datasource provided!", nameof(dataSource));

		var result = new WvDocumentFileTemplateProcessResult()
		{
			Template = Template,
			GroupByDataColumns = GroupDataByColumns,
			ResultItems = new()
		};
		if (Template is null)
		{
			result.ResultItems.Add(new WvDocumentFileTemplateProcessResultItem
			{
				WordDocument = null
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
			result.WordDocument = WordprocessingDocument.Open(result.Template!, true);
		}
		catch
		{
			throw new Exception("The provided template memory stream cannot be opened as XLWorkbook. Wrong file?");
		}
		var datasourceGroups = dataSource.GroupBy(GroupDataByColumns);
		foreach (var grouptedDs in datasourceGroups)
		{
			var resultMs = new MemoryStream();
			var resultItem = new WvDocumentFileTemplateProcessResultItem
			{
				WordDocument = WordprocessingDocument.Create(resultMs,
					DocumentFormat.OpenXml.WordprocessingDocumentType.Document),
				Result = resultMs,
				DataTable = grouptedDs
			};

			new WvDocumentFileEngineUtility().ProcessDocumentTemplate(
				result: result,
				resultItem: resultItem,
				dataSource: grouptedDs,
				culture: culture);
			resultItem.WordDocument.Save();
			result.ResultItems.Add(resultItem);
		}

		return result;
	}
}
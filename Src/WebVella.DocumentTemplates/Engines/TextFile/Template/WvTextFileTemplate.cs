using System.Data;
using System.Globalization;
using System.Text;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Engines.Text;
using WebVella.DocumentTemplates.Extensions;
namespace WebVella.DocumentTemplates.Engines.TextFile;
public class WvTextFileTemplate : WvTemplateBase
{
	public MemoryStream? Template { get; set; }
	public WvTextFileTemplateProcessResult Process(DataTable dataSource, CultureInfo? culture = null, Encoding? encoding = null)
	{
		if (culture == null) culture = new CultureInfo("en-US");
		if (encoding == null) encoding = Encoding.UTF8;
		if (dataSource is null) throw new ArgumentException("No datasource provided!", nameof(dataSource));
		var result = new WvTextFileTemplateProcessResult()
		{
			Template = Template,
			GroupByDataColumns = GroupDataByColumns,
			ResultItems = new()
		};
		if (Template is null)
		{
			result.ResultItems.Add(new WvTextFileTemplateProcessResultItem
			{
				Result = null
			});
			return result;
		};
		var datasourceGroups = dataSource.GroupBy(GroupDataByColumns);
		foreach (var grouptedDs in datasourceGroups)
		{
			var resultItem = new WvTextFileTemplateProcessResultItem()
			{
				NumberOfDataTableRows = grouptedDs.Rows.Count
			};
			var context = new WvTextFileTemplateProcessContext();
			try
			{
				var fileStringContent = encoding.GetString(result.Template!.ToArray() ?? new byte[0]);

				if (String.IsNullOrWhiteSpace(fileStringContent)) return result;

				var textTemplate = new WvTextTemplate
				{
					Template = fileStringContent.RemoveZeroBitSpaceCharacters()
				};

				WvTextTemplateProcessResult textTemplateResult = textTemplate.Process(grouptedDs, culture);

				if (textTemplateResult.ResultItems.Count == 0) continue;
				resultItem.Result = new MemoryStream(encoding.GetBytes(textTemplateResult.ResultItems[0].Result ?? String.Empty));
			}
			catch (Exception ex)
			{
				context.Errors.Add(ex.Message);
			}
			resultItem.Contexts.Add(context);
			result.ResultItems.Add(resultItem);
		}
		return result;
	}
}
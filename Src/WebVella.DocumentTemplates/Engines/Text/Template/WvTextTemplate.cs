using System.Data;
using System.Globalization;
using System.Text;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Extensions;
namespace WebVella.DocumentTemplates.Engines.Text;
public class WvTextTemplate : WvTemplateBase
{
	public string? Template { get; set; }
	public WvTextTemplateProcessResult Process(DataTable? dataSource, CultureInfo? culture = null)
	{
		if (culture == null) culture = new CultureInfo("en-US");
		if (dataSource is null) throw new ArgumentException("No datasource provided!", nameof(dataSource));

		var result = new WvTextTemplateProcessResult()
		{
			Template = Template,
			GroupByDataColumns = GroupDataByColumns,
			ResultItems = new()
		};
		if (String.IsNullOrWhiteSpace(Template))
		{
			result.ResultItems.Add(new WvTextTemplateProcessResultItem
			{
				Result = Template
			});
			return result;
		};

		var datasourceGroups = dataSource.GroupBy(GroupDataByColumns);

		foreach (var grouptedDs in datasourceGroups)
		{
			var resultItem = new WvTextTemplateProcessResultItem()
			{
				DataTable = grouptedDs
			};
			var resultContext = new WvTextTemplateProcessContext();
			try
			{
				var sb = new StringBuilder();
				var lines = Template.Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries);
				var endWithNewLine = Template.EndsWith(Environment.NewLine);
				if (lines.Count() == 1)
				{
					var tagProcessResult = new WvTemplateUtility().ProcessTemplateTag(lines.First(), grouptedDs, culture);
					if (tagProcessResult.Values.Count == 1)
					{
						sb.Append(tagProcessResult.Values[0]?.ToString());
					}
					else if (tagProcessResult.Values.Count > 1)
					{
						foreach (var value in tagProcessResult.Values)
						{
							if (endWithNewLine)
								sb.AppendLine(value?.ToString());
							else
								sb.Append(value?.ToString());
						}
					}
				}
				else if (lines.Count() > 1)
				{
					foreach (string line in lines)
					{
						var tagProcessResult = new WvTemplateUtility().ProcessTemplateTag(line, grouptedDs, culture);
						foreach (var value in tagProcessResult.Values)
						{
							sb.AppendLine(value?.ToString());
						}
					}
				}
				resultItem.Result = sb.ToString();
			}
			catch (Exception ex)
			{
				resultContext.Errors.Add(ex.Message);
			}
			resultItem.Contexts.Add(resultContext);
			result.ResultItems.Add(resultItem);
		}
		return result;
	}
}
using HtmlAgilityPack;
using System.Data;
using System.Globalization;
using System.Web;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Extensions;
namespace WebVella.DocumentTemplates.Engines.Html;
public class WvHtmlTemplate : WvTemplateBase
{
	public string? Template { get; set; }

	/// <summary>
	/// If you are using HTML editor do not forget to Decode the text 
	/// as editors encode ' and "
	/// Template = HttpUtility.HtmlDecode(Template);
	/// </summary>
	public WvHtmlTemplateProcessResult Process(DataTable? dataSource, CultureInfo? culture = null)
	{
		if (culture == null) culture = new CultureInfo("en-US");
		if (dataSource is null) throw new ArgumentException("No datasource provided!", nameof(dataSource));

		var result = new WvHtmlTemplateProcessResult()
		{
			Template = Template,
			GroupByDataColumns = GroupDataByColumns,
			ResultItems = new()
		};
		if (String.IsNullOrWhiteSpace(Template))
		{
			result.ResultItems.Add(new WvHtmlTemplateProcessResultItem
			{
				Result = Template
			});
			return result;
		};

		var datasourceGroups = dataSource.GroupBy(GroupDataByColumns);
		foreach (var grouptedDs in datasourceGroups)
		{
			HtmlDocument doc = new HtmlDocument();
			doc.OptionEmptyCollection = true; //Prevent from returning null
			HtmlDocument docResult = new HtmlDocument();
			doc.LoadHtml(result.Template);

			foreach (HtmlNode node in doc.DocumentNode.ChildNodes)
			{
				var nameLowered = node.Name.ToLowerInvariant();

				var tagProcessResult = WvTemplateUtility.ProcessTemplateTag(node.InnerHtml, grouptedDs, culture);
				foreach (var value in tagProcessResult.Values)
				{
					if (node.NodeType == HtmlNodeType.Text)
					{
						var textResult = docResult.CreateTextNode();
						textResult.InnerHtml = value.ToString();
						docResult.DocumentNode.AppendChild(textResult);
					}
					else
					{
						var divResult = docResult.CreateElement(node.Name);
						divResult.InnerHtml = value.ToString();
						docResult.DocumentNode.AppendChild(divResult);
					}
				}
			}
			result.ResultItems.Add(new WvHtmlTemplateProcessResultItem
			{
				Result = docResult.DocumentNode.InnerHtml
			});
		}
		return result;
	}
}
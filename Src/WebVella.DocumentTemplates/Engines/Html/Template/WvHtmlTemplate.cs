using HtmlAgilityPack;
using System.Data;
using System.Globalization;
using System.Web;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
namespace WebVella.DocumentTemplates.Engines.Html;
public class WvHtmlTemplate : WvTemplateBase
{
	public string? Template { get; set; }

	/// <summary>
	/// If you are using HTML editor do not forget to Decode the text 
	/// as editors encode ' and "
	/// Template = HttpUtility.HtmlDecode(Template);
	/// </summary>
	/// <param name="dataSource"></param>
	/// <param name="culture"></param>
	/// <returns></returns>
	/// <exception cref="ArgumentException"></exception>
	public WvHtmlTemplateResult? Process(DataTable dataSource, CultureInfo? culture = null)
	{
		if (culture == null) culture = new CultureInfo("en-US");
		if (dataSource is null) throw new ArgumentException("No datasource provided!", nameof(dataSource));

		var result = new WvHtmlTemplateResult()
		{
			Template = Template
		};
		if (String.IsNullOrWhiteSpace(Template))
		{
			result.Result = Template;
			return result;
		};
		HtmlDocument doc = new HtmlDocument();
		doc.OptionEmptyCollection = true; //Prevent from returning null
		HtmlDocument docResult = new HtmlDocument();
		doc.LoadHtml(result.Template);

		foreach (HtmlNode node in doc.DocumentNode.ChildNodes)
		{
			var nameLowered = node.Name.ToLowerInvariant();

			var tagProcessResult = WvTemplateUtility.ProcessTemplateTag(node.InnerHtml, dataSource, culture);
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
		result.Result = docResult.DocumentNode.InnerHtml;
		return result;
	}
}
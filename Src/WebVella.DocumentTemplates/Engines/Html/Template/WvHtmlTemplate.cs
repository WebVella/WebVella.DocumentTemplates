using DocumentFormat.OpenXml.Drawing.Diagrams;
using HtmlAgilityPack;
using System.Data;
using System.Globalization;
using System.Text;
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
			var resultItem = new WvHtmlTemplateProcessResultItem()
			{
				DataTable = grouptedDs
			};
			foreach (HtmlNode templateNode in doc.DocumentNode.ChildNodes)
			{
				//if root node is text
				if (templateNode.NodeType == HtmlNodeType.Text
					|| templateNode.NodeType == HtmlNodeType.Comment)
				{
					var resultNode = docResult.CreateElement("span");
					var context = new WvHtmlTemplateProcessContext();
					context.HtmlNode = resultNode;
					try
					{
						var tagProcessResult = new WvTemplateUtility().ProcessTemplateTag(templateNode.InnerHtml, grouptedDs, culture);
						if (tagProcessResult.ExpandCount > 0)
						{
							var sb = new StringBuilder();
							foreach (var valueObj in tagProcessResult.Values)
							{
								string? value = valueObj?.ToString();
								if (!String.IsNullOrWhiteSpace(value))
								{
									sb.Append(value);
								}
							}
							resultNode.InnerHtml = sb.ToString();
							//Some tags leave empty <span></span>
							if(!String.IsNullOrWhiteSpace(resultNode.InnerHtml))
								docResult.DocumentNode.AppendChild(resultNode);
						}
					}
					catch (Exception ex) {
						context.Errors.Add(ex.Message);
					}
					resultItem.Contexts.Add(context);
					continue;
				}
				else if (templateNode.NodeType == HtmlNodeType.Element)
				{
					processNode(templateNode, docResult.DocumentNode, docResult, grouptedDs, culture, resultItem.Contexts);
				}
			}
			resultItem.Result = docResult.DocumentNode.InnerHtml;
			result.ResultItems.Add(resultItem);
		}
		return result;
	}

	private void processNode(HtmlNode templateNode, HtmlNode resultParentNode, HtmlDocument resultDoc,
		DataTable grouptedDs, CultureInfo culture, List<WvHtmlTemplateProcessContext> contexts)
	{
		//Check if it is an empty node first
		if (String.IsNullOrWhiteSpace(templateNode.InnerHtml))
		{
			var resultNode = resultDoc.CreateElement(templateNode.Name);
			var context = new WvHtmlTemplateProcessContext();
			context.HtmlNode = templateNode;
			resultParentNode.AppendChild(resultNode);
			contexts.Add(context);
			return;
		}
		//Process if element with content
		var expandCount = 1;
		foreach (var childNode in templateNode.ChildNodes)
		{
			if (childNode.NodeType == HtmlNodeType.Text)
			{
				var tagProcessResult = new WvTemplateUtility().ProcessTemplateTag(childNode.InnerHtml, grouptedDs, culture);
				if (expandCount < tagProcessResult.ExpandCount) expandCount = tagProcessResult.ExpandCount;
			}
		}
		for (var i = 0; i < expandCount; i++)
		{
			var resultNode = resultDoc.CreateElement(templateNode.Name);
			var context = new WvHtmlTemplateProcessContext
			{
				HtmlNode = resultNode,
				Errors = new List<string>()
			};
			foreach (var childNode in templateNode.ChildNodes)
			{
				if (childNode.NodeType == HtmlNodeType.Text)
				{
					try
					{
						var tagProcessResult = new WvTemplateUtility().ProcessTemplateTag(childNode.InnerHtml, grouptedDs, culture);
						string? value = String.Empty;
						if (tagProcessResult.Values.Count > 0)
						{
							value = ((tagProcessResult.Values.Count >= i + 1) ? tagProcessResult.Values[i] : tagProcessResult.Values[0])?.ToString() ?? String.Empty;
						}
						var textResult = resultDoc.CreateTextNode();
						textResult.InnerHtml = value as string;
						resultNode.AppendChild(textResult);
					}
					catch (Exception ex)
					{
						context.Errors.Add(ex.Message);
					}

				}
				else
				{
					processNode(childNode, resultNode, resultDoc, grouptedDs, culture, contexts);
				}
			}
			resultParentNode.AppendChild(resultNode);
			contexts.Add(context);
		}
	}
}
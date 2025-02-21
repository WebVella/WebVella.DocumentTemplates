using HtmlAgilityPack;
using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.Html;
public class WvHtmlTemplateProcessContext : WvTemplateProcessContextBase
{
	public HtmlNode? HtmlNode { get; set; }
}
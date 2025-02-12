using HtmlAgilityPack;
using Markdig;

namespace WebVella.DocumentTemplates.Core.Utility;
public static partial class WvTemplateUtility
{
	public static string ConvertPlainTextToHtml(string text)
	{
		if (String.IsNullOrWhiteSpace(text)) return text;

		var pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();
		return Markdown.ToHtml(text, pipeline).TrimStart(Environment.NewLine.ToCharArray()).TrimEnd(Environment.NewLine.ToCharArray());
	}

	public static string ConvertHtmlToPlainText(string html)
	{
		if (String.IsNullOrWhiteSpace(html)) return html;

		try
		{
			if (string.IsNullOrWhiteSpace(html))
				return string.Empty;

			HtmlDocument doc = new HtmlDocument();
			doc.LoadHtml(html);

			StringWriter sw = new StringWriter();
			ConvertTo(doc.DocumentNode, sw);
			sw.Flush();
			return sw.ToString().TrimStart(Environment.NewLine.ToCharArray()).TrimEnd(Environment.NewLine.ToCharArray());
		}
		catch
		{
			return string.Empty;
		}
	}

	private static void ConvertTo(HtmlNode node, TextWriter outText)
	{
		string html;
		switch (node.NodeType)
		{
			case HtmlNodeType.Comment:
				// don't output comments
				break;

			case HtmlNodeType.Document:
				ConvertContentTo(node, outText);
				break;

			case HtmlNodeType.Text:
				// script and style must not be output
				string parentName = node.ParentNode.Name;
				if ((parentName == "script") || (parentName == "style"))
					break;

				// get text
				html = ((HtmlTextNode)node).Text;

				// is it in fact a special closing node output as text?
				if (HtmlNode.IsOverlappedClosingElement(html))
					break;

				// check the text is meaningful and not a bunch of white spaces
				if (html.Trim().Length > 0)
				{
					outText.Write(HtmlEntity.DeEntitize(html));
				}

				break;

			case HtmlNodeType.Element:
				switch (node.Name)
				{
					case "p":
						// treat paragraphs as crlf
						outText.Write(Environment.NewLine);
						break;
					case "br":
						outText.Write(Environment.NewLine);
						break;
					case "a":
						HtmlAttribute att = node.Attributes["href"];
						outText.Write($"<{att.Value}>");
						break;
				}

				if (node.HasChildNodes)
				{
					ConvertContentTo(node, outText);
				}

				break;
		}
	}

	private static void ConvertContentTo(HtmlNode node, TextWriter outText)
	{
		foreach (HtmlNode subnode in node.ChildNodes)
		{
			ConvertTo(subnode, outText);
		}
	}

}

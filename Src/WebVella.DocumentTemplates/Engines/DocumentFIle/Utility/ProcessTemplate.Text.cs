using System.Data;
using System.Globalization;
using WebVella.DocumentTemplates.Engines.Text;
using Word = DocumentFormat.OpenXml.Wordprocessing;
using System.Text;
using DocumentFormat.OpenXml;

namespace WebVella.DocumentTemplates.Engines.DocumentFile.Utility;
public partial class WvDocumentFileEngineUtility
{
	private List<OpenXmlElement> _processDocumentText(Word.Text template,
		DataTable dataSource, CultureInfo culture)
	{
		Word.Text resultEl = (Word.Text)template.CloneNode(true);
		resultEl.Space = SpaceProcessingModeValues.Preserve;
		resultEl.Text = String.Empty;
		if (String.IsNullOrWhiteSpace(template.InnerText)) return [resultEl];

		var textTemplate = new WvTextTemplate
		{
			Template = template.InnerText
		};
		var textTemplateResults = textTemplate.Process(dataSource, culture);
		var sb = new StringBuilder();
		foreach (var item in textTemplateResults.ResultItems)
		{
			if(String.IsNullOrWhiteSpace(item.Result)) continue;
			sb.Append(item.Result!);
		}

		var resultText = sb.ToString();
		if(String.IsNullOrWhiteSpace(resultText)) return [resultEl];
		if (resultText.Contains(Environment.NewLine))
		{
			var result = new List<OpenXmlElement>();
			var lines = resultText.Split(Environment.NewLine);
			var count = 1;
			foreach (var line in lines)
			{
				Word.Text nodeEl = (Word.Text)template.CloneNode(true);
				nodeEl.Text = line;
				result.Add(nodeEl);
				if (count < lines.Length)
					result.Add(new Word.Break());

				count++;
			}

			return result;
		}
		else
		{
			resultEl.Text = sb.ToString();
			return [resultEl];
		}
	}
}

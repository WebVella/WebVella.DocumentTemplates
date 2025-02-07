using System.Data;
using System.Globalization;
using System.Text;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
namespace WebVella.DocumentTemplates.Engines.Text;
public class WvTextTemplate : WvTemplateBase
{
	public string? Template { get; set; }
	public WvTextTemplateResult? Process(DataTable? dataSource, CultureInfo? culture = null)
	{
		if (culture == null) culture = new CultureInfo("en-US");
		if (dataSource is null) throw new ArgumentException("No datasource provided!",nameof(dataSource));

		var result = new WvTextTemplateResult(){
			Template = Template
		};
		if(String.IsNullOrWhiteSpace(Template)){ 
			result.Result = Template;
			return result;
		};

		var sb = new StringBuilder();
		var lines = Template.Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries);
		var endWithNewLine = Template.EndsWith(Environment.NewLine);
		if (lines.Count() == 1)
		{
			var tagProcessResult = WvTemplateUtility.ProcessTemplateTag(lines.First(), dataSource, culture);
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
				var tagProcessResult =  WvTemplateUtility.ProcessTemplateTag(line, dataSource, culture);
				foreach (var value in tagProcessResult.Values)
				{
					sb.AppendLine(value?.ToString());
				}
			}
		}
		result.Result = sb.ToString();

		return result;
	}
}
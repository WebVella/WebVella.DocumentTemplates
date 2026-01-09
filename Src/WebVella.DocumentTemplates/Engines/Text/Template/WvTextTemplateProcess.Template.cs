using System.Data;
using System.Globalization;
using System.Text;
using WebVella.DocumentTemplates.Core.Utility;

namespace WebVella.DocumentTemplates.Engines.Text;

public partial class WvTextTemplateProcess
{
    public string? ProcessTemplate(string? processedTemplate, DataTable dataSource, CultureInfo culture)
    {
        if(string.IsNullOrWhiteSpace(processedTemplate)) return  processedTemplate;
        var sb = new StringBuilder();
        var lines = processedTemplate.Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries);
        var endWithNewLine = processedTemplate.EndsWith(Environment.NewLine);
        if (lines.Count() == 1)
        {
            var tagProcessResult =
                new WvTemplateUtility().ProcessTemplateTag(lines.First(), dataSource, culture);
            if (tagProcessResult.Values.Count == 1)
            {
                sb.Append(tagProcessResult.Values[0]);
            }
            else if (tagProcessResult.Values.Count > 1)
            {
                foreach (var value in tagProcessResult.Values)
                {
                    if (endWithNewLine)
                        sb.AppendLine(value.ToString());
                    else
                        sb.Append(value);
                }
            }
        }
        else if (lines.Count() > 1)
        {
            foreach (string line in lines)
            {
                var tagProcessResult =
                    new WvTemplateUtility().ProcessTemplateTag(line, dataSource, culture);
                foreach (var value in tagProcessResult.Values)
                {
                    sb.AppendLine(value.ToString());
                }
            }
        }

        return sb.ToString();
    }
}
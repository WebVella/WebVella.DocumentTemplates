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
        }

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
                var processor = new WvTextTemplateProcess();
                var processedTemplate = Template;
                processedTemplate = processor.ProcessConditions(processedTemplate, grouptedDs, culture);                
                processedTemplate = processor.ProcessInlineTemplates(processedTemplate, grouptedDs, culture);
                processedTemplate = processor.ProcessTemplate(processedTemplate, grouptedDs, culture);
                resultItem.Result = processedTemplate;
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
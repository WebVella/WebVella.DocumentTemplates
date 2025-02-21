using ClosedXML.Excel;
using System.Data;
using System.Globalization;
using System.Text;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Engines.Email.Models;
using WebVella.DocumentTemplates.Engines.ExcelFile;
using WebVella.DocumentTemplates.Engines.Html;
using WebVella.DocumentTemplates.Engines.Text;
using WebVella.DocumentTemplates.Engines.TextFile;
using WebVella.DocumentTemplates.Extensions;
namespace WebVella.DocumentTemplates.Engines.Email;
public class WvEmailTemplate : WvTemplateBase
{
	public WvEmail? Template { get; set; }

	public WvEmailTemplateProcessResult Process(DataTable? dataSource, CultureInfo? culture = null)
	{
		if (culture == null) culture = new CultureInfo("en-US");
		if (dataSource is null) throw new ArgumentException("No datasource provided!", nameof(dataSource));
		var result = new WvEmailTemplateProcessResult()
		{
			Template = Template,
			GroupByDataColumns = GroupDataByColumns,
			ResultItems = new()
		};
		if (Template is null)
		{
			result.ResultItems.Add(new WvEmailTemplateProcessResultItem
			{
				Result = null
			});
			return result;
		};

		var datasourceGroups = dataSource.GroupBy(GroupDataByColumns);
		foreach (var grouptedDs in datasourceGroups)
		{
			var resultItem = new WvEmailTemplateProcessResultItem
			{
				Result = new WvEmail(),
				NumberOfDataTableRows = grouptedDs.Rows.Count
			};
			var context = new WvEmailTemplateProcessContext();
			try
			{
				ProcessEmailSender(Template, resultItem, grouptedDs, culture);
				ProcessEmailRecipients(Template, resultItem, grouptedDs, culture);
				ProcessEmailCcRecipients(Template, resultItem, grouptedDs, culture);
				ProcessEmailBccRecipients(Template, resultItem, grouptedDs, culture);
				ProcessEmailSubject(Template, resultItem, grouptedDs, culture);
				ProcessEmailHtmlContent(Template, resultItem, grouptedDs, culture);
				if (String.IsNullOrWhiteSpace(Template.TextContent) && !String.IsNullOrWhiteSpace(Template.HtmlContent))
				{
					resultItem.Result.TextContent = new WvTemplateUtility().ConvertHtmlToPlainText(Template.HtmlContent);
				}
				else
				{
					ProcessEmailTextContent(Template, resultItem, grouptedDs, culture);
				}
				if (String.IsNullOrWhiteSpace(Template.HtmlContent) && !String.IsNullOrWhiteSpace(Template.TextContent))
				{
					resultItem.Result.HtmlContent = new WvTemplateUtility().ConvertPlainTextToHtml(Template.TextContent);
				}

				int attachmentIndex = 0;
				foreach (var item in Template.AttachmentItems)
				{
					var itemDataSources = grouptedDs.GroupBy(item.GroupDataByColumns);
					int attachmentDsIndex = 0;
					foreach (var itemDataSource in itemDataSources)
					{
						if (item.Type == WvEmailAttachmentType.TextFile)
						{
							var attachment = ProcessEmailTextAttachment(
								template: item.Content,
								fileName: item.Filename ?? $"file{(attachmentIndex == 0 ? "" : attachmentIndex.ToString())}.txt",
								dsIndex: attachmentDsIndex,
								dataSource: itemDataSource,
								culture: culture,
								encoding: item.Encoding);
							if (attachment is not null) resultItem.Result.AttachmentItems.Add(attachment);
						}
						else if (item.Type == WvEmailAttachmentType.ExcelFile)
						{
							var attachment = ProcessEmailExcelAttachment(
								template: item.Content,
								fileName: item.Filename ?? $"file{(attachmentIndex == 0 ? "" : attachmentIndex.ToString())}.xlsx",
								dsIndex: attachmentDsIndex,
								dataSource: itemDataSource,
								culture: culture);
							if (attachment is not null) resultItem.Result.AttachmentItems.Add(attachment);
						}
					}
				}
			}
			catch (Exception ex) {
				context.Errors.Add(ex.Message);
			}
			resultItem.Contexts.Add(context);
			result.ResultItems.Add(resultItem);

		}
		return result;
	}

	public void ProcessEmailSender(WvEmail template, WvEmailTemplateProcessResultItem resultItem, DataTable dataSource, CultureInfo culture)
	{
		if (resultItem is null) throw new Exception("No result provided!");
		if (dataSource is null) throw new Exception("No datasource provided!");
		if (template is null) throw new Exception("No Template provided!");
		if (resultItem.Result is null) throw new Exception("Result should not be null!");
		var senderTemplate = new WvTextTemplate
		{
			Template = template.Sender
		};
		var firstRowOnly = dataSource.CreateAsNew(new List<int> { 0 });
		var senderTemplateResults = senderTemplate.Process(firstRowOnly, culture);
		resultItem.Result.Sender = senderTemplateResults.ResultItems.Count > 0 ? senderTemplateResults.ResultItems[0].Result : template.Sender;
	}
	public void ProcessEmailRecipients(WvEmail template, WvEmailTemplateProcessResultItem resultItem, DataTable dataSource, CultureInfo culture)
	{
		if (resultItem is null) throw new Exception("No result provided!");
		if (dataSource is null) throw new Exception("No datasource provided!");
		if (template is null) throw new Exception("No Template provided!");
		if (resultItem.Result is null) throw new Exception("Result should not be null!");

		resultItem.Result.Recipients = ProcessEmailListTemplate(template.Recipients, dataSource, culture);
	}
	public void ProcessEmailCcRecipients(WvEmail template, WvEmailTemplateProcessResultItem resultItem, DataTable dataSource, CultureInfo culture)
	{
		if (resultItem is null) throw new Exception("No result provided!");
		if (dataSource is null) throw new Exception("No datasource provided!");
		if (template is null) throw new Exception("No Template provided!");
		if (resultItem.Result is null) throw new Exception("Result should not be null!");

		resultItem.Result.CcRecipients = ProcessEmailListTemplate(template.CcRecipients, dataSource, culture);
	}
	public void ProcessEmailBccRecipients(WvEmail template, WvEmailTemplateProcessResultItem resultItem, DataTable dataSource, CultureInfo culture)
	{
		if (resultItem is null) throw new Exception("No result provided!");
		if (dataSource is null) throw new Exception("No datasource provided!");
		if (template is null) throw new Exception("No Template provided!");
		if (resultItem.Result is null) throw new Exception("Result should not be null!");

		resultItem.Result.BccRecipients = ProcessEmailListTemplate(template.BccRecipients, dataSource, culture);
	}
	public void ProcessEmailSubject(WvEmail template, WvEmailTemplateProcessResultItem resultItem, DataTable dataSource, CultureInfo culture)
	{
		if (resultItem is null) throw new Exception("No result provided!");
		if (dataSource is null) throw new Exception("No datasource provided!");
		if (template is null) throw new Exception("No Template provided!");
		if (resultItem.Result is null) throw new Exception("Result should not be null!");
		var senderTemplate = new WvTextTemplate
		{
			Template = template.Subject
		};
		var senderTemplateResults = senderTemplate.Process(dataSource, culture);
		resultItem.Result.Subject = senderTemplateResults.ResultItems.Count > 0 ? senderTemplateResults.ResultItems[0].Result : template.Subject;
	}
	public void ProcessEmailHtmlContent(WvEmail template, WvEmailTemplateProcessResultItem resultItem, DataTable dataSource, CultureInfo culture)
	{
		if (resultItem is null) throw new Exception("No result provided!");
		if (dataSource is null) throw new Exception("No datasource provided!");
		if (template is null) throw new Exception("No Template provided!");
		if (resultItem.Result is null) throw new Exception("Result should not be null!");
		var senderTemplate = new WvHtmlTemplate
		{
			Template = template.HtmlContent
		};
		var senderTemplateResults = senderTemplate.Process(dataSource, culture);
		resultItem.Result.HtmlContent = senderTemplateResults.ResultItems.Count > 0 ? senderTemplateResults.ResultItems[0].Result : template.HtmlContent;
	}
	public void ProcessEmailTextContent(WvEmail template, WvEmailTemplateProcessResultItem resultItem, DataTable dataSource, CultureInfo culture)
	{
		if (resultItem is null) throw new Exception("No result provided!");
		if (dataSource is null) throw new Exception("No datasource provided!");
		if (template is null) throw new Exception("No Template provided!");
		if (resultItem.Result is null) throw new Exception("Result should not be null!");
		var senderTemplate = new WvTextTemplate
		{
			Template = template.TextContent
		};
		var senderTemplateResults = senderTemplate.Process(dataSource, culture);
		resultItem.Result.TextContent = senderTemplateResults.ResultItems.Count > 0 ? senderTemplateResults.ResultItems[0].Result : template.TextContent;
	}
	public string ProcessEmailListTemplate(string? template, DataTable dataSource, CultureInfo culture)
	{
		if (String.IsNullOrWhiteSpace(template)) return String.Empty;
		var senderTemplate = new WvTextTemplate
		{
			Template = template
		};
		var senderTemplateResults = senderTemplate.Process(dataSource, culture);

		if (senderTemplateResults.ResultItems.Count == 0)
		{
			return String.Empty;
		}

		// For convinience if the input has no flow separator ";", join them with code
		// This is why we first check if separator is used
		if (senderTemplateResults.ResultItems.Count == 1)
		{
			if (String.IsNullOrWhiteSpace(senderTemplateResults.ResultItems[0].Result))
			{
				return String.Empty;
			}
			var atCount = senderTemplateResults.ResultItems[0].Result!.Count(x => x == '@');
			var separatorCount = senderTemplateResults.ResultItems[0].Result!.Count(x => x == ';');
			if (atCount <= 1 || separatorCount > 0)
			{
				return senderTemplateResults.ResultItems[0].Result!;
			}
		}
		//If no separator process row by row
		var resultList = new List<string>();
		for (int i = 0; i < dataSource.Rows.Count; i++)
		{
			var singleRowDs = dataSource.CreateAsNew(new List<int> { i });
			var singleRowResult = senderTemplate.Process(singleRowDs, culture);
			if (singleRowResult.ResultItems.Count == 0
				|| String.IsNullOrWhiteSpace(singleRowResult.ResultItems[0].Result))
				continue;

			resultList.Add(singleRowResult.ResultItems[0].Result!);

		}
		return String.Join(";", resultList);
	}

	public WvEmailAttachment? ProcessEmailTextAttachment(byte[]? template, string fileName, int dsIndex, DataTable dataSource, CultureInfo culture, Encoding? encoding)
	{
		if (template is null) return null;
		WvEmailAttachment? attachment = null;
		var attachmentTemplate = new WvTextFileTemplate
		{
			Template = template
		};

		var attachmentTemplateResult = attachmentTemplate.Process(dataSource, culture, encoding);
		if (attachmentTemplateResult.ResultItems.Count == 0
			|| attachmentTemplateResult.ResultItems[0].Result is null)
		{
			return null;
		}
		var ext = Path.GetExtension(fileName);
		var name = Path.GetFileNameWithoutExtension(fileName);
		attachment = new WvEmailAttachment
		{
			Encoding = encoding,
			GroupDataByColumns = new(),
			Type = WvEmailAttachmentType.TextFile,
			Filename = dsIndex == 0 ? fileName : $"{name}-{dsIndex}{ext}",
			Content = attachmentTemplateResult.ResultItems[0].Result
		};


		return attachment;
	}

	public WvEmailAttachment? ProcessEmailExcelAttachment(byte[]? template, string fileName, int dsIndex, DataTable dataSource, CultureInfo culture)
	{
		if (template is null) return null;
		WvEmailAttachment? attachment = null;
		XLWorkbook? workbook = null;
		using (MemoryStream memoryStream = new MemoryStream(template))
		{
			// Use the stream as needed
			workbook = new XLWorkbook(memoryStream);
		}
		if (workbook is null) return null;

		var attachmentTemplate = new WvExcelFileTemplate
		{
			Template = workbook
		};

		var attachmentTemplateResult = attachmentTemplate.Process(dataSource, culture);
		if (attachmentTemplateResult.ResultItems.Count == 0
					|| attachmentTemplateResult.ResultItems[0].Result is null)
		{
			return null;
		}
		var ext = Path.GetExtension(fileName);
		var name = Path.GetFileNameWithoutExtension(fileName);
		byte[]? content = null;
		using (var ms = new MemoryStream())
		{
			attachmentTemplateResult.ResultItems[0]!.Result!.SaveAs(ms);
			content = ms?.ToArray();
		}
		if (content is null) return null;

		attachment = new WvEmailAttachment
		{
			GroupDataByColumns = new(),
			Type = WvEmailAttachmentType.ExcelFile,
			Filename = dsIndex == 0 ? fileName : $"{name}-{dsIndex}{ext}",
			Content = content
		};


		return attachment;
	}
}
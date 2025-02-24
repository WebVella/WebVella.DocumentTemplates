using ClosedXML.Excel;
using System.Data;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Engines.Email;
using WebVella.DocumentTemplates.Engines.Email.Models;
using WebVella.DocumentTemplates.Extensions;
using WebVella.DocumentTemplates.Tests.Models;

namespace WebVella.DocumentTemplates.Tests.Engines;
public class EmailEngineTests : TestBase
{
	private static readonly object locker = new();
	public EmailEngineTests() : base() { }

	#region << Arguments >>
	[Fact]
	public void ShouldHaveDataSource()
	{
		//Given
		var template = new WvEmailTemplate
		{
			Template = new WvEmail
			{
				Sender = "test@test.com"
			}
		};
		WvEmailTemplateProcessResult? result = null;
		//When
		var action = () => result = template.Process(null, DefaultCulture);
		//Then
		var ex = Record.Exception(action);
		Assert.NotNull(ex);
		Assert.IsType<ArgumentException>(ex);
		var argEx = (ArgumentException)ex;
		Assert.Equal("dataSource", argEx.ParamName);
		Assert.StartsWith("No datasource provided!", argEx.Message);
	}
	#endregion

	#region << Sender >>
	[Fact]
	public void Sender_HardcodedEmail()
	{
		//Given
		var template = new WvEmailTemplate
		{
			Template = new WvEmail
			{
				Sender = "test@test.com"
			}
		};
		var ds = EmailData;

		//When
		WvEmailTemplateProcessResult? result = template.Process(ds, DefaultCulture); ;

		//Then
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.NotNull(result.ResultItems[0].Result);
		Assert.Equal(template.Template.Sender, result.ResultItems[0].Result!.Sender);
	}
	[Fact]
	public void Sender_SingleData()
	{
		//Given
		var template = new WvEmailTemplate
		{
			Template = new WvEmail
			{
				Sender = "{{sender_email[0]}}"
			}
		};
		var ds = EmailData;

		//When
		WvEmailTemplateProcessResult? result = template.Process(ds, DefaultCulture); ;

		//Then
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.NotNull(result.ResultItems[0].Result);
		Assert.Equal(ds.Rows[0]["sender_email"], result.ResultItems[0].Result!.Sender);
	}

	[Fact]
	public void Sender_ShouldReturnSingle()
	{
		//Given
		var template = new WvEmailTemplate
		{
			Template = new WvEmail
			{
				Sender = "{{sender_email}}"
			}
		};
		var ds = EmailData;

		//When
		WvEmailTemplateProcessResult? result = template.Process(ds, DefaultCulture); ;

		//Then
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.NotNull(result.ResultItems[0].Result);
		Assert.Equal(ds.Rows[0]["sender_email"], result.ResultItems[0].Result!.Sender);
	}
	#endregion

	#region << Recipients >>
	[Fact]
	public void Recipients_HardcodedEmail()
	{
		//Given
		var template = new WvEmailTemplate
		{
			Template = new WvEmail
			{
				Recipients = "test@test.com"
			}
		};
		var ds = EmailData;

		//When
		WvEmailTemplateProcessResult? result = template.Process(ds, DefaultCulture);

		//Then
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.NotNull(result.ResultItems[0].Result);
		Assert.Equal(template.Template.Recipients, result.ResultItems[0].Result!.Recipients);
	}

	[Fact]
	public void Recipients_TemplateSingle()
	{
		//Given
		var template = new WvEmailTemplate
		{
			Template = new WvEmail
			{
				Recipients = "{{recipient_email[0]}}"
			}
		};
		var ds = EmailData;

		//When
		WvEmailTemplateProcessResult? result = template.Process(ds, DefaultCulture);

		//Then
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.NotNull(result.ResultItems[0].Result);
		Assert.Equal(EmailData.Rows[0]["recipient_email"], result.ResultItems[0].Result!.Recipients);
	}
	[Fact]
	public void Recipients_TemplateWithSeparator()
	{
		//Given
		var template = new WvEmailTemplate
		{
			Template = new WvEmail
			{
				Recipients = "{{recipient_email(S=';')}}"
			}
		};
		var ds = EmailData.CreateAsNew(new List<int> { 0, 1 });

		//When
		WvEmailTemplateProcessResult? result = template.Process(ds, DefaultCulture);

		//Then
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.NotNull(result.ResultItems[0].Result);
		Assert.Equal($"{EmailData.Rows[0]["recipient_email"]};{EmailData.Rows[1]["recipient_email"]}", result.ResultItems[0].Result!.Recipients);
	}

	[Fact]
	public void Recipients_TemplateWithoutSeparatorShouldWorkToo()
	{
		//Given
		var template = new WvEmailTemplate
		{
			Template = new WvEmail
			{
				Recipients = "{{recipient_email}}"
			}
		};
		var ds = EmailData.CreateAsNew(new List<int> { 0, 1 });

		//When
		WvEmailTemplateProcessResult? result = template.Process(ds, DefaultCulture);

		//Then
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.NotNull(result.ResultItems[0].Result);
		Assert.Equal($"{EmailData.Rows[0]["recipient_email"]};{EmailData.Rows[1]["recipient_email"]}", result.ResultItems[0].Result!.Recipients);
	}
	#endregion

	#region << CcRecipients >>
	[Fact]
	public void CcRecipients_HardcodedEmail()
	{
		//Given
		var template = new WvEmailTemplate
		{
			Template = new WvEmail
			{
				CcRecipients = "test@test.com"
			}
		};
		var ds = EmailData;

		//When
		WvEmailTemplateProcessResult? result = template.Process(ds, DefaultCulture);

		//Then
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.NotNull(result.ResultItems[0].Result);
		Assert.Equal(template.Template.CcRecipients, result.ResultItems[0].Result!.CcRecipients);
	}

	[Fact]
	public void CcRecipients_TemplateSingle()
	{
		//Given
		var template = new WvEmailTemplate
		{
			Template = new WvEmail
			{
				CcRecipients = "{{recipient_email[0]}}"
			}
		};
		var ds = EmailData;

		//When
		WvEmailTemplateProcessResult? result = template.Process(ds, DefaultCulture);

		//Then
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.NotNull(result.ResultItems[0].Result);
		Assert.Equal(EmailData.Rows[0]["recipient_email"], result.ResultItems[0].Result!.CcRecipients);
	}
	[Fact]
	public void CcRecipients_TemplateWithSeparator()
	{
		//Given
		var template = new WvEmailTemplate
		{
			Template = new WvEmail
			{
				CcRecipients = "{{recipient_email(S=';')}}"
			}
		};
		var ds = EmailData.CreateAsNew(new List<int> { 0, 1 });

		//When
		WvEmailTemplateProcessResult? result = template.Process(ds, DefaultCulture);

		//Then
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.NotNull(result.ResultItems[0].Result);
		Assert.Equal($"{EmailData.Rows[0]["recipient_email"]};{EmailData.Rows[1]["recipient_email"]}", result.ResultItems[0].Result!.CcRecipients);
	}

	[Fact]
	public void CcRecipients_TemplateWithoutSeparatorShouldWorkToo()
	{
		//Given
		var template = new WvEmailTemplate
		{
			Template = new WvEmail
			{
				CcRecipients = "{{recipient_email}}"
			}
		};
		var ds = EmailData.CreateAsNew(new List<int> { 0, 1 });

		//When
		WvEmailTemplateProcessResult? result = template.Process(ds, DefaultCulture);

		//Then
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.NotNull(result.ResultItems[0].Result);
		Assert.Equal($"{EmailData.Rows[0]["recipient_email"]};{EmailData.Rows[1]["recipient_email"]}", result.ResultItems[0].Result!.CcRecipients);
	}
	#endregion

	#region << BccRecipients >>
	[Fact]
	public void BccRecipients_HardcodedEmail()
	{
		//Given
		var template = new WvEmailTemplate
		{
			Template = new WvEmail
			{
				BccRecipients = "test@test.com"
			}
		};
		var ds = EmailData;

		//When
		WvEmailTemplateProcessResult? result = template.Process(ds, DefaultCulture);

		//Then
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.NotNull(result.ResultItems[0].Result);
		Assert.Equal(template.Template.BccRecipients, result.ResultItems[0].Result!.BccRecipients);
	}

	[Fact]
	public void BccRecipients_TemplateSingle()
	{
		//Given
		var template = new WvEmailTemplate
		{
			Template = new WvEmail
			{
				BccRecipients = "{{recipient_email[0]}}"
			}
		};
		var ds = EmailData;

		//When
		WvEmailTemplateProcessResult? result = template.Process(ds, DefaultCulture);

		//Then
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.NotNull(result.ResultItems[0].Result);
		Assert.Equal(EmailData.Rows[0]["recipient_email"], result.ResultItems[0].Result!.BccRecipients);
	}
	[Fact]
	public void BccRecipients_TemplateWithSeparator()
	{
		//Given
		var template = new WvEmailTemplate
		{
			Template = new WvEmail
			{
				BccRecipients = "{{recipient_email(S=';')}}"
			}
		};
		var ds = EmailData.CreateAsNew(new List<int> { 0, 1 });

		//When
		WvEmailTemplateProcessResult? result = template.Process(ds, DefaultCulture);

		//Then
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.NotNull(result.ResultItems[0].Result);
		Assert.Equal($"{EmailData.Rows[0]["recipient_email"]};{EmailData.Rows[1]["recipient_email"]}", result.ResultItems[0].Result!.BccRecipients);
	}

	[Fact]
	public void BccRecipients_TemplateWithoutSeparatorShouldWorkToo()
	{
		//Given
		var template = new WvEmailTemplate
		{
			Template = new WvEmail
			{
				BccRecipients = "{{recipient_email}}"
			}
		};
		var ds = EmailData.CreateAsNew(new List<int> { 0, 1 });

		//When
		WvEmailTemplateProcessResult? result = template.Process(ds, DefaultCulture);

		//Then
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.NotNull(result.ResultItems[0].Result);
		Assert.Equal($"{EmailData.Rows[0]["recipient_email"]};{EmailData.Rows[1]["recipient_email"]}", result.ResultItems[0].Result!.BccRecipients);
	}
	#endregion

	#region << Subject >>
	[Fact]
	public void Subject_Hardcoded()
	{
		//Given
		var template = new WvEmailTemplate
		{
			Template = new WvEmail
			{
				Subject = "some text"
			}
		};
		var ds = EmailData;

		//When
		WvEmailTemplateProcessResult? result = template.Process(ds, DefaultCulture); ;

		//Then
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.NotNull(result.ResultItems[0].Result);
		Assert.Equal(template.Template.Subject, result.ResultItems[0].Result!.Subject);
	}
	[Fact]
	public void Subject_SingleData()
	{
		//Given
		var template = new WvEmailTemplate
		{
			Template = new WvEmail
			{
				Subject = "{{subject[0]}}"
			}
		};
		var ds = EmailData;

		//When
		WvEmailTemplateProcessResult? result = template.Process(ds, DefaultCulture); ;

		//Then
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.NotNull(result.ResultItems[0].Result);
		Assert.Equal(ds.Rows[0]["subject"], result.ResultItems[0].Result!.Subject);
	}

	[Fact]
	public void Subject_MultiData()
	{
		//Given
		var template = new WvEmailTemplate
		{
			Template = new WvEmail
			{
				Subject = "{{subject(S=',')}}"
			}
		};
		var ds = EmailData.CreateAsNew(new List<int> { 0, 1 });

		//When
		WvEmailTemplateProcessResult? result = template.Process(ds, DefaultCulture); ;

		//Then
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.NotNull(result.ResultItems[0].Result);
		Assert.Equal($"{ds.Rows[0]["subject"]},{ds.Rows[1]["subject"]}", result.ResultItems[0].Result!.Subject);
	}
	#endregion

	#region << HTML content >>
	[Fact]
	public void HtmlContent_Hardcoded()
	{
		//Given
		var template = new WvEmailTemplate
		{
			Template = new WvEmail
			{
				HtmlContent = "some text"
			}
		};
		var ds = EmailData;

		//When
		WvEmailTemplateProcessResult? result = template.Process(ds, DefaultCulture); ;

		//Then
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.NotNull(result.ResultItems[0].Result);
		Assert.Equal($"<span>{template.Template.HtmlContent}</span>", result.ResultItems[0].Result!.HtmlContent);
		Assert.Equal(template.Template.HtmlContent, result.ResultItems[0].Result!.TextContent);
	}
	[Fact]
	public void HtmlContent_Hardcoded2()
	{
		//Given
		var template = new WvEmailTemplate
		{
			Template = new WvEmail
			{
				HtmlContent = "Hello<br/>test"
			}
		};
		var ds = EmailData;

		//When
		WvEmailTemplateProcessResult? result = template.Process(ds, DefaultCulture); ;

		//Then
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.NotNull(result.ResultItems[0].Result);
		Assert.Equal("<span>Hello</span><br><span>test</span>", result.ResultItems[0].Result!.HtmlContent);
		Assert.Equal($"Hello{Environment.NewLine}test", result.ResultItems[0].Result!.TextContent);
	}
	[Fact]
	public void HtmlContent_Hardcoded3()
	{
		//Given
		var template = new WvEmailTemplate
		{
			Template = new WvEmail
			{
				HtmlContent = "<p>Hello</p><p>test</p>"
			}
		};
		var ds = EmailData;

		//When
		WvEmailTemplateProcessResult? result = template.Process(ds, DefaultCulture); ;

		//Then
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.NotNull(result.ResultItems[0].Result);
		Assert.Equal(template.Template.HtmlContent, result.ResultItems[0].Result!.HtmlContent);
		Assert.Equal($"Hello{Environment.NewLine}test", result.ResultItems[0].Result!.TextContent);
	}
	[Fact]
	public void HtmlContent_DoNotReplaceTextIfProvided()
	{
		//Given
		var template = new WvEmailTemplate
		{
			Template = new WvEmail
			{
				HtmlContent = "<p>Hello</p><p>test</p>",
				TextContent = "test"
			}
		};
		var ds = EmailData;

		//When
		WvEmailTemplateProcessResult? result = template.Process(ds, DefaultCulture); ;

		//Then
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.NotNull(result.ResultItems[0].Result);
		Assert.Equal(template.Template.HtmlContent, result.ResultItems[0].Result!.HtmlContent);
		Assert.Equal(template.Template.TextContent, result.ResultItems[0].Result!.TextContent);
	}
	#endregion
	#region <<TextContent>>
	[Fact]
	public void TextContent_Hardcoded()
	{
		//Given
		var template = new WvEmailTemplate
		{
			Template = new WvEmail
			{
				TextContent = "some text"
			}
		};
		var ds = EmailData;

		//When
		WvEmailTemplateProcessResult? result = template.Process(ds, DefaultCulture); ;

		//Then
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.NotNull(result.ResultItems[0].Result);
		Assert.Equal(template.Template.TextContent, result.ResultItems[0].Result!.TextContent);
		Assert.Equal($"<p>{template.Template.TextContent}</p>", result.ResultItems[0].Result!.HtmlContent);
	}
	[Fact]
	public void TextContent_DoNotReplaceHtmlIfProvided()
	{
		//Given
		var template = new WvEmailTemplate
		{
			Template = new WvEmail
			{
				TextContent = "some text",
				HtmlContent = "<span>test</span>"
			}
		};
		var ds = EmailData;

		//When
		WvEmailTemplateProcessResult? result = template.Process(ds, DefaultCulture); ;

		//Then
		Assert.NotNull(result);
		Assert.Single(result.ResultItems);
		Assert.NotNull(result.ResultItems[0].Result);
		Assert.Equal(template.Template.TextContent, result.ResultItems[0].Result!.TextContent);
		Assert.Equal(template.Template.HtmlContent, result.ResultItems[0].Result!.HtmlContent);
	}
	#endregion

	#region <<Attachments>>
	[Fact]
	public void Attachments_EmptyShouldReturnEmpty()
	{
		lock (locker)
		{
			//Given
			var template = new WvEmailTemplate
			{
				Template = new WvEmail
				{
					AttachmentItems = new()
				}
			};
			var ds = EmailData;

			//When
			WvEmailTemplateProcessResult? result = template.Process(ds, DefaultCulture);

			//Then
			Assert.NotNull(result);
			Assert.Single(result.ResultItems);
			Assert.NotNull(result.ResultItems[0].Result);
			Assert.NotNull(result.ResultItems[0].Result!.AttachmentItems);
			Assert.Empty(result.ResultItems[0].Result!.AttachmentItems);
		}
	}

	[Fact]
	public void Attachments_Test()
	{
		lock (locker)
		{
			var fileName = "TemplateEmail1.xlsx";
			//Given
			var template = new WvEmailTemplate
			{
				Template = new WvEmail
				{
					AttachmentItems = new(){
						new WvEmailAttachment{
							Filename = fileName,
							Content = LoadWorkbookAsMemoryStream(fileName),
							Type =  WvEmailAttachmentType.ExcelFile
						}
					}
				}
			};
			var ds = EmailData;

			//When
			WvEmailTemplateProcessResult? result = template.Process(ds, DefaultCulture);

			//Then
			Assert.NotNull(result);
			Assert.Single(result.ResultItems);
			Assert.NotNull(result.ResultItems[0].Result);
			Assert.NotNull(result.ResultItems[0].Result!.AttachmentItems);
			Assert.Single(result.ResultItems[0].Result!.AttachmentItems);
			Assert.NotNull(result.ResultItems[0].Result!.AttachmentItems[0]!.Content);

			XLWorkbook? resultWorkbook = new XLWorkbook(result.ResultItems[0].Result!.AttachmentItems[0]!.Content!);
			var ws = resultWorkbook.Worksheets.First();
			Assert.Equal(EmailData.Rows[0]["sender_email"], ws.Cell(2, 1).Value.ToString());
			Assert.Equal(EmailData.Rows[1]["sender_email"], ws.Cell(3, 1).Value.ToString());
			Assert.Equal(EmailData.Rows[2]["sender_email"], ws.Cell(4, 1).Value.ToString());

			Assert.Equal(EmailData.Rows[0]["recipient_email"], ws.Cell(2, 2).Value.ToString());
			Assert.Equal(EmailData.Rows[1]["recipient_email"], ws.Cell(3, 2).Value.ToString());
			Assert.Equal(EmailData.Rows[2]["recipient_email"], ws.Cell(4, 2).Value.ToString());

			SaveWorkbookFromMemoryStream(result.ResultItems[0].Result!.AttachmentItems[0]!.Content!, fileName);
		}
	}

	[Fact]
	public void Attachments_GroupBy()
	{
		lock (locker)
		{
			var fileName = "TemplateEmail1Grouped.xlsx";
			//Given
			var template = new WvEmailTemplate
			{
				Template = new WvEmail
				{
					AttachmentItems = new(){
						new WvEmailAttachment{
							Filename = fileName,
							Content = LoadWorkbookAsMemoryStream(fileName),
							Type =  WvEmailAttachmentType.ExcelFile,
							GroupDataByColumns = new List<string>{"sender_email"}
						}
					}
				}
			};
			var ds = EmailData;
			ds.Rows[1]["sender_email"] = ds.Rows[0]["sender_email"];
			//When
			WvEmailTemplateProcessResult? result = template.Process(ds, DefaultCulture); ;

			//Then
			Assert.NotNull(result);
			Assert.Single(result.ResultItems);
			Assert.NotNull(result.ResultItems[0].Result);
			Assert.NotNull(result.ResultItems[0].Result!.AttachmentItems);
			Assert.Equal(4, result.ResultItems[0].Result!.AttachmentItems.Count);
			Assert.NotNull(result.ResultItems[0].Result!.AttachmentItems[0]!.Content);

			XLWorkbook? firstWorkbook = new XLWorkbook(result.ResultItems[0].Result!.AttachmentItems[0]!.Content!);
			XLWorkbook? secondWorkbook = new XLWorkbook(result.ResultItems[0].Result!.AttachmentItems[1]!.Content!);

			var firstWs = firstWorkbook.Worksheets.First();
			Assert.Equal(EmailData.Rows[0]["sender_email"], firstWs.Cell(2, 1).Value.ToString());
			Assert.Equal(EmailData.Rows[0]["sender_email"], firstWs.Cell(3, 1).Value.ToString());
			Assert.Equal(String.Empty, firstWs.Cell(4, 1).Value.ToString());

			Assert.Equal(EmailData.Rows[0]["recipient_email"], firstWs.Cell(2, 2).Value.ToString());
			Assert.Equal(EmailData.Rows[1]["recipient_email"], firstWs.Cell(3, 2).Value.ToString());
			Assert.Equal(String.Empty, firstWs.Cell(4, 2).Value.ToString());

			var secondWs = secondWorkbook.Worksheets.First();
			Assert.Equal(EmailData.Rows[2]["sender_email"], secondWs.Cell(2, 1).Value.ToString());
			Assert.Equal(String.Empty, secondWs.Cell(3, 1).Value.ToString());

			Assert.Equal(EmailData.Rows[2]["recipient_email"], secondWs.Cell(2, 2).Value.ToString());
			Assert.Equal(String.Empty, secondWs.Cell(3, 2).Value.ToString());

			SaveWorkbookFromMemoryStream(result.ResultItems[0].Result!.AttachmentItems[0]!.Content!, fileName);
		}
	}
	#endregion

	#region << From Docs >>
	[Fact]
	public void EmailDocs_test1()
	{
		//Creating the DataTable
		DataTable dt = new DataTable();
		dt.Columns.Add("email", typeof(string));
		dt.Columns.Add("name", typeof(string));
		dt.Columns.Add("item", typeof(string));
		//Row 1
		var row = dt.NewRow();
		row["email"] = $"john@domain.com";
		row["name"] = $"John Smith";
		row["item"] = $"Bike";
		dt.Rows.Add(row);
		//Row 2
		var row2 = dt.NewRow();
		row2["email"] = $"peter@domain.com";
		row2["name"] = $"Peter Jackson";
		row2["item"] = $"Car";
		dt.Rows.Add(row2);

		//Creating the template
		WvEmailTemplate template = new()
		{
			GroupDataByColumns = new List<string>() { "email" },
			Template = new WvEmail
			{
				Sender = "some@email.com",
				Recipients = "{{email}}",
				BccRecipients = null,
				CcRecipients = null,
				Subject = "About the {{item}}",
				AttachmentItems = new(), //no attachments needed
				HtmlContent = null, //will be autigenerated when null from the text
				TextContent = "Dear {{name}},\r\nThis email is about the {{item}}"

			}
		};

		//Execution
		WvEmailTemplateProcessResult result = template.Process(dt);
		Assert.Equal(2, result.ResultItems.Count);
		Assert.NotNull(result.ResultItems[0].Result);
		Assert.NotNull(result.ResultItems[1].Result);

		Assert.Equal("john@domain.com",result.ResultItems[0].Result!.Recipients);
		Assert.Equal("About the Bike",result.ResultItems[0].Result!.Subject);

		Assert.Equal("peter@domain.com",result.ResultItems[1].Result!.Recipients);
		Assert.Equal("About the Car",result.ResultItems[1].Result!.Subject);
	}
	#endregion
}

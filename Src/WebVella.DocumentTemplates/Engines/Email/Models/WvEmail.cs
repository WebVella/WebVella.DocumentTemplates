using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebVella.DocumentTemplates.Engines.Email.Models;
public class WvEmail
{
	public string? Sender { get; set; }
	public string? Recipients { get; set; }
	public string? CcRecipients { get; set; }
	public string? BccRecipients { get; set; }
	public string? Subject { get; set; }
	public string? HtmlContent { get; set; }
	public string? TextContent { get; set; }
	public List<string> GroupBy { get; set; } = new();
	public List<WvEmailAttachment> AttachmentItems { get; set; } = new();
}

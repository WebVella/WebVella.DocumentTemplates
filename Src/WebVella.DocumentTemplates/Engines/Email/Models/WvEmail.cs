namespace WebVella.DocumentTemplates.Engines.Email.Models;
public class WvEmail
{
	/// <summary>
	/// Single email or template resulting in a single email
	/// </summary>
	public string? Sender { get; set; }
	/// <summary>
	/// Email list or templated result. Emails should be separated by ";"
	/// </summary>
	public string? Recipients { get; set; }
	/// <summary>
	/// Empty, Email list or templated result. Emails should be separated by ";"
	/// </summary>
	public string? CcRecipients { get; set; }
	/// <summary>
	/// Empty, Email list or templated result. Emails should be separated by ";"
	/// </summary>
	public string? BccRecipients { get; set; }
	/// <summary>
	/// String or template
	/// </summary>
	public string? Subject { get; set; }
	/// <summary>
	/// Default email body. Html or template html. 
	/// </summary>
	public string? HtmlContent { get; set; }
	/// <summary>
	/// If empty, will be generated from the HtmlContent
	/// </summary>
	public string? TextContent { get; set; }
	public List<WvEmailAttachment> AttachmentItems { get; set; } = new();
}

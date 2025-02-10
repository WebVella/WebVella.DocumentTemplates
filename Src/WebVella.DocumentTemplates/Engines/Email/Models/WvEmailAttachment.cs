using System.Text;

namespace WebVella.DocumentTemplates.Engines.Email.Models;
public class WvEmailAttachment
{
	public Encoding? Encoding { get; set; } = null;
	public List<string> GroupDataByColumns { get; set; } = new List<string>();
	public WvEmailAttachmentType Type { get; set; } = WvEmailAttachmentType.TextFile;
	public string? Filename { get; set; }
	public byte[]? Content { get; set; }
}

public enum WvEmailAttachmentType{ 
	TextFile = 0,
	ExcelFile = 1
}
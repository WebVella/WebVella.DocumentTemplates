using DocumentFormat.OpenXml.Packaging;
using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.DocumentFile;
public class WvDocumentFileTemplateProcessResult : WvTemplateProcessResultBase
{
	public WordprocessingDocument? WordDocument { get; set; } = null;
	public new MemoryStream? Template { get; set; } = null;
	public new List<WvDocumentFileTemplateProcessResultItem> ResultItems { get; set; } = new();
}
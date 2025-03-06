using DocumentFormat.OpenXml.Packaging;
using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.DocumentFile;
public class WvDocumentFileTemplateProcessResultItem : WvTemplateProcessResultItemBase
{
	public new MemoryStream? Result { get; set; } = new();
	public WordprocessingDocument? WordDocument { get; set; } = null;
	public new List<WvDocumentFileTemplateProcessContext> Contexts { get; set; } = new();
}
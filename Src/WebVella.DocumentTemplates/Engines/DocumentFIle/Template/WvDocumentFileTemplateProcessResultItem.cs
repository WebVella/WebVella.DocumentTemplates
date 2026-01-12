using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.DocumentFile;

public class WvDocumentFileTemplateProcessResultItem : WvTemplateProcessResultItemBase
{
    public new MemoryStream? Result { get; set; } = new();
    public WordprocessingDocument? WordDocument { get; set; } = null;
    public new List<WvDocumentFileTemplateProcessContext> Contexts { get; set; } = new();

    public List<ValidationErrorInfo> Validate(FileFormatVersions version = FileFormatVersions.Microsoft365)
    {
        if (Result is null) return new List<ValidationErrorInfo>();
        using WordprocessingDocument wordDoc = WordprocessingDocument.Open(Result, false);
        OpenXmlValidator validator = new OpenXmlValidator(version);
        return validator.Validate(wordDoc).ToList();
    }
}
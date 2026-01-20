using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using WebVella.DocumentTemplates.Core;
namespace WebVella.DocumentTemplates.Engines.SpreadsheetFile;
public class WvSpreadsheetFileTemplateProcessResultItem : WvTemplateProcessResultItemBase
{
	public new MemoryStream? Result { get; set; } = new();
	public XLWorkbook? Workbook { get; set; } = new();
	public new List<WvSpreadsheetFileTemplateProcessContext> Contexts { get; set; } = new();
	public List<WvSpreadsheetFileTemplateProcessResultItemRow> ResultRows { get; set; } = new();
	public Dictionary<Guid, int> ContextProcessLog { get; set; } = new();
	
	public List<ValidationErrorInfo> Validate(FileFormatVersions version = FileFormatVersions.Microsoft365)
	{
		if (Result is null) return new List<ValidationErrorInfo>();
		using SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(Result, false);
		OpenXmlValidator validator = new OpenXmlValidator(version);
		return validator.Validate(excelDoc).ToList();
	}	
}


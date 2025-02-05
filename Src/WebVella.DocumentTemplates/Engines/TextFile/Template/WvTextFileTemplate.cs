using System.Data;
using System.Globalization;
using System.Text;
using WebVella.DocumentTemplates.Core;
using WebVella.DocumentTemplates.Core.Utility;
using WebVella.DocumentTemplates.Extensions;
namespace WebVella.DocumentTemplates.Engines.Text;
public class WvTextFileTemplate : WvTemplateBase
{
	public byte[]? Template { get; set; } = null;
	public WvTextFileTemplateResult? Process(DataTable dataSource, CultureInfo? culture = null, Encoding? encoding = null)
	{
		if (culture == null) culture = new CultureInfo("en-US");
		if (encoding == null) encoding = Encoding.UTF8;
		if (dataSource is null) throw new ArgumentException("No datasource provided!", nameof(dataSource));
		var result = new WvTextFileTemplateResult()
		{
			Template = Template
		};

		result.Result = Template;

		if (Template is null) return result;

		var fileStringContent = encoding.GetString(result.Template ?? new byte[0]);

		if (String.IsNullOrWhiteSpace(fileStringContent)) return result;

		var textTemplate = new WvTextTemplate
		{
			Template = fileStringContent.RemoveZeroBitSpaceCharacters()
		};

		WvTextTemplateResult? textTemplateResult = textTemplate.Process(dataSource, culture);

		if (textTemplateResult == null) return result;
		result.Result = null;

		if(!String.IsNullOrWhiteSpace(textTemplateResult.Result)){ 
			result.Result = encoding.GetBytes(textTemplateResult.Result ?? String.Empty);
		}
		return result;
	}
}
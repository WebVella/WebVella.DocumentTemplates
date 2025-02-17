using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;

namespace WebVella.DocumentTemplates.Core.Utility;
public partial class WvTemplateUtility
{
	public (string?, object?) ProcessTagInTemplate(string? templateResultString, object? templateResultObject,
		WvTemplateTag tag, DataTable dataSource, int? contextRowIndex, CultureInfo culture)
	{
		var currentCulture = Thread.CurrentThread.CurrentCulture;
		var currentUICulture = Thread.CurrentThread.CurrentUICulture;
		try
		{
			Thread.CurrentThread.CurrentCulture = culture;
			Thread.CurrentThread.CurrentCulture = culture;
			object? newResultObject = null;
			//Rules:
			//If tag not found or cannot be used return tag full string (no substitution)
			//if tag has index it is check for applicability if not applicable return tag full string for this template
			//if tag has no index - apply the submitted index
			//if not index is requested get the first if present		
			if (tag.Type == WvTemplateTagType.Data)
			{
				if (String.IsNullOrWhiteSpace(tag.Name)) return (templateResultString, templateResultObject);
				if (tag.Flow == WvTemplateTagDataFlow.Horizontal && contextRowIndex is null)
				{
					//If horizontal all items should be return in one template 
					(templateResultString, newResultObject) = ReplaceAllValuesDataTagInTemplate(templateResultString, templateResultObject, tag, dataSource);
				}
				else
				{
					(templateResultString, newResultObject) = ReplaceSingleValueDataTagInTemplate(templateResultString, templateResultObject, tag, dataSource, contextRowIndex);
				}
			}
			else if (tag.Type == WvTemplateTagType.Function)
			{
				newResultObject = templateResultString;//temporary
				//Processed by processor
			}
			else if (tag.Type == WvTemplateTagType.ExcelFunction)
			{
				newResultObject = templateResultString;//temporary
				//Processed by processor
			}
			return (templateResultString, newResultObject);
		}
		finally
		{
			Thread.CurrentThread.CurrentCulture = currentCulture;
			Thread.CurrentThread.CurrentUICulture = currentUICulture;
		}
	}

}

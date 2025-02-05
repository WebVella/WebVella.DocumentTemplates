using System.ComponentModel;
using System.Globalization;

namespace WebVella.DocumentTemplates.Extensions;

public static class StringExtensions
{
	public static string RemoveZeroBitSpaceCharacters(this string text)
	{
		if(String.IsNullOrWhiteSpace(text)) return text;
		return text.Replace("\uFEFF", "");;
	}

}

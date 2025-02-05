using System.ComponentModel;
using System.Globalization;

namespace WebVella.DocumentTemplates.Extensions;

public static class EnumExtensions
{
	public static string ToDescriptionString<TEnum>(this TEnum e) where TEnum : IConvertible
	{
		string description = "";

		if (e is Enum)
		{
			Type type = e.GetType();
			string? enumName = type.GetEnumName(e.ToInt32(CultureInfo.InvariantCulture));
			if(String.IsNullOrWhiteSpace(enumName))
				return description;

			var memInfo = type.GetMember(enumName);
			var soAttributes = memInfo[0].GetCustomAttributes(typeof(DescriptionAttribute), false);
			if (soAttributes.Length > 0)
			{
				// we're only getting the first description we find
				// others will be ignored
				description = ((DescriptionAttribute)soAttributes[0]).Description;
			}
		}

		return description;
	}

	public static TEnum ConvertIntToEnum<TEnum>(int value, TEnum defaultValue) where TEnum : IConvertible
	{
		if (Enum.IsDefined(typeof(TEnum), value))
		{
			return (TEnum)Enum.ToObject(typeof(TEnum), value);
		}
		return defaultValue;

	}

	public static TEnum ConvertStringToEnum<TEnum>(string value, TEnum defaultValue) where TEnum : IConvertible
	{
		if (String.IsNullOrEmpty(value)) return defaultValue;
		if (int.TryParse(value, out var n))
		{
			return ConvertIntToEnum(n, defaultValue);
		}
		return defaultValue;
	}

	public static TEnum2 ConvertSafeToEnum<TEnum, TEnum2>(this TEnum value) where TEnum : struct, IConvertible
																							 where TEnum2 : struct, IConvertible
	{
		if (Enum.TryParse<TEnum2>(value.ToString(), out TEnum2 result) && result.ToInt16(null) == value.ToInt16(null))
			return result;

		throw new Exception("Cannot be safely converted");

	}
}

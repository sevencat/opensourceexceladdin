namespace Sevencat.ExcelAddin.Common.Util;

public static class StringUtil
{
	public static bool IsNullOrEmpty(this string str)
	{
		return string.IsNullOrEmpty(str);
	}

	public static bool IsNullOrWhiteSpace(this string str)
	{
		return string.IsNullOrWhiteSpace(str);
	}

	public static int ToInt(this string str, int defaultValue)
	{
		if (str.IsNullOrEmpty())
			return defaultValue;
		return int.TryParse(str, out var result) ? result : defaultValue;
	}
}
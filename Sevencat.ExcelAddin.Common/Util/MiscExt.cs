using System;

namespace Sevencat.ExcelAddin.Common.Util;

public static class MiscExt
{
	public static void LogException(this NLog.Logger log, Exception ex)
	{
		log.Error("出错:{0},{1}", ex.Message, ex.StackTrace);
	}
}
using System;
using System.IO;

namespace Sevencat.ExcelAddin.Common.Util;

public class EnvUtil
{
	private static string _appdir;

	public static string AppDatDir
	{
		get
		{
			if (_appdir == null)
			{
				const string strProgramDataPath = "%PROGRAMDATA%";
				var directoryPath = Environment.ExpandEnvironmentVariables(strProgramDataPath);
				_appdir = Path.Combine(directoryPath, AppConstant.AppName);
			}

			return _appdir;
		}
		set => _appdir = value;
	}

	public static string AppDbDir => Path.Combine(AppDatDir, "app.db");

	private static string _workdir;

	//修改为location 
	public static string WorkDir
	{
		get
		{
			if (_workdir == null)
				_workdir = AppDomain.CurrentDomain.BaseDirectory;
			return _workdir;
		}
		set => _workdir = value;
	}
}
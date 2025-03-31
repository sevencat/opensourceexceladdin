using System.Drawing;
using System.IO;
using System.Text;
using Sevencat.ExcelAddin.Common.Service;
using Sevencat.ExcelAddin.Common.Util;

namespace Sevencat.ExcelAddin.Core.Service;

public class FileSystemResourceManager : IResourceManager
{
	private static readonly NLog.Logger Log = NLog.LogManager.GetCurrentClassLogger();
	private readonly string _basePath;

	public FileSystemResourceManager()
	{
		_basePath = Path.Combine(EnvUtil.WorkDir, "res");
		Log.Info("资源路径为:{0}", _basePath);
		Directory.CreateDirectory(_basePath);
	}

	public Image GetImage(string ImageName)
	{
		var fullfn = Path.Combine(_basePath, "img", ImageName);
		if (!File.Exists(fullfn))
		{
			Log.Error("资源文件不存在:{0}", fullfn);
			return null;
		}

		var bindat = File.ReadAllBytes(fullfn);
		using (var ms = new MemoryStream(bindat))
		{
			return Image.FromStream(ms);
		}
	}

	public string GetXml(string xmlfilename)
	{
		var fullfn = Path.Combine(_basePath, "xml", xmlfilename);
		if (!File.Exists(fullfn))
		{
			Log.Error("资源文件不存在:{0}", fullfn);
			return null;
		}

		var bindat = File.ReadAllText(fullfn, Encoding.UTF8);
		return bindat;
	}
}
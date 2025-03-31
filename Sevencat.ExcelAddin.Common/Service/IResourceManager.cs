using System.Drawing;

namespace Sevencat.ExcelAddin.Common.Service;

//外部的图片，资源，以及其他静态文件都从这里获取
public interface IResourceManager
{
	Image GetImage(string ImageName);
	string GetXml(string xmlfilename);
}
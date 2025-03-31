using System.Drawing;

namespace Sevencat.ExcelAddin.Common.Service;

public interface IResourceManager
{
	Image GetImage(string ImageName);
	string GetXml(string xmlfilename);
}
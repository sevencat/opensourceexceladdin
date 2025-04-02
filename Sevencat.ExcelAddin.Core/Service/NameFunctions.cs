using Autofac.Core;
using Sevencat.ExcelAddin.Core.Ui;

namespace Sevencat.ExcelAddin.Core.Service;

public class NameFunctions
{
	public static void NameSplit()
	{
		var win = new WinNameSplitInput();
		win.ShowDialog();
	}

	public static void NameAddBlank()
	{
		
	}

	public static void NameGenRandom()
	{
		
	}

	public static void NameAddStarMask()
	{
		
	}
	
	public static void Register(Container ct)
	{
	}
}
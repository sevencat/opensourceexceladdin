using Autofac;

namespace Sevencat.ExcelAddin.Common;

//这个一般是用在和界面绑定在一起的panel上的。其他一般不需要
public class ExcelWbContext
{
	public ILifetimeScope Scope { get; set; }

	public TService Resolve<TService>()
	{
		return Scope.Resolve<TService>();
	}

	public NetOffice.ExcelApi.Workbook Wb { get; set; }
	public object CustomTaskPane { get; set; }
}
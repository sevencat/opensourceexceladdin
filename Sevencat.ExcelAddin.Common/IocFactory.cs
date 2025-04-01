using Autofac;

namespace Sevencat.ExcelAddin.Common;

public class IocFactory
{
	public static ILifetimeScope ServiceProvider { get; set; }

	public static NetOffice.ExcelApi.Application Application { get; set; }

	public static NetOffice.ExcelApi.Application GetExcelApplication()
	{
		return Application;
	}

	public static T Get<T>()
	{
		return ServiceProvider.Resolve<T>();
	}

	public static T GetByName<T>(string name)
	{
		return ServiceProvider.ResolveNamed<T>(name);
	}

	public static T GetOptional<T>() where T : class
	{
		return ServiceProvider.ResolveOptional<T>();
	}

	public static T GetByNameOptional<T>(string name) where T : class
	{
		return ServiceProvider.ResolveOptionalNamed<T>(name);
	}
}
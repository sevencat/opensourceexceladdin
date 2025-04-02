using System;
using System.Net;
using Autofac;
using NetOffice.OfficeApi;
using Sevencat.ExcelAddin.Common;
using Sevencat.ExcelAddin.Common.Service;
using Sevencat.ExcelAddin.Core.Service;

namespace Sevencat.ExcelAddin.Core;

public class CoreModule : Module
{
	protected override void Load(ContainerBuilder builder)
	{
		ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

		builder.RegisterType<NameFuncService>().AsSelf().SingleInstance();

		builder.RegisterType<FileSystemResourceManager>().AsSelf().As<IResourceManager>().SingleInstance();
		//RegisterNamedUserControl<WpfMainFrame>(builder, AppConstant.Uc_MainFrame);

		RegisterExcelFunc(builder, "NameSplit", NameFunctions.NameSplit);
		RegisterExcelFunc(builder, "NameAddBlank", NameFunctions.NameAddBlank);
		RegisterExcelFunc(builder, "NameGenRandom", NameFunctions.NameGenRandom);
		RegisterExcelFunc(builder, "NameAddStarMask", NameFunctions.NameAddStarMask);
	}

	private static void RegisterExcelFunc(ContainerBuilder builder, string name, Action act)
	{
		Action<IRibbonControl> ribbonControl = _ => act();
		builder.RegisterInstance(ribbonControl).Named<Action<IRibbonControl>>(name.ToLower()).SingleInstance();
	}

	private static void RegisterNamedUserControl<T>(ContainerBuilder builder, string name)
		where T : System.Windows.Controls.UserControl
	{
		builder.RegisterType<T>().AsSelf().Named<System.Windows.Controls.UserControl>(name)
			.InstancePerLifetimeScope();
	}
}
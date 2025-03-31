using System.Net;
using Autofac;
using Sevencat.ExcelAddin.Common;
using Sevencat.ExcelAddin.Common.Service;
using Sevencat.ExcelAddin.Core.Service;

namespace Sevencat.ExcelAddin.Core;

public class CoreModule : Module
{
	protected override void Load(ContainerBuilder builder)
	{
		ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

		builder.RegisterType<FileSystemResourceManager>().AsSelf().As<IResourceManager>().SingleInstance();
		//RegisterNamedUserControl<WpfMainFrame>(builder, AppConstant.Uc_MainFrame);
	}

	private static void RegisterNamedUserControl<T>(ContainerBuilder builder, string name)
		where T : System.Windows.Controls.UserControl
	{
		builder.RegisterType<T>().AsSelf().Named<System.Windows.Controls.UserControl>(name)
			.InstancePerLifetimeScope();
	}
}
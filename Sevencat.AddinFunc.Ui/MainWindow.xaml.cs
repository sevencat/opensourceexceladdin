using System;
using System.Windows;
using Nancy.Hosting.Self;

namespace Sevencat.AddinFunc.Ui
{
	/// <summary>
	/// MainWindow.xaml 的交互逻辑
	/// </summary>
	public partial class MainWindow : Window
	{
		public MainWindow()
		{
			InitializeComponent();

			var hostConfigs = new HostConfiguration
			{
				UrlReservations = new UrlReservations() { CreateAutomatically = true },
				MaximumConnectionCount = 2,
			};
			var uri = new Uri("http://localhost:9800");
			_host = new NancyHost(hostConfigs, uri);

			_host.Start();
		}

		private NancyHost _host;
	}
}
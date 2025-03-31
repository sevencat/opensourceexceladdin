using NetOffice.ExcelApi;
using NetOffice.ExcelApi.Tools;
using NetOffice.Tools;
using System;
using System.Data.SQLite;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Autofac;
using NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using NLog.Targets;
using Sevencat.ExcelAddin.Common;
using Sevencat.ExcelAddin.Common.Service;
using Sevencat.ExcelAddin.Common.Util;
using Sevencat.ExcelAddin.Core;

namespace Sevencat.ExcelComAddin
{
	[ComVisible(true)]
	[Guid("ac5ab611-5a1a-4d6e-b94f-3afb90e46395")]
	[ProgId("Sevencat.ExcelComAddin.MyAddin")]
	[COMAddin("SevencatExcelAddin", "Sevencat's Excel Addin", LoadBehavior.LoadAtStartup)]
	public class MyAddin : COMAddin
	{
		private static NLog.Logger Log;
		private static IFreeSql Db { get; set; }

		//这个东西比较奇怪，在vsto里面，这个东西得一开始就初始化，
		//因为ribbon的回调非常非常早，在startup以前，这个设计好奇怪了。
		private static IResourceManager _resourceManager;

		//启动的时候只初始化数据库和日志
		public MyAddin()
		{
			AppConstant.AppName = "SevencatExcelAddin";

			InitLog();
			Log = NLog.LogManager.GetCurrentClassLogger();

			var codebase = typeof(MyAddin).Assembly.CodeBase;
			var dllfn = new Uri(codebase).LocalPath;
			var workdir = Path.GetDirectoryName(dllfn);
			EnvUtil.WorkDir = workdir;
			Log.Info("工作路径为:{0}", workdir);

			Log.Info("开始初始化数据库");
			var dbfn = EnvUtil.AppDbDir;
			InitDb(dbfn);
			Log.Info("结束初始化数据库");


			this.OnConnection += MyAddin_OnConnection;
		}

		#region 初始化

		public static IFreeSql InitDb(string dbfn)
		{
			if (!File.Exists(dbfn))
				SQLiteConnection.CreateFile(dbfn);
			var blder = new SQLiteConnectionStringBuilder();
			blder.Pooling = true;
			blder.JournalMode = SQLiteJournalModeEnum.Wal;
			blder.DataSource = dbfn;
			var connstr = blder.ToString();
			Log.Info("数据库开始初始化:{0}", connstr);

			Db = new FreeSql.FreeSqlBuilder()
				.UseConnectionString(FreeSql.DataType.Sqlite, connstr)
				.UseAutoSyncStructure(false)
				.UseLazyLoading(true)
				.Build();

			Log.Info("数据库初始化成功:{0}", dbfn);
			return Db;
		}

		public static void InitLog()
		{
			var config = new NLog.Config.LoggingConfiguration();

			var logdir = Path.Combine(EnvUtil.AppDatDir, "logs");
			Directory.CreateDirectory(logdir);
			var logfn = Path.Combine(logdir, "excel_log.txt");
			var fileTarget = new FileTarget("logfile") { FileName = logfn };
			fileTarget.ArchiveEvery = FileArchivePeriod.Day;
			fileTarget.ArchiveFileName = logdir + "\\" + "excel_log_{########}.txt";
			fileTarget.ArchiveNumbering = ArchiveNumberingMode.Date;
			fileTarget.ArchiveDateFormat = "yyyyMMdd";
			fileTarget.ArchiveOldFileOnStartup = true;
			fileTarget.MaxArchiveDays = 64;

			var asyncFileTarget = new NLog.Targets.Wrappers.AsyncTargetWrapper(fileTarget)
			{
				Name = fileTarget.Name,
				QueueLimit = 128,
				OverflowAction = NLog.Targets.Wrappers.AsyncTargetWrapperOverflowAction.Discard
			};

			config.AddRule(NLog.LogLevel.Debug, NLog.LogLevel.Fatal, asyncFileTarget);
			NLog.LogManager.Configuration = config;
		}

		#endregion

		private void MyAddin_OnConnection(object application, ext_ConnectMode connectMode, object addInInst,
			ref Array custom)
		{
			var builder = new ContainerBuilder();
			builder.RegisterInstance(Application);
			builder.RegisterInstance(Db);
			builder.RegisterModule<CoreModule>();
			builder.RegisterModule<CommonModule>();
			IocFactory.ServiceProvider = builder.Build();
			_resourceManager = IocFactory.Get<IResourceManager>();
			this.Application.WorkbookOpenEvent += Application_WorkbookOpenEvent;
		}

		private void Application_WorkbookOpenEvent(Workbook workbook)
		{
			using (workbook)
			{
				// start working with the workbook
			}
		}

		#region 回调

		public override string GetCustomUI(string RibbonID)
		{
			return _resourceManager.GetXml("excelribbonui.xml");
		}

		public void Ribbon_Load(IRibbonUI ribbonUI)
		{
			Log.Info("Ribbon_Load");
		}

		public void RibbonBtnClick(IRibbonControl control)
		{
			var id = control.Id.ToLower();
			var func = IocFactory.GetByNameOptional<Action<IRibbonControl>>(id);
			if (func == null)
			{
				Log.Error("没有{0}的处理函数", id);
				MessageBox.Show("没有" + id + "的处理函数", "出错");
				return;
			}

			try
			{
				func(control);
			}
			catch (Exception ex)
			{
				Log.LogException(ex);
				MessageBox.Show("执行异常:" + ex.Message, "异常");
			}
		}

		public void CommonWordFunc_Click(IRibbonControl control)
		{
			var id = control.Id;
			var tag = control.Tag;
			if (id == "aboutButton")
			{
				var panel = this.TaskPaneFactory.CreateCTP("Bastet.OfficeAddin.AddinPanel", "Example");
				panel.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
				panel.Width = 550;
				panel.Visible = true;
			}
		}

		#endregion
	}
}
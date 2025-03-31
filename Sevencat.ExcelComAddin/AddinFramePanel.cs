using System.Runtime.InteropServices;
using System.Windows.Forms;
using NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;

namespace Sevencat.ExcelComAddin;

[ProgId("Sevencat.ExcelComAddin.AddinFramePanel")]
[ClassInterface(ClassInterfaceType.AutoDispatch)]
public partial class AddinFramePanel : UserControl, NetOffice.ExcelApi.Tools.ITaskPane
{
	public AddinFramePanel()
	{
		InitializeComponent();
	}

	#region excel相关

	public void OnConnection(NetOffice.ExcelApi.Application application, _CustomTaskPane parentPane,
		object[] customArguments)
	{
	}

	void NetOffice.ExcelApi.Tools.ITaskPane.OnDockPositionChanged(MsoCTPDockPosition position)
	{
	}

	void NetOffice.ExcelApi.Tools.ITaskPane.OnVisibleStateChanged(bool visible)
	{
	}

	void NetOffice.ExcelApi.Tools.ITaskPane.OnDisconnection()
	{
	}

	#endregion
}
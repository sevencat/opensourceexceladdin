using NetOffice.ExcelApi;
using NetOffice.ExcelApi.Tools;
using NetOffice.Tools;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace Sevencat.ExcelComAddin
{
    [ComVisible(true)]
    [Guid("ac5ab611-5a1a-4d6e-b94f-3afb90e46395")]
    [ProgId("Sevencat.ExcelComAddin.MyAddin")]
    [COMAddin("MyAddin", "Addin description.", LoadBehavior.LoadAtStartup)]
    public class MyAddin : COMAddin
    {
        public MyAddin()
        {
            this.OnConnection += MyAddin_OnConnection;
        }

        private void MyAddin_OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            this.Application.WorkbookOpenEvent += Application_WorkbookOpenEvent;
        }

        private void Application_WorkbookOpenEvent(Workbook workbook)
        {
            using (workbook)
            {
                // start working with the workbook
            }
        }
    }
}

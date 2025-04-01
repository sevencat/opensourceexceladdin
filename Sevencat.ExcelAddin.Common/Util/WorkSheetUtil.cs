using NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;

namespace Sevencat.ExcelAddin.Common.Util;

public static class WorkSheetUtil
{
	//好多种方法,这个可能是最快的方法
	//https://stackoverflow.com/questions/7674573/programmatically-getting-the-last-filled-excel-row-using-c-sharp
	public static int GetLastRow(this Worksheet ws)
	{
		var lastcell = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);
		return lastcell.Row;
	}
}
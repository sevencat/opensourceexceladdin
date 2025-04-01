using System;
using System.Collections.Generic;
using Sevencat.ExcelAddin.Common.Model;

namespace Sevencat.ExcelAddin.Common.Util;

public static class RangeUtil
{
	public static bool IsOneColumn(this string range)
	{
		if (range.IsNullOrWhiteSpace())
			return true;
		var areas = range.ToAreaItems();
		if (areas.Count <= 0)
			return true;
		var firstarea = areas[0];
		if (!firstarea.IsSameCol())
			return false;
		var col = firstarea.From.Col;
		for (var i = 1; i < areas.Count; i++)
		{
			if (!areas[i].IsAllCol(col))
				return false;
		}

		return true;
	}

	public static List<ExcelAreaItem> ToAreaItems(this string range)
	{
		if (range.IsNullOrWhiteSpace())
			return [];
		var retlist = new List<ExcelAreaItem>();
		var subareas = range.Split(',');
		foreach (var area in subareas)
		{
			var curarea = area.Trim();
			retlist.Add(curarea.ConvertToArea());
		}

		return retlist;
	}

	public static ExcelAreaItem ConvertToArea(this string rangeitem)
	{
		//$D$3:$L$3
		//$D$8
		var sepPos = rangeitem.IndexOf(':');
		if (sepPos < 0)
		{
			var cell = ConvertToCell(rangeitem);
			return new ExcelAreaItem()
			{
				From = cell,
				To = cell,
			};
		}

		var frompart = rangeitem.Substring(0, sepPos);
		var fromcell = frompart.ConvertToCell();
		var topart = rangeitem.Substring(sepPos + 1);
		var tocell = topart.ConvertToCell();
		return new ExcelAreaItem()
		{
			From = fromcell,
			To = tocell,
		};
	}

	//$F:$F
	public static ExcelCellItem ConvertToCell(this string rangeitem)
	{
		//$D$8
		if (!rangeitem.StartsWith("$"))
		{
			throw new Exception("选择位置错误" + rangeitem);
		}

		var rowpos = rangeitem.IndexOf('$', 1);
		if (rowpos < 0)
		{
			var scol = rangeitem.Substring(1);
			return new ExcelCellItem()
			{
				Col = scol,
				Row = -1
			};
		}
		else
		{
			var scol = rangeitem.Substring(1, rowpos - 1);
			var srow = rangeitem.Substring(rowpos + 1);
			return new ExcelCellItem()
			{
				Col = scol,
				Row = int.Parse(srow)
			};
		}
	}
}

// 只有一个单元格：没问题 $D$8 
// 横向的 有问题 $D$3:$L$3
// 纵向的，确实是同一列 $D$4:$D$14
// 两个区域这个不行 $B$5:$B$13,$D$5:$F$5
// 多个区域 这个全部在C列，所以是可以的 $C$3:$C$7,$C$10:$C$12,$C$16:$C$19,$C$24:$C$27
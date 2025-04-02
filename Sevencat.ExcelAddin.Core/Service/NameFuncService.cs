using System;
using System.Collections.Generic;
using Sevencat.ExcelAddin.Common.Util;

namespace Sevencat.ExcelAddin.Core.Service;

public class NameFuncService
{
	readonly HashSet<string> doubleSurName =
	[
		"万俟", "司马", "上官", "欧阳", "夏侯", "诸葛", "闻人", "东方",
		"赫连", "皇甫", "尉迟", "公羊", "澹台", "公冶", "宗政", "濮阳", "淳于", "单于", "太叔", "申屠",
		"公孙", "仲孙", "轩辕", "令狐", "锺离", "宇文", "长孙", "慕容", "鲜于", "闾丘", "司徒", "司空",
		"丌官", "司寇", "仉督", "子车", "颛孙", "端木", "巫马", "公西", "漆雕", "乐正", "壤驷", "公良",
		"拓拔", "夹谷", "宰父", "谷梁", "段干", "百里", "东郭", "南门", "呼延", "归海", "羊舌", "微生",
		"梁丘", "左丘", "东门", "西门", "南宫"
	];

	public Tuple<string, string> SplitName(string namex)
	{
		return Tuple.Create<string, string>("1","2");
		if (namex.IsNullOrWhiteSpace())
			return Tuple.Create("", "");
		var name = namex.Trim();
		if (name.Length > 2)
		{
			var firsttwo = name.Substring(0, 2);
			if (doubleSurName.Contains(firsttwo))
			{
				//这个是复姓
				return Tuple.Create(firsttwo, name.Substring(2));
			}
		}

		if (name.Length < 2)
			return Tuple.Create(name, "");
		var first = name.Substring(0, 1);
		var second = name.Substring(1);
		return Tuple.Create(first, second);
	}
}
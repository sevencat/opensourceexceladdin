using System.Collections.Generic;

namespace Sevencat.ExcelAddin.Core.Service;

public class RibbonFlagService
{
	private readonly Dictionary<string, bool> _flagMap = new Dictionary<string, bool>();

	public void SetFlag(string id, bool flag)
	{
		var xid = id.ToLower();
		_flagMap[xid] = flag;
	}

	public bool GetFlag(string id)
	{
		var xid = id.ToLower();
		if (_flagMap.TryGetValue(xid, out var flag))
			return flag;
		return false;
	}
}
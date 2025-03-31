using System;
using System.Collections.Generic;
using System.Linq;
using Autofac;
using NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;

namespace Sevencat.ExcelAddin.Common;

public class SubScopeManager
{
	private static readonly NLog.Logger Log = NLog.LogManager.GetCurrentClassLogger();

	private static readonly string UidPropKey = AppConstant.AppName + "_uid";

	public static string GetOrCreateUuid(object props)
	{
		var properties = (DocumentProperties)props;
		foreach (var prop in properties.Where(prop => prop.Name == UidPropKey))
		{
			return (string)prop.Value;
		}

		var uid = Guid.NewGuid().ToString();
		properties.Add(UidPropKey, false, MsoDocProperties.msoPropertyTypeString, uid);
		return uid;
	}

	private static readonly Dictionary<string, ExcelWbContext> _excelScopes = new();

	public static ExcelWbContext GetExcelWbContext2(NetOffice.ExcelApi.Workbook wb)
	{
		lock (_excelScopes)
		{
			foreach (var curctx in _excelScopes.Values)
			{
				if (curctx.Wb.UnderlyingObject == wb.UnderlyingObject)
				{
					return curctx;
				}
			}

			return null;
		}
	}

	public static ExcelWbContext GetExcelWbContext(NetOffice.ExcelApi.Workbook wb)
	{
		string docid = null;
		try
		{
			docid = GetOrCreateUuid(wb.CustomDocumentProperties);
		}
		catch (Exception ex)
		{
			//这里是拿不到，只有用笨办法
			return GetExcelWbContext2(wb);
		}

		Log.Info("workbook id is {0}", docid);
		lock (_excelScopes)
		{
			if (_excelScopes.TryGetValue(docid, out var ctx))
				return ctx;
			var scope = IocFactory.ServiceProvider.BeginLifetimeScope();
			ctx = scope.Resolve<ExcelWbContext>();
			ctx.Scope = scope;
			ctx.Wb = wb;
			_excelScopes[docid] = ctx;
			return ctx;
		}
	}

	public static void Remove(NetOffice.ExcelApi.Workbook wb)
	{
		var docid = GetOrCreateUuid(wb.CustomDocumentProperties);
		lock (_excelScopes)
		{
			_excelScopes.Remove(docid);
		}
	}
}
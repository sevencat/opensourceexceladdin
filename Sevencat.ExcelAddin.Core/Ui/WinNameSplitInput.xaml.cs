using System;
using System.Collections.Generic;
using System.Windows;
using NetOffice.ExcelApi;
using Sevencat.ExcelAddin.Common;
using Sevencat.ExcelAddin.Common.Util;
using Sevencat.ExcelAddin.Core.Service;
using Window = System.Windows.Window;

namespace Sevencat.ExcelAddin.Core.Ui
{
	/// <summary>
	/// WinNameSplitInput.xaml 的交互逻辑
	/// </summary>
	public partial class WinNameSplitInput : Window
	{
		private Range _selrange;

		readonly NameFuncService _nameFuncService;

		public WinNameSplitInput()
		{
			_nameFuncService = IocFactory.Get<NameFuncService>();
			InitializeComponent();
		}

		private void BtnSelectRange_OnClick(object sender, RoutedEventArgs e)
		{
			var app = IocFactory.Application;
			var defaultRange = this.TextBoxExcelRange.Text;
			var ret = app.InputBox("请选择单列", "请选择姓名所在列", defaultRange, Type.Missing, Type.Missing, Type.Missing,
				Type.Missing, 8);
			if (ret is bool)
			{
				return;
			}

			if (ret is Range range)
			{
				var srange = range.Address;
				if (!srange.IsOneColumn())
				{
					MessageBox.Show("区域不在同一列", "错误");
					return;
				}

				_selrange = range;
				this.TextBoxExcelRange.Text = srange;
			}
		}

		private void Handle()
		{
			var srange = this.TextBoxExcelRange.Text;
			var areas = srange.ToAreaItems();
			if (areas.Count == 0)
				return;
			//准备开始合并
			var col = areas[0].From.Col;

			var lastrow = _selrange.Worksheet.GetLastRow();
			var rows = new SortedSet<int>();
			foreach (var curarea in areas)
			{
				var fromrow = curarea.From.Row;
				if (fromrow <= 0)
					fromrow = 1;
				var torow = curarea.To.Row;
				if (torow < 0)
					torow = lastrow;
				for (var i = fromrow; i <= torow; i++)
					rows.Add(i);
			}

			Handle(_selrange.Worksheet, col, rows);
		}

		private void Handle(Worksheet ws, string col, SortedSet<int> rows)
		{
			var app = ws.Application;
			app.ScreenUpdating = false;
			try
			{
				foreach (var row in rows)
				{
					using (var cell = ws.Cells[row, col])
					{
						var orgvalue = cell.Value2?.ToString();
						if (orgvalue.IsNullOrWhiteSpace())
							continue;
						var namepair = _nameFuncService.SplitName(orgvalue);
						if (namepair.Item1.IsNullOrWhiteSpace())
							continue;
						using (var nextcell = cell.Offset(0, 1))
						{
							cell.Value = namepair.Item1;
							nextcell.Value = namepair.Item2;
						}
					}
				}
			}
			finally
			{
				app.ScreenUpdating = true;
			}
		}


		private void BtnOk_OnClick(object sender, RoutedEventArgs e)
		{
			Handle();
			e.Handled = true;
			this.Close();
		}
	}
}
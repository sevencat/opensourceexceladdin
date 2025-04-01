namespace Sevencat.ExcelAddin.Common.Model;

public class ExcelAreaItem
{
	public ExcelCellItem From { get; set; }
	public ExcelCellItem To { get; set; }

	public bool IsSameCol()
	{
		return From.Col == To.Col;
	}

	public bool IsAllCol(string col)
	{
		return (From.Col == col && To.Col == col);
	}
}
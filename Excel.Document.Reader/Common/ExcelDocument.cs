namespace Excel.Document.Reader.Common
{
	public class ExcelDocument
	{
		public string Name { get; set; }

		public Table[] Tables { get; set; }
	}

	public class Table
	{
		public string Name { get; set; }

		public Cell[,] Cells { get; set; }

		public int[] ColumnWidth { get; set; }
	}

	public class Cell
	{
		public string Value { get; set; }
	}
}
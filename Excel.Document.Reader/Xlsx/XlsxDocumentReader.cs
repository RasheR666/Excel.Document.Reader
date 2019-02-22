using System;
using System.Collections.Generic;
using System.IO;
using Excel.Document.Reader.Common;
using ExcelDataReader;

namespace Excel.Document.Reader.Xlsx
{
	public class XlsxDocumentReader
	{
		public ExcelDocument Read(string xlsxFilepath)
		{
			using(var fileStream = File.Open(xlsxFilepath, FileMode.Open, FileAccess.Read))
			{
				using(var reader = ExcelReaderFactory.CreateOpenXmlReader(fileStream))
				{
					return ReadXlsxDocument(Path.GetFileName(xlsxFilepath), reader);
				}
			}
		}

		private ExcelDocument ReadXlsxDocument(string xlsxFilename, IExcelDataReader reader)
		{
			var tables = new List<Table>();
			do
			{
				tables.Add(ReadTable(reader));
			} while(reader.NextResult()); // get next sheet

			return new ExcelDocument
			{
				Name = xlsxFilename,
				Tables = tables.ToArray()
			};
		}

		private Table ReadTable(IExcelDataReader reader)
		{
			var table = new Table
			{
				Name = reader.Name,
				Cells = new Cell[reader.RowCount, reader.FieldCount],
				ColumnWidth = new int[reader.FieldCount]
			};

			for(var rowIndex = 0; reader.Read(); rowIndex++)
			{
				for(var columnIndex = 0; columnIndex < reader.FieldCount; columnIndex++)
				{
					var cellValue = reader.GetValue(columnIndex)?.ToString() ?? string.Empty;
					table.Cells[rowIndex, columnIndex] = new Cell {Value = cellValue};
					table.ColumnWidth[columnIndex] = Math.Max(table.ColumnWidth[columnIndex], cellValue.Length);
				}
			}

			return table;
		}
	}
}
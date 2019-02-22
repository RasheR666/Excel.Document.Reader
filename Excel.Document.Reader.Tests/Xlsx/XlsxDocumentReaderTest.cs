using System.Collections.Generic;
using System.IO;
using System.Text;
using ApprovalTests;
using ApprovalTests.Namers;
using ApprovalTests.Reporters;
using Excel.Document.Reader.Common;
using Excel.Document.Reader.Xlsx;
using NUnit.Framework;
using Testing.Tools.Directory;

namespace Excel.Document.Reader.Tests.Xlsx
{
	[TestFixture]
	[UseReporter(typeof(DiffReporter))]
	[UseApprovalSubdirectory(@"TestData\ApprovalFiles")]
	public class XlsxDocumentReaderTest
	{
		public XlsxDocumentReaderTest()
		{
			reader = new XlsxDocumentReader();
		}

		[Test]
		[TestCaseSource(nameof(XlsxTestCases))]
		public void Read(FileInfo file)
		{
			using(ApprovalResults.ForScenario(file.Name))
			{
				var xlsx = reader.Read(file.FullName);
				Approvals.Verify(ToString(xlsx));
			}
		}

		private string ToString(ExcelDocument document)
		{
			var sb = new StringBuilder();

			sb.AppendLine($"Document name: {document.Name}");
			foreach(var table in document.Tables)
			{
				sb.AppendLine($"Table name: {table.Name}");
				sb.AppendLine("Cells:");
				for(var rowIndex = 0; rowIndex < table.Cells.GetLength(0); rowIndex++)
				{
					for(var columnIndex = 0; columnIndex < table.Cells.GetLength(1); columnIndex++)
					{
						sb.Append($"{table.Cells[rowIndex, columnIndex].Value} ");
					}

					sb.AppendLine();
				}

				sb.AppendLine();
				sb.AppendLine("ColumnWidths:");
				foreach(var width in table.ColumnWidth)
				{
					sb.Append($"{width} ");
				}
			}

			return sb.ToString();
		}

		private static IEnumerable<TestCaseData> XlsxTestCases
		{
			get
			{
				var xlsxFilesDirectory = TestDataDirectory.GetOrCreateSubDirectory("InputData");
				foreach(var xlsxFile in xlsxFilesDirectory.GetFiles("*.xlsx"))
				{
					yield return new TestCaseData(xlsxFile).SetName(xlsxFile.Name);
				}
			}
		}

		private static DirectoryInfo TestDataDirectory
		{
			get
			{
				return TestingProjectDirectoryProvider.ExcelDocumentReaderTestsDirectory
					.GetOrCreateSubDirectory("Xlsx")
					.GetOrCreateSubDirectory("TestData");
			}
		}

		private readonly XlsxDocumentReader reader;
	}
}
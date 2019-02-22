using System.IO;
using Testing.Tools.Directory;

namespace Excel.Document.Reader.Tests
{
	public static class TestingProjectDirectoryProvider
	{
		public static DirectoryInfo SolutionDirectory { get; } = SolutionDirectoryProvider.Get("Excel.Document.Reader.sln");
		public static DirectoryInfo ExcelDocumentReaderTestsDirectory { get; } = SolutionDirectory.GetOrCreateSubDirectory("Excel.Document.Reader.Tests");
	}
}
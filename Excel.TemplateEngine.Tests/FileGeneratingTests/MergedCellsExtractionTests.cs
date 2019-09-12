using System.IO;
using System.Linq;

using FluentAssertions;

using NUnit.Framework;

using SkbKontur.Excel.TemplateEngine.FileGenerating;

using Vostok.Logging.Console;

namespace SkbKontur.Excel.TemplateEngine.Tests.FileGeneratingTests
{
    [TestFixture]
    public class MergedCellsExtractionTests : FileBasedTestBase
    {
        [Test]
        public void MergedCellsExtractionTest()
        {
            var document = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("template.xlsx")), logger);
            var worksheet = document.GetWorksheet(0);

            var cells = worksheet.MergedCells.ToArray();

            cells.Length.Should().Be(2);
            cells[0].Item1.CellReference.Should().Be("B32");
            cells[0].Item2.CellReference.Should().Be("C35");
            cells[1].Item1.CellReference.Should().Be("I2");
            cells[1].Item2.CellReference.Should().Be("J3");

            document.Dispose();
        }

        [Test]
        public void MergedCellsExtractionEmptyTest()
        {
            var document = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("empty.xlsx")), logger);
            var worksheet = document.GetWorksheet(0);

            var cells = worksheet.MergedCells.ToArray();

            cells.Length.Should().Be(0);

            document.Dispose();
        }

        private readonly ConsoleLog logger = new ConsoleLog();
    }
}
using System.IO;
using System.Linq;

using FluentAssertions;

using NUnit.Framework;

using SkbKontur.Excel.TemplateEngine.FileGenerating;
using SkbKontur.Excel.TemplateEngine.FileGenerating.DataTypes;

using Vostok.Logging.Console;

namespace SkbKontur.Excel.TemplateEngine.Tests.FileGeneratingTests
{
    public class DocumentPartsExtractionTests : FileBasedTestBase
    {
        [Test]
        public void CellsInRangeTest()
        {
            var document = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("template.xlsx")), logger);
            var worksheet = document.GetWorksheet(0);

            var upperLeft = new ExcelCellIndex("A22");
            var lowerRight = new ExcelCellIndex("A24");
            var cells = worksheet.GetSortedCellsInRange(upperLeft, lowerRight).ToArray();

            cells.Length.Should().Be(3);
            cells[1].GetStringValue().Should().Be("Value:String:Name");
            cells[2].GetStringValue().Should().Be("Value:String:Address");

            document.Dispose();
        }

        [Test]
        public void GetCellTest()
        {
            var document = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("template.xlsx")), logger);
            var worksheet = document.GetWorksheet(0);

            var position = new ExcelCellIndex("B9");
            var cell = worksheet.GetCell(position);
            cell.GetStringValue().Should().Be("Model:Document:B10:C11");

            document.Dispose();
        }

        [Test]
        public void SearchCellsByTextTest()
        {
            var document = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("template.xlsx")), logger);
            var worksheet = document.GetWorksheet(0);

            var cell = worksheet.SearchCellsByText("Value:String").FirstOrDefault();
            cell.Should().NotBeNull();
// ReSharper disable PossibleNullReferenceException
            cell.GetStringValue().Should().Be("Value:String:Name");
// ReSharper restore PossibleNullReferenceException

            document.Dispose();
        }

        private readonly ConsoleLog logger = new ConsoleLog();
    }
}
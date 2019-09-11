using System.IO;
using System.Linq;

using Excel.TemplateEngine.FileGenerating;
using Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitivesImplementation;
using Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;

using NUnit.Framework;

using Vostok.Logging.Console;

namespace Excel.TemplateEngine.Tests.ObjectPrintingTests
{
    [TestFixture]
    public class ExcelPrimitivesWrappersTests : FileBasedTestBase
    {
        [Test]
        public void ExcelTablePartExtractionTest()
        {
            var document = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("template.xlsx")), logger);
            var table = new ExcelTable(document.GetWorksheet(0));

            var rows = table.GetTablePart(new Rectangle(new CellPosition("B9"), new CellPosition("D11")))
                            .Cells
                            .Select(row => row.ToArray())
                            .ToArray();

            foreach (var cell in rows.SelectMany(row => row))
                Assert.AreNotEqual(null, cell);

            Assert.AreEqual("Template:RootTemplate:B10:D11", rows[0][0].StringValue);
            Assert.AreEqual("", rows[0][1].StringValue + "");
            Assert.AreEqual("", rows[0][2].StringValue + "");
            Assert.AreEqual("Покупатель:", rows[1][0].StringValue);
            Assert.AreEqual("Поставщик:", rows[1][1].StringValue);
            Assert.AreEqual("", rows[1][2].StringValue + "");
            Assert.AreEqual("Value:Organization:Buyer", rows[2][0].StringValue);
            Assert.AreEqual("Value:Organization:Supplier", rows[2][1].StringValue);
            Assert.AreEqual("Value::TypeName", rows[2][2].StringValue);
        }

        [Test]
        public void ExcelCellExtractionTest()
        {
            var document = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("template.xlsx")), logger);
            var table = new ExcelTable(document.GetWorksheet(0));

            var cell = table.GetCell(new CellPosition("B9"));
            Assert.AreNotEqual(null, cell);
            Assert.AreEqual("Template:RootTemplate:B10:D11", cell.StringValue);

            cell = table.GetCell(new CellPosition("ABCD4234"));
            Assert.AreEqual(null, cell);
        }

        private readonly ConsoleLog logger = new ConsoleLog();
    }
}
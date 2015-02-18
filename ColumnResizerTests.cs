using System.IO;
using System.Linq;

using NUnit.Framework;

using SKBKontur.Catalogue.ExcelFileGenerator;
using SKBKontur.Catalogue.ExcelObjectPrinter;
using SKBKontur.Catalogue.ExcelObjectPrinter.ExcelDocumentPrimitivesImplementation;
using SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives;
using SKBKontur.Catalogue.ExcelObjectPrinter.PostBuildActions;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder;

namespace SKBKontur.Catalogue.Core.Tests.ExcelObjectPrinterTests
{
    [TestFixture]
    public class ColumnResizerTests
    {
        [Test]
        public void ColumnResizerTest()
        {
            var document = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(templateFileName));
            var template = new ExcelTable(document.GetWorksheet(0));
            var resizer = new ColumnResizer(template);

            var targetDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(emptyFileName));
            var target = new ExcelTable(targetDocument.GetWorksheet(0));
            resizer.ResizeColumns(new TableBuilder(target, new CellPosition("A1")));

            Assert.AreEqual(40.5703125, target.Columns.First().Width);
            Assert.AreEqual(25.7109375, target.Columns.Last().Width);

            //var result = targetDocument.GetDocumentBytes();
            //File.WriteAllBytes("output.xlsx", result);

            document.Dispose();
            targetDocument.Dispose();
        }

        private const string templateFileName = @"ExcelObjectPrinterTests\Files\template.xlsx";
        private const string emptyFileName = @"ExcelObjectPrinterTests\Files\empty.xlsx";
    }
}
using NUnit.Framework;

using SKBKontur.Catalogue.ExcelObjectPrinter.FakeDocumentPrimitivesImplementation;
using SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives;
using SKBKontur.Catalogue.ExcelObjectPrinter.PostBuildActions;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder;

namespace SKBKontur.Catalogue.Core.Tests.ExcelObjectPrinterTests
{
    [TestFixture]
    public class CellsMergerTests
    {
        [Test]
        public void Test()
        {
            var stringTemplate = new[]
                {
                    new[] {"", "", "", "", ""},
                    new[] {"", "", "", "", ""},
                    new[] {"", "", "", "", ""},
                    new[] {"", "Покупатель:", "Поставщик:", "", ""},
                    new[] {"", "", "", "", ""},
                    new[] {"", "", "", "", ""},
                    new[] {"", "", "", "", ""},
                    new[] {"", "MergeCells:A1:B2", "", "", ""},
                    new[] {"", "", "", "", "MergeCells:B5:D8"}
                };
            var template = FakeTable.GenerateFromStringArray(stringTemplate);
            var cellsMerger = new CellsMerger(template);
            var target = new FakeTable(100, 100);
            var tableBuilder = new TableBuilder(target, new CellPosition("A1"));
            cellsMerger.MergeCells(tableBuilder);

            Assert.AreEqual(new CellPosition("A1").CellReference, target.MergedCells[0].Item1.CellReference);
            Assert.AreEqual(new CellPosition("B2").CellReference, target.MergedCells[0].Item2.CellReference);
            Assert.AreEqual(new CellPosition("B5").CellReference, target.MergedCells[1].Item1.CellReference);
            Assert.AreEqual(new CellPosition("D8").CellReference, target.MergedCells[1].Item2.CellReference);
        }
    }
}
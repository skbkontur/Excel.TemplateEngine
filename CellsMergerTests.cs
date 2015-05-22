using System.Linq;

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

            var mergedCells = target.MergedCells.ToArray();
            Assert.AreEqual(new CellPosition("A1").CellReference, mergedCells[0].UpperLeft.CellReference);
            Assert.AreEqual(new CellPosition("B2").CellReference, mergedCells[0].LowerRight.CellReference);
            Assert.AreEqual(new CellPosition("B5").CellReference, mergedCells[1].UpperLeft.CellReference);
            Assert.AreEqual(new CellPosition("D8").CellReference, mergedCells[1].LowerRight.CellReference);
        }
    }
}
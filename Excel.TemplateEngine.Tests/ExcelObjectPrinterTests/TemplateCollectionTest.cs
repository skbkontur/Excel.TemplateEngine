using System.Linq;

using NUnit.Framework;

using SKBKontur.Catalogue.ExcelObjectPrinter.FakeDocumentPrimitivesImplementation;
using SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives;
using SKBKontur.Catalogue.ExcelObjectPrinter.RenderingTemplates;

namespace SKBKontur.Catalogue.Core.Tests.ExcelObjectPrinterTests
{
    [TestFixture]
    public class TemplateCollectionTest
    {
        [Test]
        public void TemplateExtractionTest()
        {
            var template = FakeTable.GenerateFromStringArray(stringTemplate);
            var templateCollection = new TemplateCollection(template);

            var rootTemplateContent = new[]
                {
                    new[] {"Тест:", "Value::Test"},
                    new[] {"Birne:", "Value::Pear"}
                };
            var rootTemplate = templateCollection.GetTemplate("RootTemplate");

            var rowNumber = 0;
            foreach (var row in rootTemplate.Content.Cells)
            {
                var cellNumber = 0;
                foreach (var cell in row)
                {
                    Assert.AreEqual(rootTemplateContent[rowNumber][cellNumber], cell.StringValue);
                    ++cellNumber;
                }
                ++rowNumber;
            }

            var thrashTemplate = templateCollection.GetTemplate("Thrash");
            Assert.AreEqual("Value::Metallica", thrashTemplate.Content.Cells.First().First().StringValue);

            var emptyTemplate = templateCollection.GetTemplate("Тест:");
            Assert.AreEqual(null, emptyTemplate);

            Assert.AreEqual("D3", thrashTemplate.Range.UpperLeft.CellReference);
            Assert.AreEqual("D3", thrashTemplate.Range.LowerRight.CellReference);
        }

        [Test]
        public void WithMergeCellsTemplateExtractionTest()
        {
            var template = FakeTable.GenerateFromStringArray(stringTemplate);
            template.MergeCells(new Rectangle(new CellPosition("A3"), new CellPosition("B3")));
            var templateCollection = new TemplateCollection(template);

            var rootTemplate = templateCollection.GetTemplate("RootTemplate");
            var mergedCells = rootTemplate.MergedCells.ToArray();

            Assert.AreEqual(1, mergedCells.Length);
            Assert.AreEqual("A2", mergedCells[0].UpperLeft.CellReference);
            Assert.AreEqual("B2", mergedCells[0].LowerRight.CellReference);
        }

        private readonly string[][] stringTemplate =
            {
                new[] {"Template:RootTemplate:A2:B3", "", "", ""},
                new[] {"Тест:", "Value::Test", "", "Template:Thrash:D3:D3"},
                new[] {"Birne:", "Value::Pear", "", "Value::Metallica"}
            };
    }
}
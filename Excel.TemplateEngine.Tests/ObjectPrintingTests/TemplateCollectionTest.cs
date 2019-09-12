using System.Linq;

using FluentAssertions;

using NUnit.Framework;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives.Implementations;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.RenderingTemplates;
using SkbKontur.Excel.TemplateEngine.Tests.ObjectPrintingTests.FakeDocumentPrimitivesImplementation;

namespace SkbKontur.Excel.TemplateEngine.Tests.ObjectPrintingTests
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
                    cell.StringValue.Should().Be(rootTemplateContent[rowNumber][cellNumber]);
                    ++cellNumber;
                }
                ++rowNumber;
            }

            var thrashTemplate = templateCollection.GetTemplate("Thrash");
            thrashTemplate.Content.Cells.First().First().StringValue.Should().Be("Value::Metallica");

            var emptyTemplate = templateCollection.GetTemplate("Тест:");
            emptyTemplate.Should().BeNull();

            thrashTemplate.Range.UpperLeft.CellReference.Should().Be("D3");
            thrashTemplate.Range.LowerRight.CellReference.Should().Be("D3");
        }

        [Test]
        public void WithMergeCellsTemplateExtractionTest()
        {
            var template = FakeTable.GenerateFromStringArray(stringTemplate);
            template.MergeCells(new Rectangle(new CellPosition("A3"), new CellPosition("B3")));
            var templateCollection = new TemplateCollection(template);

            var rootTemplate = templateCollection.GetTemplate("RootTemplate");
            var mergedCells = rootTemplate.MergedCells.ToArray();

            mergedCells.Length.Should().Be(1);
            mergedCells[0].UpperLeft.CellReference.Should().Be("A2");
            mergedCells[0].LowerRight.CellReference.Should().Be("B2");
        }

        private readonly string[][] stringTemplate =
            {
                new[] {"Template:RootTemplate:A2:B3", "", "", ""},
                new[] {"Тест:", "Value::Test", "", "Template:Thrash:D3:D3"},
                new[] {"Birne:", "Value::Pear", "", "Value::Metallica"}
            };
    }
}
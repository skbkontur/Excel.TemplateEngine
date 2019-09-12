using SkbKontur.Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;

namespace SkbKontur.Excel.TemplateEngine.Tests.ObjectPrintingTests.FakeDocumentPrimitivesImplementation
{
    internal class FakeCell : ICell
    {
        public FakeCell(ICellPosition position)
        {
            CellPosition = position;
        }

        public string StringValue { get; set; }
        public CellType CellType { get; set; }
        public ICellPosition CellPosition { get; }

        public void CopyStyle(ICell templateCell)
        {
            var fakeCell = (FakeCell)templateCell;
            StyleId = fakeCell.StyleId;
        }

        public string StyleId { get; set; }
    }
}
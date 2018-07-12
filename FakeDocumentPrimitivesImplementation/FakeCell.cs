using SKBKontur.Catalogue.ExcelObjectPrinter.DataTypes;
using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;
using SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.FakeDocumentPrimitivesImplementation
{
    public class FakeCell : ICell
    {
        public FakeCell(ICellPosition position)
        {
            CellPosition = position;
        }

        public string StringValue { get; set; }
        public CellType CellType { get; set; }
        public ICellPosition CellPosition { get; private set; }

        public void CopyStyle(ICell templateCell)
        {
            var fakeCell = (FakeCell)templateCell;
            StyleId = fakeCell.StyleId;
        }

        public string StyleId { get; set; }
    }
}
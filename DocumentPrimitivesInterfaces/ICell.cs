using SKBKontur.Catalogue.ExcelObjectPrinter.DataTypes;
using SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces
{
    public interface ICell
    {
        void CopyStyle(ICell templateCell);
        string StringValue { get; set; }
        CellType CellType { set; }
        ICellPosition CellPosition { get; }
    }
}
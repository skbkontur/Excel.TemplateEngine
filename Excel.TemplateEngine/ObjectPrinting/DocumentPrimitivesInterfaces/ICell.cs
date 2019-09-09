using Excel.TemplateEngine.ObjectPrinting.DataTypes;
using Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;

namespace Excel.TemplateEngine.ObjectPrinting.DocumentPrimitivesInterfaces
{
    public interface ICell
    {
        void CopyStyle(ICell templateCell);
        string StringValue { get; set; }
        CellType CellType { set; }
        ICellPosition CellPosition { get; }
    }
}
using Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;

namespace Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives
{
    public interface ICell
    {
        void CopyStyle(ICell templateCell);
        string StringValue { get; set; }
        CellType CellType { set; }
        ICellPosition CellPosition { get; }
    }
}
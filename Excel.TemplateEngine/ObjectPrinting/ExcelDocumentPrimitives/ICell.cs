using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives
{
    public interface ICell
    {
        void CopyStyle(ICell templateCell);
        string StringValue { get; set; }
        CellType CellType { set; }
        ICellPosition CellPosition { get; }
    }
}
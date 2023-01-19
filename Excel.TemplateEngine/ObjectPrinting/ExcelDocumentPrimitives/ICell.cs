using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives
{
    public interface ICell
    {
        void CopyStyle(ICell templateCell);
        string StringValue { get; }
        void SetValue(string stringValue, CellType cellType = CellType.String);
        ICellPosition CellPosition { get; }
    }
}
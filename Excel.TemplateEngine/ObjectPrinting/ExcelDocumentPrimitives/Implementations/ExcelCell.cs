using SkbKontur.Excel.TemplateEngine.FileGenerating.Primitives;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives.Implementations;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives.Implementations
{
    public class ExcelCell : ICell
    {
        public ExcelCell(IExcelCell excelCell)
        {
            internalCell = excelCell;
        }

        public string StringValue => internalCell.GetStringValue();

        public void CopyStyle(ICell templateCell)
        {
            var excelCell = ((ExcelCell)templateCell);
            internalCell.SetStyle(excelCell.internalCell.GetStyle());
        }

        public void SetValue(string stringValue, CellType cellType)
        {
            if (cellType == CellType.String)
                internalCell.SetStringValue(stringValue);
            else
                internalCell.SetNumericValue(stringValue);
        }

        public ICellPosition CellPosition => new CellPosition(internalCell.GetCellIndex());

        private readonly IExcelCell internalCell;
    }
}
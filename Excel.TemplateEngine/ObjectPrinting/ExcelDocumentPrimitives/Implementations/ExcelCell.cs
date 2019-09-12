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

        public string StringValue
        {
            get => internalCell.GetStringValue();
            set => internalCell.SetStringValue(value);
        }

        public CellType CellType
        {
            set
            {
                if (value == CellType.String)
                    internalCell.SetStringValue(StringValue);
                else
                    internalCell.SetNumericValue(StringValue);
            }
        }

        public void CopyStyle(ICell templateCell)
        {
            var excelCell = ((ExcelCell)templateCell);
            internalCell.SetStyle(excelCell.internalCell.GetStyle());
        }

        public ICellPosition CellPosition => new CellPosition(internalCell.GetCellIndex());

        private readonly IExcelCell internalCell;
    }
}
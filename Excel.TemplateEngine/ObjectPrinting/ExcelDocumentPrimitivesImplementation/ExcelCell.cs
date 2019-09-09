using Excel.TemplateEngine.ObjectPrinting.DataTypes;
using Excel.TemplateEngine.ObjectPrinting.DocumentPrimitivesInterfaces;
using Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;

namespace Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitivesImplementation
{
    public class ExcelCell : ICell
    {
        public ExcelCell(IExcelCell excelCell)
        {
            internalCell = excelCell;
        }

        public string StringValue { get { return internalCell.GetStringValue(); } set { internalCell.SetStringValue(value); } }

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

        public ICellPosition CellPosition { get { return new CellPosition(internalCell.GetCellIndex()); } }

        private readonly IExcelCell internalCell;
    }
}
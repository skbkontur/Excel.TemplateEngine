using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SKBKontur.Catalogue.ExcelFileGenerator
{
    public interface IExcelCell
    {
        void SetStringValue(string value);
    }

    internal class ExcelCell : IExcelCell
    {
        public ExcelCell(Cell cell)
        {
            this.cell = cell;
        }

        public void SetStringValue(string value)
        {
            cell.CellValue = new CellValue(value);
            cell.DataType = new EnumValue<CellValues>(CellValues.String);
        }

        private readonly Cell cell;
    }
}
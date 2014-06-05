using System.Globalization;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SKBKontur.Catalogue.ExcelFileGenerator
{
    public interface IExcelCell
    {
        void SetStringValue(string value);
        void SetNumericValue(double value);
        void SetStyle(ExcelCellStyle style);
    }

    internal class ExcelCell : IExcelCell
    {
        public ExcelCell(Cell cell, IExcelDocumentStyle documentStyle)
        {
            this.cell = cell;
            this.documentStyle = documentStyle;
        }

        public void SetStringValue(string value)
        {
            cell.CellValue = new CellValue(value);
            cell.DataType = new EnumValue<CellValues>(CellValues.String);
        }

        public void SetNumericValue(double value)
        {
            cell.CellValue = new CellValue(value.ToString(CultureInfo.InvariantCulture));
            cell.DataType = new EnumValue<CellValues>(CellValues.Number);
        }

        public void SetStyle(ExcelCellStyle style)
        {
            cell.StyleIndex = documentStyle.SaveStyle(style);
        }

        private readonly Cell cell;
        private readonly IExcelDocumentStyle documentStyle;
    }
}
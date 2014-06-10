using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Interfaces
{
    public interface IExcelCell
    {
        void SetStringValue(string value);
        void SetNumericValue(double value);
        void SetStyle(ExcelCellStyle style);
        void SetFormattedStringValue(FormattedStringValue value);
    }
}
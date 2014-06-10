using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Interfaces
{
    public interface IExcelCell
    {
        void SetNumericValue(double value);
        void SetStringValue(string value);
        void SetFormattedStringValue(FormattedStringValue value);
        void SetStyle(ExcelCellStyle style);
    }
}
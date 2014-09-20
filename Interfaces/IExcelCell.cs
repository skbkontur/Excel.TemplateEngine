using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Interfaces
{
    public interface IExcelCell
    {
        IExcelCell SetNumericValue(double value);
        IExcelCell SetNumericValue(decimal value);
        IExcelCell SetStringValue(string value);
        IExcelCell SetFormattedStringValue(FormattedStringValue value);
        IExcelCell SetStyle(ExcelCellStyle style);
        
        string GetStringValue();
    }
}
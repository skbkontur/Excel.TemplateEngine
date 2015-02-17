using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;
using SKBKontur.Catalogue.ExcelFileGenerator.Implementation;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Interfaces
{
    public interface IExcelCell
    {
        IExcelCell SetNumericValue(double value);
        IExcelCell SetNumericValue(string value);
        IExcelCell SetNumericValue(decimal value);
        IExcelCell SetStringValue(string value);
        IExcelCell SetFormattedStringValue(FormattedStringValue value);
        IExcelCell SetStyle(ExcelCellStyle style);
        ExcelCellStyle GetStyle();

        string GetStringValue();
        ExcelCellIndex GetCellReference();
    }
}
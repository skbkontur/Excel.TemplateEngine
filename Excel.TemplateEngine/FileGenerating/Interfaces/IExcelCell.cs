using Excel.TemplateEngine.FileGenerating.DataTypes;
using Excel.TemplateEngine.FileGenerating.Implementation;

namespace Excel.TemplateEngine.FileGenerating.Interfaces
{
    public interface IExcelCell
    {
        IExcelCell SetNumericValue(double value);
        IExcelCell SetNumericValue(string value);
        IExcelCell SetNumericValue(decimal value);
        IExcelCell SetStringValue(string value);
        IExcelCell SetFormattedStringValue(FormattedStringValue value);
        IExcelCell SetFormula(string formula);
        IExcelCell SetStyle(ExcelCellStyle style);
        ExcelCellStyle GetStyle();

        string GetStringValue();
        ExcelCellIndex GetCellIndex();
    }
}
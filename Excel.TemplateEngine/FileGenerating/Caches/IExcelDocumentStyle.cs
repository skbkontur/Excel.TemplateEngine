using Excel.TemplateEngine.FileGenerating.DataTypes;

namespace Excel.TemplateEngine.FileGenerating.Caches
{
    internal interface IExcelDocumentStyle
    {
        uint AddStyle(ExcelCellStyle style);
        ExcelCellStyle GetStyle(int styleIndex);
    }
}
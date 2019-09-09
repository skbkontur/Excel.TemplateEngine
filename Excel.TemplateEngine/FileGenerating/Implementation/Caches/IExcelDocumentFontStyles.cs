using Excel.TemplateEngine.FileGenerating.DataTypes;

namespace Excel.TemplateEngine.FileGenerating.Implementation.Caches
{
    internal interface IExcelDocumentFontStyles
    {
        uint AddFont(ExcelCellFontStyle style);
    }
}
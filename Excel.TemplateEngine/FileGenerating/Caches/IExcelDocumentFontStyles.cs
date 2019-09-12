using Excel.TemplateEngine.FileGenerating.DataTypes;

namespace Excel.TemplateEngine.FileGenerating.Caches
{
    internal interface IExcelDocumentFontStyles
    {
        uint AddFont(ExcelCellFontStyle style);
    }
}
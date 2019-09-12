using SkbKontur.Excel.TemplateEngine.FileGenerating.DataTypes;

namespace SkbKontur.Excel.TemplateEngine.FileGenerating.Caches
{
    internal interface IExcelDocumentFontStyles
    {
        uint AddFont(ExcelCellFontStyle style);
    }
}
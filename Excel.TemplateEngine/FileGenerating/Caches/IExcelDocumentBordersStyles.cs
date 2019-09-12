using SkbKontur.Excel.TemplateEngine.FileGenerating.DataTypes;

namespace SkbKontur.Excel.TemplateEngine.FileGenerating.Caches
{
    internal interface IExcelDocumentBordersStyles
    {
        uint AddStyle(ExcelCellBordersStyle style);
    }
}
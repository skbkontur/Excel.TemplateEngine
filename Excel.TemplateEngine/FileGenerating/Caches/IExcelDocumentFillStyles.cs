using SkbKontur.Excel.TemplateEngine.FileGenerating.DataTypes;

namespace SkbKontur.Excel.TemplateEngine.FileGenerating.Caches
{
    internal interface IExcelDocumentFillStyles
    {
        uint AddStyle(ExcelCellFillStyle style);
    }
}
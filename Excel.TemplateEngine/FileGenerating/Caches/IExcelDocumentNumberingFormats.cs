using SkbKontur.Excel.TemplateEngine.FileGenerating.DataTypes;

namespace SkbKontur.Excel.TemplateEngine.FileGenerating.Caches
{
    internal interface IExcelDocumentNumberingFormats
    {
        uint AddFormat(ExcelCellNumberingFormat format);
    }
}
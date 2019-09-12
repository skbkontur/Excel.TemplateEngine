using Excel.TemplateEngine.FileGenerating.DataTypes;

namespace Excel.TemplateEngine.FileGenerating.Caches
{
    internal interface IExcelDocumentFillStyles
    {
        uint AddStyle(ExcelCellFillStyle style);
    }
}
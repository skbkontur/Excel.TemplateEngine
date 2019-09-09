using Excel.TemplateEngine.FileGenerating.DataTypes;

namespace Excel.TemplateEngine.FileGenerating.Implementation.Caches
{
    internal interface IExcelDocumentFillStyles
    {
        uint AddStyle(ExcelCellFillStyle style);
    }
}
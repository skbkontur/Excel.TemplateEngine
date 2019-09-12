using Excel.TemplateEngine.FileGenerating.DataTypes;

namespace Excel.TemplateEngine.FileGenerating.Caches
{
    internal interface IExcelDocumentBordersStyles
    {
        uint AddStyle(ExcelCellBordersStyle style);
    }
}
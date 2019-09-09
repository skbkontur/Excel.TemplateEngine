using Excel.TemplateEngine.FileGenerating.DataTypes;

namespace Excel.TemplateEngine.FileGenerating.Implementation.Caches
{
    internal interface IExcelDocumentBordersStyles
    {
        uint AddStyle(ExcelCellBordersStyle style);
    }
}
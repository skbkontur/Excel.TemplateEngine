using Excel.TemplateEngine.FileGenerating.DataTypes;

namespace Excel.TemplateEngine.FileGenerating.Implementation.Caches
{
    internal interface IExcelDocumentNumberingFormats
    {
        uint AddFormat(ExcelCellNumberingFormat format);
    }
}
using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Caches
{
    internal interface IExcelDocumentNumberingFormats
    {
        uint AddFormat(ExcelCellNumberingFormat format);
    }
}
using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Caches
{
    internal interface IExcelDocumentBordersStyles
    {
        uint AddStyle(ExcelCellBordersStyle style);
    }
}
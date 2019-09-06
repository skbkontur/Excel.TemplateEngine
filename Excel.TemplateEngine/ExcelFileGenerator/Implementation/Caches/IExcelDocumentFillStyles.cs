using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Caches
{
    internal interface IExcelDocumentFillStyles
    {
        uint AddStyle(ExcelCellFillStyle style);
    }
}
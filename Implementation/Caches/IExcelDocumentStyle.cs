using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Caches
{
    internal interface IExcelDocumentStyle
    {
        uint AddStyle(ExcelCellStyle style);
        void Save();
    }
}
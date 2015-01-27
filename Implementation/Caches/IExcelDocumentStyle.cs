using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Caches
{
    public interface IExcelDocumentStyle
    {
        uint AddStyle(ExcelCellStyle style);
        void Save();
    }
}
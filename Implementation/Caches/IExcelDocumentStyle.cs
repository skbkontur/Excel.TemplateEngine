using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Caches
{
    public interface IExcelDocumentStyle
    {
        uint AddStyle(ExcelCellStyle style);
        ExcelCellStyle GetStyle(int styleIndex);
        void Save();
    }
}
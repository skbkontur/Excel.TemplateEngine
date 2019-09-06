using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Caches
{
    public interface IExcelSharedStrings
    {
        uint AddSharedString(FormattedStringValue value);
        string GetSharedString(uint index);
    }
}
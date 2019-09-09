using Excel.TemplateEngine.FileGenerating.DataTypes;

namespace Excel.TemplateEngine.FileGenerating.Implementation.Caches
{
    public interface IExcelSharedStrings
    {
        uint AddSharedString(FormattedStringValue value);
        string GetSharedString(uint index);
    }
}
using Excel.TemplateEngine.FileGenerating.DataTypes;

namespace Excel.TemplateEngine.FileGenerating.Caches
{
    internal interface IExcelSharedStrings
    {
        uint AddSharedString(FormattedStringValue value);
        string GetSharedString(uint index);
    }
}
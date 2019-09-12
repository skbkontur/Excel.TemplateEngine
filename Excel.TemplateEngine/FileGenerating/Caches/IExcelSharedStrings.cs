using SkbKontur.Excel.TemplateEngine.FileGenerating.DataTypes;

namespace SkbKontur.Excel.TemplateEngine.FileGenerating.Caches
{
    internal interface IExcelSharedStrings
    {
        uint AddSharedString(FormattedStringValue value);
        string GetSharedString(uint index);
    }
}
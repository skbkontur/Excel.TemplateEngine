using System.Linq;

namespace SkbKontur.Excel.TemplateEngine.FileGenerating.DataTypes
{
    public class ExcelCellNumberingFormat
    {
        public ExcelCellNumberingFormat(int precision)
        {
            FormatCode = "0." + string.Join("", Enumerable.Repeat("0", precision));
        }

        public ExcelCellNumberingFormat(string formatCode)
        {
            FormatCode = formatCode;
        }

        public override string ToString() => $"FormatCode = {FormatCode}";

        public string FormatCode { get; }
    }
}
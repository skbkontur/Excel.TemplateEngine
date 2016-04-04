using System.Linq;

namespace SKBKontur.Catalogue.ExcelFileGenerator.DataTypes
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

        public override string ToString()
        {
            return string.Format("FormatCode = {0}", FormatCode);
        }

        public string FormatCode { get; private set; }
    }
}
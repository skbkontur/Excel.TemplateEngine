using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SKBKontur.Catalogue.ExcelFileGenerator
{
    internal interface IExcelDocumentStyle
    {
        void Save();
        uint CreateNumericTableStyle(int precision);
    }

    internal class ExcelDocumentStyle : IExcelDocumentStyle
    {
        public ExcelDocumentStyle(Stylesheet stylesheet)
        {
            this.stylesheet = stylesheet;
        }

        public void Save()
        {
            stylesheet.Save();
        }

        public uint CreateNumericTableStyle(int precision)
        {
            var formatCode = "0." + string.Join("", Enumerable.Repeat("0", precision));
            if(stylesheet.NumberingFormats == null)
            {
                var numberingFormats = new NumberingFormats {Count = new UInt32Value(0u)};
                stylesheet.InsertAt(numberingFormats, 0);
            }
            var numberFormatId = ++stylesheet.NumberingFormats.Count;
            stylesheet.NumberingFormats.AppendChild(new NumberingFormat
                {
                    FormatCode = new StringValue(formatCode),
                    NumberFormatId = numberFormatId
                });
            var styleFormatId = stylesheet.CellFormats.Count.Value;
            stylesheet.CellFormats.Count++;
            stylesheet.CellFormats.AppendChild(new CellFormat
                {
                    NumberFormatId = numberFormatId,
                    FormatId = 0,
                    FontId = 0,
                    FillId = 0,
                    BorderId = 0,
                    ApplyNumberFormat = new BooleanValue(true)
                });
            return styleFormatId;
        }

        private readonly Stylesheet stylesheet;
    }
}
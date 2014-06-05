using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SKBKontur.Catalogue.ExcelFileGenerator
{
    internal interface IExcelDocumentStyle
    {
        void Save();
        uint CreateNumericTableStyle(int precision);
        uint SaveStyle(ExcelCellStyle style);
    }

    internal class ExcelDocumentStyle : IExcelDocumentStyle
    {
        public ExcelDocumentStyle(Stylesheet stylesheet)
        {
            this.stylesheet = stylesheet;
            numberingFormats = new ExcelDocumentNumberingFormats(stylesheet);
            fillStyles = new ExcelDocumentFillStyles(stylesheet);
        }

        public void Save()
        {
            stylesheet.Save();
        }

        public uint CreateNumericTableStyle(int precision)
        {
            var numberFormatId = numberingFormats.AddFormat(new ExcelCellNumberingFormat {Precision = precision});
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

        public uint SaveStyle(ExcelCellStyle style)
        {
            var fillId = fillStyles.AddStyle(style.FillStyle);
            var styleFormatId = stylesheet.CellFormats.Count.Value;
            stylesheet.CellFormats.Count++;
            stylesheet.CellFormats.AppendChild(new CellFormat
                {
                    NumberFormatId = 0,
                    FormatId = 0,
                    FontId = 0,
                    FillId = fillId,
                    BorderId = 0,
                    ApplyFill = fillId == 0 ? null : new BooleanValue(true)
                });
            return styleFormatId;
        }

        private readonly Stylesheet stylesheet;
        private readonly ExcelDocumentNumberingFormats numberingFormats;
        private readonly ExcelDocumentFillStyles fillStyles;
    }
}
using System;

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
            throw new NotImplementedException();
        }

        private readonly Stylesheet stylesheet;
        private readonly ExcelDocumentNumberingFormats numberingFormats;
    }
}
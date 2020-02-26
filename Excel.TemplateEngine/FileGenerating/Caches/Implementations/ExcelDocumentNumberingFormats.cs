using System.Collections.Generic;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

using SkbKontur.Excel.TemplateEngine.FileGenerating.DataTypes;

namespace SkbKontur.Excel.TemplateEngine.FileGenerating.Caches.Implementations
{
    internal class ExcelDocumentNumberingFormats : IExcelDocumentNumberingFormats
    {
        public ExcelDocumentNumberingFormats(Stylesheet stylesheet)
        {
            this.stylesheet = stylesheet;
            cache = new HashSet<uint>();
        }

        public uint AddFormat(ExcelCellNumberingFormat format)
        {
            if (format == null)
                return 0;

            if (format.Code == null || cache.Contains(format.Id))
                return format.Id;

            if (stylesheet.NumberingFormats == null)
            {
                var numberingFormats = new NumberingFormats {Count = new UInt32Value(0u)};
                stylesheet.InsertAt(numberingFormats, 0);
            }

            stylesheet.NumberingFormats.AppendChild(new NumberingFormat {FormatCode = new StringValue(format.Code), NumberFormatId = format.Id});
            cache.Add(format.Id);
            return format.Id;
        }

        private readonly Stylesheet stylesheet;

        private readonly HashSet<uint> cache;
    }
}
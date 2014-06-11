using System.Collections.Generic;

using DocumentFormat.OpenXml.Spreadsheet;

using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Caches
{
    internal class ExcelDocumentStyle : IExcelDocumentStyle
    {
        public ExcelDocumentStyle(Stylesheet stylesheet)
        {
            this.stylesheet = stylesheet;
            numberingFormats = new ExcelDocumentNumberingFormats(stylesheet);
            fillStyles = new ExcelDocumentFillStyles(stylesheet);
            bordersStyles = new ExcelDocumentBordersStyles(stylesheet);
            fontStyles = new ExcelDocumentFontStyles(stylesheet);
            cache = new Dictionary<CellStyleCacheItem, uint>();
        }

        public void Save()
        {
            stylesheet.Save();
        }

        public uint AddStyle(ExcelCellStyle style)
        {
            var fillId = fillStyles.AddStyle(style.FillStyle);
            var fontId = fontStyles.AddFont(style.FontStyle);
            var borderId = bordersStyles.AddStyle(style.BordersStyle);
            var numberFormatId = numberingFormats.AddFormat(style.NumberingFormat);
            var alignment = Alignment(style.Alignment);
            var cacheItem = new CellStyleCacheItem
                {
                    FillId = fillId,
                    FontId = fontId,
                    BorderId = borderId,
                    NumberFormatId = numberFormatId,
                    Alignment = alignment
                };
            uint result;
            if(!cache.TryGetValue(cacheItem, out result))
            {
                result = stylesheet.CellFormats.Count;
                stylesheet.CellFormats.Count++;
                stylesheet.CellFormats.AppendChild(cacheItem.ToCellFormat());
                cache.Add(cacheItem, result);
            }
            return result;
        }

        private static AlignmentCacheItem Alignment(ExcelCellAlignment cellAlignment)
        {
            if(cellAlignment == null)
                return null;
            if(cellAlignment.HorizontalAlignment == ExcelHorizontalAlignment.Default && cellAlignment.VerticalAlignment == ExcelVerticalAlignment.Default && !cellAlignment.WrapText)
                return null;
            return new AlignmentCacheItem(cellAlignment);
        }

        private readonly Stylesheet stylesheet;
        private readonly ExcelDocumentNumberingFormats numberingFormats;
        private readonly ExcelDocumentFillStyles fillStyles;
        private readonly ExcelDocumentBordersStyles bordersStyles;
        private readonly IExcelDocumentFontStyles fontStyles;
        private readonly IDictionary<CellStyleCacheItem, uint> cache;
    }
}
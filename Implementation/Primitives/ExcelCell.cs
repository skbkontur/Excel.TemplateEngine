using System.Globalization;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;
using SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Caches;
using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Primitives
{
    internal class ExcelCell : IExcelCell
    {
        public ExcelCell(Cell cell, IExcelDocumentStyle documentStyle, IExcelSharedStrings excelSharedStrings)
        {
            this.cell = cell;
            this.documentStyle = documentStyle;
            this.excelSharedStrings = excelSharedStrings;
        }

        public IExcelCell SetStringValue(string value)
        {
            return SetFormattedStringValue(new FormattedStringValue
                {
                    Blocks = new[]
                        {
                            new FormattedStringValueBlock
                                {
                                    Value = value
                                }
                        }
                });
        }

        public IExcelCell SetNumericValue(double value)
        {
            cell.CellValue = new CellValue(value.ToString(CultureInfo.InvariantCulture));
            cell.DataType = new EnumValue<CellValues>(CellValues.Number);
            return this;
        }

        public IExcelCell SetStyle(ExcelCellStyle style)
        {
            cell.StyleIndex = documentStyle.AddStyle(style);
            return this;
        }

        public IExcelCell SetFormattedStringValue(FormattedStringValue value)
        {
            var index = excelSharedStrings.AddSharedString(value);
            cell.CellValue = new CellValue(index.ToString(CultureInfo.InvariantCulture));
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            return this;
        }

        private readonly Cell cell;
        private readonly IExcelDocumentStyle documentStyle;
        private readonly IExcelSharedStrings excelSharedStrings;
    }
}
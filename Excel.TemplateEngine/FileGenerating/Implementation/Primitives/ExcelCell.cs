using System.Globalization;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

using Excel.TemplateEngine.FileGenerating.DataTypes;
using Excel.TemplateEngine.FileGenerating.Implementation.Caches;
using Excel.TemplateEngine.FileGenerating.Interfaces;

namespace Excel.TemplateEngine.FileGenerating.Implementation.Primitives
{
    public class ExcelCell : IExcelCell
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

        public IExcelCell SetNumericValue(string value)
        {
            cell.CellValue = new CellValue(value);
            cell.DataType = new EnumValue<CellValues>(CellValues.Number);
            return this;
        }

        public IExcelCell SetNumericValue(decimal value)
        {
            cell.CellValue = new CellValue(value.ToString(CultureInfo.InvariantCulture));
            cell.DataType = new EnumValue<CellValues>(CellValues.Number);
            return this;
        }

        public IExcelCell SetStyle(ExcelCellStyle style)
        {
            if (style != null)
                cell.StyleIndex = documentStyle.AddStyle(style);
            return this;
        }

        public ExcelCellStyle GetStyle()
        {
            return cell?.StyleIndex == null ? null : documentStyle.GetStyle((int)cell.StyleIndex.Value);
        }

        public string GetStringValue()
        {
            if (cell?.DataType?.Value == CellValues.SharedString)
                return excelSharedStrings.GetSharedString(uint.Parse(cell.CellValue.Text));
            return cell?.CellValue?.Text;
        }

        public ExcelCellIndex GetCellIndex()
        {
            return new ExcelCellIndex(cell.CellReference.Value);
        }

        public IExcelCell SetFormattedStringValue(FormattedStringValue value)
        {
            var index = excelSharedStrings.AddSharedString(value);
            cell.CellValue = new CellValue(index.ToString(CultureInfo.InvariantCulture));
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            return this;
        }

        public IExcelCell SetFormula(string formula)
        {
            cell.CellFormula = new CellFormula {Text = formula};
            return this;
        }

        private readonly Cell cell;
        private readonly IExcelDocumentStyle documentStyle;
        private readonly IExcelSharedStrings excelSharedStrings;
    }
}
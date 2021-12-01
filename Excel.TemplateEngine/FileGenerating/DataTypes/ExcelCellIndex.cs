using System.Text.RegularExpressions;

namespace SkbKontur.Excel.TemplateEngine.FileGenerating.DataTypes
{
    public class ExcelCellIndex
    {
        public ExcelCellIndex(int row, int column)
        {
            RowIndex = row;
            ColumnIndex = column;
            CellReference = ToCellReference(row, column);
        }

        public ExcelCellIndex(string cellReference)
        {
            CellReference = cellReference;
            RowIndex = ToRowIndex(cellReference);
            ColumnIndex = ToColumnIndex(cellReference);
        }

        public ExcelCellIndex Add(ExcelCellIndex other)
        {
            return new ExcelCellIndex(RowIndex + other.RowIndex - 1, ColumnIndex + other.ColumnIndex - 1);
        }

        public ExcelCellIndex Subtract(ExcelCellIndex other)
        {
            return new ExcelCellIndex(RowIndex - other.RowIndex + 1, ColumnIndex - other.ColumnIndex + 1);
        }

        public string CellReference { get; }
        public int RowIndex { get; }
        public int ColumnIndex { get; }

        private static string ToCellReference(int rowIndex, int columnIndex)
            => $"{ToColumnName(columnIndex)}{rowIndex}";

        private static string ToCellReference(uint rowIndex, int columnIndex)
            => ToCellReference((int)rowIndex, columnIndex);

        private static int ToRowIndex(string cellReference)
            => int.Parse(new Regex("[A-Z]+").Replace(cellReference, ""));

        private static int ToColumnIndex(string cellReference)
        {
            var prefix = new Regex("[0-9]+").Replace(cellReference, "");
            var result = 0;
            foreach (var c in prefix)
            {
                result *= 26;
                result += c - 'A' + 1;
            }
            return result;
        }

        private static string ToColumnName(int columnIndex)
        {
            columnIndex -= 1;
            var prefixIndex = columnIndex / 26;
            var tail = ((char)(columnIndex % 26 + 'A')).ToString();
            return prefixIndex > 0 ? ToColumnName(prefixIndex) + tail : tail;
        }
    }
}
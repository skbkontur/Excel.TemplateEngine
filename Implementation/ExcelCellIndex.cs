using System.Text.RegularExpressions;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation
{
    public class ExcelCellIndex
    {
        public ExcelCellIndex(int row, int column)
        {
            this.row = row;
            this.column = column;
            cellReference = ToCellReference(row, column);
        }

        public ExcelCellIndex(string cellReference)
        {
            this.cellReference = cellReference;
            row = ToRowIndex(cellReference);
            column = ToColumnIndex(cellReference);
        }

        public ExcelCellIndex Add(ExcelCellIndex other)
        {
            return new ExcelCellIndex(RowIndex + other.RowIndex - 1, ColumnIndex + other.ColumnIndex - 1);
        }

        public ExcelCellIndex Subtract(ExcelCellIndex other)
        {
            return new ExcelCellIndex(RowIndex - other.RowIndex + 1, ColumnIndex - other.ColumnIndex + 1);
        }

        public string CellReference { get { return cellReference; } }
        public int RowIndex { get { return row; } }
        public int ColumnIndex { get { return column; } }

        private static string ToCellReference(int rowIndex, int columnIndex)
        {
            return string.Format("{0}{1}", ToColumnName(columnIndex), rowIndex);
        }

        private static string ToCellReference(uint rowIndex, int columnIndex)
        {
            return ToCellReference((int)rowIndex, columnIndex);
        }

        private static int ToRowIndex(string cellReference)
        {
            return int.Parse(new Regex("[A-Z]+").Replace(cellReference, ""));
        }

        private static int ToColumnIndex(string cellReference)
        {
            var prefix = new Regex("[0-9]+").Replace(cellReference, "");
            var result = 0;
            foreach(var c in prefix)
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

        private readonly string cellReference;
        private readonly int row;
        private readonly int column;
    }
}
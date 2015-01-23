using System.Text.RegularExpressions;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation
{
    internal static class IndexHelpers
    {
        public static string ToCellName(int rowIndex, int columnIndex)
        {
            return string.Format("{0}{1}", ToColumnName(columnIndex), rowIndex);
        }

        public static string ToCellName(uint rowIndex, int columnIndex)
        {
            return ToCellName((int)rowIndex, columnIndex);
        }

        public static int GetRowIndex(string cellReference)
        {
            return int.Parse(new Regex("[A-Z]+").Replace(cellReference, "")) - 1;
        }

        public static int GetColumnIndex(string cellReference)
        {
            var prefix = new Regex("[0-9]+").Replace(cellReference, "");
            var result = 0;
            foreach(var c in prefix)
            {
                result *= 26;
                result += c - 'A';
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
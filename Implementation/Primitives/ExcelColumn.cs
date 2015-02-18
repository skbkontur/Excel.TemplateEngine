using DocumentFormat.OpenXml.Spreadsheet;

using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Primitives
{
    public class ExcelColumn : IExcelColumn
    {
        public ExcelColumn(Column column)
        {
            this.column = column;
        }

        public double Width { get { return column.Width; } }
        public int Index { get { return (int)column.Min.Value; } }

        private readonly Column column;
    }
}
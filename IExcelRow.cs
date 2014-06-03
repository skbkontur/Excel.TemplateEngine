using DocumentFormat.OpenXml.Spreadsheet;

namespace SKBKontur.Catalogue.ExcelFileGenerator
{
    public interface IExcelRow
    {
        IExcelCell CreateCell(int index);
    }

    internal class ExcelRow : IExcelRow
    {
        public ExcelRow(Row row)
        {
            this.row = row;
        }

        private readonly Row row;
        public IExcelCell CreateCell(int index)
        {
            var cell = new Cell
                {
                    CellReference = IndexHelpers.ToColumnName(index) + row.RowIndex
                };
            row.AppendChild(cell);
            return new ExcelCell(cell);
        }
    }
}
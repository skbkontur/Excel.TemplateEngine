using DocumentFormat.OpenXml.Spreadsheet;

namespace SKBKontur.Catalogue.ExcelFileGenerator
{
    public interface IExcelRow
    {
        IExcelCell CreateCell(int index);
    }

    internal class ExcelRow : IExcelRow
    {
        public ExcelRow(Row row, IExcelDocumentStyle documentStyle, ISharedStringsCache sharedStringsCache)
        {
            this.row = row;
            this.documentStyle = documentStyle;
            this.sharedStringsCache = sharedStringsCache;
        }

        public IExcelCell CreateCell(int index)
        {
            var cell = new Cell
                {
                    CellReference = IndexHelpers.ToColumnName(index) + row.RowIndex
                };
            row.AppendChild(cell);
            return new ExcelCell(cell, documentStyle, sharedStringsCache);
        }

        private readonly Row row;
        private readonly IExcelDocumentStyle documentStyle;
        private readonly ISharedStringsCache sharedStringsCache;
    }
}
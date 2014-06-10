using DocumentFormat.OpenXml.Spreadsheet;

namespace SKBKontur.Catalogue.ExcelFileGenerator
{
    public interface IExcelRow
    {
        IExcelCell CreateCell(int index);
    }

    internal class ExcelRow : IExcelRow
    {
        public ExcelRow(Row row, IExcelDocumentStyle documentStyle, IExcelSharedStrings excelSharedStrings)
        {
            this.row = row;
            this.documentStyle = documentStyle;
            this.excelSharedStrings = excelSharedStrings;
        }

        public IExcelCell CreateCell(int index)
        {
            var cell = new Cell
                {
                    CellReference = IndexHelpers.ToColumnName(index) + row.RowIndex
                };
            row.AppendChild(cell);
            return new ExcelCell(cell, documentStyle, excelSharedStrings);
        }

        private readonly Row row;
        private readonly IExcelDocumentStyle documentStyle;
        private readonly IExcelSharedStrings excelSharedStrings;
    }
}
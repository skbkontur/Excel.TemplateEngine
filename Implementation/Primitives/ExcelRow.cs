using DocumentFormat.OpenXml.Spreadsheet;

using SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Caches;
using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Primitives
{
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
                    CellReference = IndexHelpers.ToCellName(row.RowIndex, index)
                };
            row.AppendChild(cell);
            return new ExcelCell(cell, documentStyle, excelSharedStrings);
        }

        private readonly Row row;
        private readonly IExcelDocumentStyle documentStyle;
        private readonly IExcelSharedStrings excelSharedStrings;
    }
}
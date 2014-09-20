using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
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

        public IEnumerable<IExcelCell> Cells { get { return row.ChildElements.OfType<Cell>().Select(x => new ExcelCell(x, documentStyle, excelSharedStrings)); } }

        public void SetHeight(double value)
        {
            row.Height = value;
            row.CustomHeight = new BooleanValue(true);
        }

        private readonly Row row;
        private readonly IExcelDocumentStyle documentStyle;
        private readonly IExcelSharedStrings excelSharedStrings;
    }
}
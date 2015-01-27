using System;
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
                    CellReference = new ExcelCellIndex((int)row.RowIndex.Value, index).CellReference
                };
            row.AppendChild(cell);
            return new ExcelCell(cell, documentStyle, excelSharedStrings);
        }

        public IEnumerable<IExcelCell> Cells
        {
            get
            {
                var currentIndex = 1;
                foreach(var cell in row.ChildElements.OfType<Cell>())
                {
                    var index = GetCellXIndex(cell);
                    while(currentIndex < index)
                    {
                        yield return null;
                        currentIndex++;
                    }
                    yield return new ExcelCell(cell, documentStyle, excelSharedStrings);
                    currentIndex++;
                }
            }
        }

        public void SetHeight(double value)
        {
            row.Height = value;
            row.CustomHeight = new BooleanValue(true);
        }

        private static int GetCellXIndex(Cell cell)
        {
            return cell
                .CellReference.Value
                .TakeWhile(char.IsLetter)
                .Select(c => c - 'A' + 1)
                .Reverse()
                .Select((v, i) => v * (int)Math.Pow(26, i))
                .Sum();
        }

        private readonly Row row;
        private readonly IExcelDocumentStyle documentStyle;
        private readonly IExcelSharedStrings excelSharedStrings;
    }
}
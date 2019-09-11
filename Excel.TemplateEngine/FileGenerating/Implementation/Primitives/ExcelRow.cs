using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

using Excel.TemplateEngine.FileGenerating.Implementation.Caches;
using Excel.TemplateEngine.FileGenerating.Interfaces;

namespace Excel.TemplateEngine.FileGenerating.Implementation.Primitives
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
            var cellReference = new ExcelCellIndex((int)row.RowIndex.Value, index).CellReference;

            if (row.Elements<Cell>().Any(c => c.CellReference.Value == cellReference))
                return new ExcelCell(row.Elements<Cell>().First(c => c.CellReference.Value == cellReference), documentStyle, excelSharedStrings);

            var refCell = row.Elements<Cell>().FirstOrDefault(cell => IsCellReferenceGreaterThan(cell.CellReference.Value, cellReference));
            var newCell = new Cell
                {
                    CellReference = new ExcelCellIndex((int)row.RowIndex.Value, index).CellReference
                };

            row.InsertBefore(newCell, refCell);

            return new ExcelCell(newCell, documentStyle, excelSharedStrings);
        }

        private static bool IsCellReferenceGreaterThan(string first, string second)
        {
            if (first.Length == second.Length)
                return string.Compare(first, second, StringComparison.OrdinalIgnoreCase) > 0;
            return first.Length > second.Length;
        }

        public IEnumerable<IExcelCell> Cells
        {
            get
            {
                var currentIndex = 1;
                foreach (var cell in row.ChildElements.OfType<Cell>())
                {
                    var index = GetCellXIndex(cell);
                    while (currentIndex < index)
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
                   .CellReference.Value.TakeWhile(char.IsLetter)
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
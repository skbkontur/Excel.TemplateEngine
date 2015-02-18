using System.Collections.Generic;

using SKBKontur.Catalogue.ExcelFileGenerator.Implementation;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Interfaces
{
    public interface IExcelWorksheet
    {
        IExcelCell InsertCell(ExcelCellIndex cellIndex);
        IExcelRow CreateRow(int rowIndex);
        void MergeCells(ExcelCellIndex upperLeft, ExcelCellIndex lowerRight);
        void CreateAutofilter(ExcelCellIndex upperLeft, ExcelCellIndex lowerRight);
        void CreateHyperlink(ExcelCellIndex from, int toWorksheet, ExcelCellIndex to);
        void ResizeColumn(int columnIndex, double width);
        IEnumerable<IExcelCell> GetSortedCellsInRange(ExcelCellIndex upperLeft, ExcelCellIndex lowerRight);
        IExcelCell GetCell(ExcelCellIndex position);
        IEnumerable<IExcelCell> SearchCellsByText(string text);
        IEnumerable<IExcelRow> Rows { get; }
        IEnumerable<IExcelColumn> Columns { get; }
    }
}
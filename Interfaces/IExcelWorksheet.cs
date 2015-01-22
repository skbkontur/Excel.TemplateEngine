using System.Collections.Generic;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Interfaces
{
    public interface IExcelWorksheet
    {
        IExcelRow CreateRow(int rowIndex);
        void MergeCells(int fromRow, int fromCol, int toRow, int toCol);
        void CreateAutofilter(int fromRow, int fromCol, int toRow, int toCol);
        void CreateHyperlink(int row, int col, int toWorksheet, int toRow, int toCol);
        void ResizeColumn(int columnIndex, double width);
        IEnumerable<IExcelRow> Rows { get; }
    }
}
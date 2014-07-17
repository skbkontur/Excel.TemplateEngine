namespace SKBKontur.Catalogue.ExcelFileGenerator.Interfaces
{
    public interface IExcelSpreadsheet
    {
        IExcelRow CreateRow(int rowIndex);
        void MergeCells(int fromRow, int fromCol, int toRow, int toCol);
        void CreateAutofilter(int fromRow, int fromCol, int toRow, int toCol);
        void CreateHyperlink(int row, int col, int toSpreadsheet, int toRow, int toCol);
        void ResizeColumn(int columnIndex, double width);
    }
}
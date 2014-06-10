namespace SKBKontur.Catalogue.ExcelFileGenerator.Interfaces
{
    public interface IExcelSpreadsheet
    {
        void MergeCells(int fromRow, int fromCol, int toRow, int toCol);
        IExcelRow CreateRow(int rowIndex);
        void CreateHyperlink(int row, int col, int toSpreadsheet, int toRow, int toCol);
    }
}
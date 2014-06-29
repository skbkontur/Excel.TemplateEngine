using System;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Interfaces
{
    public interface IExcelDocument : IDisposable
    {
        byte[] GetDocumentBytes();
        IExcelSpreadsheet GetSpreadsheet(int index);
        void DeleteSpreadsheet(int index);
        void RenameSpreadSheet(int index, string name);
        void SetPivotTableSource(int tableIndex, int fromRow, int fromColumn, int toRow, int toColumn);
    }
}
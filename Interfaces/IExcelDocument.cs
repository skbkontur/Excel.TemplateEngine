using System;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Interfaces
{
    public interface IExcelDocument : IDisposable
    {
        byte[] GetDocumentBytes();
        IExcelWorksheet GetWorksheet(int index);
        void DeleteWorksheet(int index);
        void RenameWorksheet(int index, string name);
        void SetPivotTableSource(int tableIndex, int fromRow, int fromColumn, int toRow, int toColumn);
    }
}
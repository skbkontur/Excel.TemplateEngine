using System;

using SKBKontur.Catalogue.ExcelFileGenerator.Implementation;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Interfaces
{
    public interface IExcelDocument : IDisposable
    {
        byte[] GetDocumentBytes();
        int GetWorksheetCount();
        IExcelWorksheet GetWorksheet(int index);
        void DeleteWorksheet(int index);
        void RenameWorksheet(int index, string name);
        IExcelWorksheet AddWorksheet(string worksheetName);
        void SetPivotTableSource(int tableIndex, ExcelCellIndex upperLeft, ExcelCellIndex lowerRight);
    }
}
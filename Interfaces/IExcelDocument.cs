using System;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Interfaces
{
    public interface IExcelDocument : IDisposable
    {
        byte[] CloseAndGetDocumentBytes();
        int GetWorksheetCount();
        IExcelWorksheet GetWorksheet(int index);
        void RenameWorksheet(int index, string name);
        IExcelWorksheet AddWorksheet(string worksheetName);
        IExcelWorksheet FindWorksheet(string name);
        string GetWorksheetName(int index);
        IExcelVbaInfo GetVbaInfo();
        void AddVbaInfo(IExcelVbaInfo excelVbaInfo);
    }
}
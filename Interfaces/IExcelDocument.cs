using System;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Interfaces
{
    public interface IExcelDocument : IDisposable
    {
        byte[] GetDocumentBytes();
        IExcelSpreadsheet GetSpreadsheet(int index);
    }
}
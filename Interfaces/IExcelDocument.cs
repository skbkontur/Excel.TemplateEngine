using System;
using System.Collections.Generic;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Interfaces
{
    public interface IExcelDocument : IDisposable
    {
        byte[] CloseAndGetDocumentBytes();
        int GetWorksheetCount();
        IExcelWorksheet GetWorksheet(int index);
        void RenameWorksheet(int index, string name);
        IExcelWorksheet AddWorksheet(string worksheetName);
        List<IExcelFormControlInfo> GetFormControlInfos(int worksheetIndex); // todo (mpivko, 15.12.2017): maybe it's not the best place. Consider moving it to IExcelWorksheet
        void AddFormControlInfos(int worksheetIndex, IEnumerable<IExcelFormControlInfo> formControlInfos);
    }
}
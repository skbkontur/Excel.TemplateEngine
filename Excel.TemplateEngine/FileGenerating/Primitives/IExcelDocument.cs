#nullable enable

using System;

using JetBrains.Annotations;

namespace SkbKontur.Excel.TemplateEngine.FileGenerating.Primitives;

public interface IExcelDocument : IDisposable
{
    byte[] CloseAndGetDocumentBytes();

    int GetWorksheetCount();

    IExcelWorksheet? TryGetWorksheet(int index);

    IExcelWorksheet GetWorksheet(int index);

    void RenameWorksheet(int index, string name);

    IExcelWorksheet AddWorksheet(string worksheetName);

    IExcelWorksheet? FindWorksheet(string name);

    string? GetWorksheetName(int index);

    string? GetDescription();

    void AddDescription(string text);

    void CopyVbaInfoFrom(IExcelDocument excelDocument);

    [ContractAnnotation("=> true, value:notnull; => false, value:null")]
    bool TryGetCustomProperty(string key, out string? value);

    void SetCustomProperty(string key, string value);
}
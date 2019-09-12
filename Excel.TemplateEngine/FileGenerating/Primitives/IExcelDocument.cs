using System;

using JetBrains.Annotations;

namespace SkbKontur.Excel.TemplateEngine.FileGenerating.Primitives
{
    public interface IExcelDocument : IDisposable
    {
        [NotNull]
        byte[] CloseAndGetDocumentBytes();

        int GetWorksheetCount();

        [CanBeNull]
        IExcelWorksheet TryGetWorksheet(int index);

        [NotNull]
        IExcelWorksheet GetWorksheet(int index);

        void RenameWorksheet(int index, [NotNull] string name);

        [NotNull]
        IExcelWorksheet AddWorksheet([NotNull] string worksheetName);

        [CanBeNull]
        IExcelWorksheet FindWorksheet([NotNull] string name);

        [NotNull]
        string GetWorksheetName(int index);

        [CanBeNull]
        string GetDescription();

        void AddDescription([NotNull] string text);

        void CopyVbaInfoFrom([NotNull] IExcelDocument excelDocument);
    }
}
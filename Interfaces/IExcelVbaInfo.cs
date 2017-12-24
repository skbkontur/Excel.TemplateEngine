using DocumentFormat.OpenXml.Packaging;

using JetBrains.Annotations;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Interfaces
{
    public interface IExcelVbaInfo
    {
        [NotNull]
        VbaProjectPart VbaProjectPart { get; }
    }
}
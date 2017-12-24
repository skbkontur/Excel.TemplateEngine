using DocumentFormat.OpenXml.Packaging;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Primitives
{
    class ExcelVbaInfo : IExcelVbaInfo
    {
        public ExcelVbaInfo([NotNull] VbaProjectPart vbaProjectPart)
        {
            VbaProjectPart = vbaProjectPart;
        }

        [NotNull]
        public VbaProjectPart VbaProjectPart { get; }
    }
}
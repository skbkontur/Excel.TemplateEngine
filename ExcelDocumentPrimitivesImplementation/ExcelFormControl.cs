using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;
using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.ExcelDocumentPrimitivesImplementation
{
    public class ExcelFormControl : IFormControl
    {
        public ExcelFormControl([NotNull] IExcelFormControlInfo excelFormControlInfo)
        {
            ExcelFormControlInfo = excelFormControlInfo;
        }

        [NotNull]
        public IExcelFormControlInfo ExcelFormControlInfo { get; set; }
    }
}
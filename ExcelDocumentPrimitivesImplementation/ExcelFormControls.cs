using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;
using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.ExcelDocumentPrimitivesImplementation
{
    public class ExcelFormControls : IFormControls
    {
        public ExcelFormControls([NotNull] IExcelFormControlsInfo excelFormControlsInfo)
        {
            ExcelFormControlsInfo = excelFormControlsInfo;
        }

        [NotNull]
        public IExcelFormControlsInfo ExcelFormControlsInfo { get; set; }
    }
}
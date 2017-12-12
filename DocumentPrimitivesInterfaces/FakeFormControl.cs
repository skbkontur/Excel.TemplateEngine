using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces
{
    // todo (mpivko, 15.12.2017): what is it for?
    class FakeFormControl : IFormControl
    {
        public IExcelFormControlInfo ExcelFormControlInfo { get; set; }
    }
}
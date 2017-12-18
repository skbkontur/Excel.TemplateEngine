using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces
{
    public interface IFormControl
    {
        // todo (mpivko, 19.12.2017): what is it for?
        IExcelFormControlInfo ExcelFormControlInfo { get; }
    }
}
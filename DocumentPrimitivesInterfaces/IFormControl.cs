using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces
{
    public interface IFormControl
    {
        IExcelFormControlInfo ExcelFormControlInfo { get; }
    }
}
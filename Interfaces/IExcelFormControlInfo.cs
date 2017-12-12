using DocumentFormat.OpenXml.Packaging;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Interfaces
{
    public interface IExcelFormControlInfo
    {
        bool IsChecked { get; }
        VmlDrawingPart VmlDrawingPart { get; }
        string SelectedValue { get; }
    }
}
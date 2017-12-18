using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Interfaces
{
    public interface IExcelFormControlInfo
    {
        bool IsChecked { get; set; }
        Control Control { get; }
        DrawingsPart GlobalDrawingsPart { get; } // todo (mpivko, 19.12.2017): remove it, because it's global
        VmlDrawingPart GlobalVmlDrawingPart { get; } // todo (mpivko, 19.12.2017): remove it, because it's global
        ControlPropertiesPart ControlPropertiesPart { get; }
        string SelectedValue { get; set; }
    }
}
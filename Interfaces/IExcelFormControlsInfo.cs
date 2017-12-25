using System.Collections.Generic;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Interfaces
{
    public interface IExcelFormControlsInfo
    {
        (DrawingsPart part, string id) DrawingsPart { get; }
        (VmlDrawingPart part, string id) VmlDrawingPart { get; }
        List<(ControlPropertiesPart controlPropertiesPart, string id)> ControlPropertiesParts { get; }
        Controls Controls { get; }
    }
}
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Primitives
{
    public class ExcelFormControlsInfo : IExcelFormControlsInfo
    {
        public ExcelFormControlsInfo(WorksheetPart worksheetPart)
        {
            DrawingsPart = (worksheetPart.DrawingsPart, worksheetPart.DrawingsPart == null ? null : worksheetPart.GetIdOfPart(worksheetPart.DrawingsPart));
            var vmlPart = worksheetPart.VmlDrawingParts.SingleOrDefault(); // todo (mpivko, 25.12.2017): what if not single?
            VmlDrawingPart = (vmlPart, vmlPart == null ? null : worksheetPart.GetIdOfPart(vmlPart));
            ControlPropertiesParts = worksheetPart.ControlPropertiesParts.Select(x => (x, x == null ? null : worksheetPart.GetIdOfPart(x))).ToList();
            Controls = worksheetPart.Worksheet?.GetFirstChild<AlternateContent>()?.GetFirstChild<AlternateContentChoice>()?.GetFirstChild<Controls>();
        }

        public (DrawingsPart part, string id) DrawingsPart { get; }
        public (VmlDrawingPart part, string id) VmlDrawingPart { get; }
        public List<(ControlPropertiesPart controlPropertiesPart, string id)> ControlPropertiesParts { get; }
        public Controls Controls { get; }
    }
}
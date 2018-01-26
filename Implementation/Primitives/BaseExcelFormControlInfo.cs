using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Primitives
{
    public class BaseExcelFormControlInfo
    {
        public BaseExcelFormControlInfo([NotNull] IExcelWorksheet excelWorksheet, [NotNull] Control control, [NotNull] ControlPropertiesPart controlPropertiesPart, [NotNull] VmlDrawingPart vmlDrawingPart, [NotNull] DrawingsPart drawingsPart)
        {
            this.excelWorksheet = excelWorksheet;
            Control = control;
            GlobalDrawingsPart = drawingsPart;
            GlobalVmlDrawingPart = vmlDrawingPart;
            ControlPropertiesPart = controlPropertiesPart;
        }

        public DrawingsPart GlobalDrawingsPart { get; }
        public VmlDrawingPart GlobalVmlDrawingPart { get; }

        [NotNull]
        protected readonly IExcelWorksheet excelWorksheet;

        [NotNull]
        public Control Control { get; }

        [NotNull]
        public ControlPropertiesPart ControlPropertiesPart { get; }
    }
}
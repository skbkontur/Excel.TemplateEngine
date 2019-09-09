using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Primitives
{
    public class BaseExcelFormControlInfo : IExcelFormControlInfo
    {
        protected BaseExcelFormControlInfo([NotNull] IExcelWorksheet excelWorksheet, [NotNull] Control control, [NotNull] ControlPropertiesPart controlPropertiesPart, [NotNull] VmlDrawingPart vmlDrawingPart)
        {
            this.excelWorksheet = excelWorksheet;
            Control = control;
            GlobalVmlDrawingPart = vmlDrawingPart;
            ControlPropertiesPart = controlPropertiesPart;
        }

        [NotNull]
        protected VmlDrawingPart GlobalVmlDrawingPart { get; }

        [NotNull]
        protected Control Control { get; }

        [NotNull]
        protected ControlPropertiesPart ControlPropertiesPart { get; }

        [NotNull]
        protected readonly IExcelWorksheet excelWorksheet;
    }
}
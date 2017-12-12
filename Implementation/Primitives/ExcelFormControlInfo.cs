using System;
using System.Linq;

using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Primitives
{
    public class ExcelFormControlInfo : IExcelFormControlInfo
    {
        public ExcelFormControlInfo([NotNull] IExcelWorksheet excelWorksheet, [NotNull] VmlDrawingPart vmlDrawingPart, [NotNull] ControlPropertiesPart controlPropertiesPart)
        {
            this.excelWorksheet = excelWorksheet;
            this.VmlDrawingPart = vmlDrawingPart;
            this.controlPropertiesPart = controlPropertiesPart;
        }

        public bool IsChecked => controlPropertiesPart.FormControlProperties.Checked != null && controlPropertiesPart.FormControlProperties.Checked.HasValue && controlPropertiesPart.FormControlProperties.Checked.Value == CheckedValues.Checked;

        public string SelectedValue
        {
            get
            {
                // todo (mpivko, 18.12.2017): it does not work when range is from other worksheet
                var range = controlPropertiesPart.FormControlProperties.FmlaRange.Value;
                var parts = range.Split(':').Select(x => x.Replace("$", "")).ToList();
                if (parts.Count != 2)
                    throw new Exception($"Invalid range: '{range}'");
                return excelWorksheet.GetSortedCellsInRange(new ExcelCellIndex(parts[0]), new ExcelCellIndex(parts[1])).Skip((int)controlPropertiesPart.FormControlProperties.Selected.Value - 1).First().GetStringValue();
            }
        }

        public VmlDrawingPart VmlDrawingPart { get; }

        [NotNull]
        private readonly IExcelWorksheet excelWorksheet;
        private readonly ControlPropertiesPart controlPropertiesPart;
    }
}
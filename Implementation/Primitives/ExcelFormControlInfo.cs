using System;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Primitives
{
    public class ExcelFormControlInfo : IExcelFormControlInfo
    {
        public ExcelFormControlInfo([NotNull] IExcelWorksheet excelWorksheet, [NotNull] Control control, [NotNull] ControlPropertiesPart controlPropertiesPart, [NotNull] VmlDrawingPart vmlDrawingPart, [NotNull] DrawingsPart drawingsPart)
        {
            this.excelWorksheet = excelWorksheet;
            this.Control = control;
            this.GlobalDrawingsPart = drawingsPart;
            this.GlobalVmlDrawingPart = vmlDrawingPart;
            this.ControlPropertiesPart = controlPropertiesPart;
        }

        public bool IsChecked { get => ControlPropertiesPart.FormControlProperties.Checked != null && ControlPropertiesPart.FormControlProperties.Checked.HasValue && ControlPropertiesPart.FormControlProperties.Checked.Value == CheckedValues.Checked;
            set {
                ControlPropertiesPart.FormControlProperties.Checked = value ? CheckedValues.Checked : CheckedValues.Unchecked;
                lock(GlobalVmlDrawingPart)
                {
                    var ns = "urn:schemas-microsoft-com:office:excel";
                    var xdoc = XDocument.Load(GlobalVmlDrawingPart.GetStream());
                    var clientData = xdoc.Root.Elements().Single(x => x.Attribute("id")?.Value == Control.Name).Element(XName.Get("ClientData", ns));
                    var checkedElement = clientData.Element(XName.Get("Checked", ns));
                    checkedElement?.Remove();
                    if (value)
                        clientData.Add(new XElement(XName.Get("Checked", ns), "1"));
                    xdoc.Save(GlobalVmlDrawingPart.GetStream());
                }
            }
            
        }

        public string SelectedValue
        {
            get
            {
                // todo (mpivko, 18.12.2017): it does not work when range is from other worksheet
                var range = ControlPropertiesPart.FormControlProperties.FmlaRange.Value;
                var parts = range.Split(':').Select(x => x.Replace("$", "")).ToList();
                if (parts.Count != 2)
                    throw new Exception($"Invalid range: '{range}'");
                return excelWorksheet.GetSortedCellsInRange(new ExcelCellIndex(parts[0]), new ExcelCellIndex(parts[1])).Skip((int)ControlPropertiesPart.FormControlProperties.Selected.Value - 1).First().GetStringValue();
            }
            set
            {
                // todo (mpivko, 18.12.2017): it does not work when range is from other worksheet
                if (ControlPropertiesPart.FormControlProperties.FmlaRange == null)
                    throw new ArgumentException("This form control has no FmlaRange (maybe you are using it as dropdown, while it isn't dropdown)");
                var range = ControlPropertiesPart.FormControlProperties.FmlaRange.Value;
                var parts = range.Split(':').Select(x => x.Replace("$", "")).ToList();
                if (parts.Count != 2)
                    throw new Exception($"Invalid range: '{range}'");
                var index = excelWorksheet.GetSortedCellsInRange(new ExcelCellIndex(parts[0]), new ExcelCellIndex(parts[1])).Select(x => x.GetStringValue()).ToList().IndexOf(value);
                ControlPropertiesPart.FormControlProperties.Selected = (uint)index + 1;

                var ns = "urn:schemas-microsoft-com:office:excel";
                var xdoc = XDocument.Load(GlobalVmlDrawingPart.GetStream());
                var clientData = xdoc.Root.Elements().Single(x => x.Attribute("id")?.Value == Control.Name).Element(XName.Get("ClientData", ns));
                var checkedElement = clientData.Element(XName.Get("Sel", ns));
                checkedElement?.Remove();
                clientData.Add(new XElement(XName.Get("Sel", ns), (index + 1).ToString()));
                xdoc.Save(GlobalVmlDrawingPart.GetStream());
            }
        }

        public DrawingsPart GlobalDrawingsPart { get; }
        public VmlDrawingPart GlobalVmlDrawingPart { get; }

        [NotNull]
        private readonly IExcelWorksheet excelWorksheet;

        [NotNull]
        public Control Control { get; }

        public ControlPropertiesPart ControlPropertiesPart { get; }
    }
}
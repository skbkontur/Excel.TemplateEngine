using System;
using System.Collections.Generic;
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
    // todo (mpivko, 22.12.2017): use different classes for different form controls
    public class ExcelFormControlInfo : IExcelCheckBoxControlInfo, IExcelDropDownControlInfo
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
            set
            {
                if(value)
                    ControlPropertiesPart.FormControlProperties.Checked = CheckedValues.Checked;
                else
                    ControlPropertiesPart.FormControlProperties.Checked = null;
                lock (GlobalVmlDrawingPart)
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
                if(!ControlPropertiesPart.FormControlProperties.Selected.HasValue)
                    return null;
                return GetDropDownCells().Skip((int)ControlPropertiesPart.FormControlProperties.Selected.Value - 1).First().GetStringValue();
            }
            set
            {
                var index = GetDropDownCells().Select(x => x.GetStringValue()).ToList().IndexOf(value);
                if(index == -1)
                    return; // todo (mpivko, 21.12.2017): unknown value
                ControlPropertiesPart.FormControlProperties.Selected = (uint)index + 1;

                var ns = "urn:schemas-microsoft-com:office:excel";
                var xdoc = XDocument.Load(GlobalVmlDrawingPart.GetStream());
                var clientData = xdoc.Root.Elements().Single(x => x.Attribute("id")?.Value == Control.Name).Element(XName.Get("ClientData", ns));
                var checkedElement = clientData.Element(XName.Get("Sel", ns)); // todo (mpivko, 21.12.2017): do it more carefully
                checkedElement?.Remove();
                clientData.Add(new XElement(XName.Get("Sel", ns), (index + 1).ToString()));
                xdoc.Save(GlobalVmlDrawingPart.GetStream());
            }
        }

        private IEnumerable<IExcelCell> GetDropDownCells()
        {
            // todo (mpivko, 18.12.2017): it does not work when range is from other worksheet
            if (ControlPropertiesPart.FormControlProperties.FmlaRange == null)
                throw new ArgumentException("This form control has no FmlaRange (maybe you are using it as dropdown, while it isn't dropdown)");
            var absoluteRange = ControlPropertiesPart.FormControlProperties.FmlaRange.Value;
            var (worksheetName, range) = SplitAbsoluteRange(absoluteRange);
            var parts = range.Split(':').Select(x => x.Replace("$", "")).ToList(); // todo (mpivko, 21.12.2017): extract method
            if (parts.Count != 2)
                throw new Exception($"Invalid range: '{range}'"); // todo (mpivko, 22.12.2017): use other exception
            var worksheet = worksheetName == null ? excelWorksheet : excelWorksheet.ExcelDocument.FindWorksheet(worksheetName);
            if (worksheet == null)
                throw new Exception($"Worksheet with name {worksheetName} not found, but used in dropDown"); // todo (mpivko, 22.12.2017): use other exception
            return worksheet.GetSortedCellsInRange(new ExcelCellIndex(parts[0]), new ExcelCellIndex(parts[1]));
        }

        private (string worksheetName, string cellRange) SplitAbsoluteRange(string absoluteRange)
        {
            var parts = absoluteRange.Split('!').ToList();
            if(parts.Count == 1)
                return (null, parts[0]);
            if(parts.Count == 2)
            {
                if (parts[0].StartsWith("'") && parts[0].EndsWith("'"))
                    return (parts[0].Substring(1, parts[0].Length - 2), parts[1]);
                return (parts[0], parts[1]);
            }
            throw new ArgumentException($"Invalid range: '{absoluteRange}'");
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
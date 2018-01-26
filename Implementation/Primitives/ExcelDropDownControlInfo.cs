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

using SKBKontur.Catalogue.ExcelFileGenerator.Exceptions;
using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Primitives
{
    public class ExcelDropDownControlInfo : BaseExcelFormControlInfo, IExcelDropDownControlInfo
    {
        public ExcelDropDownControlInfo([NotNull] IExcelWorksheet excelWorksheet, [NotNull] Control control, [NotNull] ControlPropertiesPart controlPropertiesPart, [NotNull] VmlDrawingPart vmlDrawingPart, [NotNull] DrawingsPart drawingsPart)
            : base(excelWorksheet, control, controlPropertiesPart, vmlDrawingPart, drawingsPart)
        {
        }

        public string SelectedValue
        {
            get
            {
                if (ControlPropertiesPart.FormControlProperties?.Selected == null || !ControlPropertiesPart.FormControlProperties.Selected.HasValue)
                    return null;
                return GetDropDownCells().ElementAt((int)ControlPropertiesPart.FormControlProperties.Selected.Value - 1).GetStringValue();
            }
            set
            {

                var index = GetDropDownCells().Select(x => x.GetStringValue()).ToList().IndexOf(value);
                if (index == -1)
                    throw new ArgumentException($"Unable to set dropbox value: value '{value}' is unknown");
                if(ControlPropertiesPart.FormControlProperties == null)
                    ControlPropertiesPart.FormControlProperties = new FormControlProperties();
                ControlPropertiesPart.FormControlProperties.Selected = (uint)index + 1;
                var ns = "urn:schemas-microsoft-com:office:excel";
                var xdoc = XDocument.Load(GlobalVmlDrawingPart.GetStream());
                var clientData = xdoc.Root?.Elements()?.Single(x => x.Attribute("id")?.Value == Control.Name)?.Element(XName.Get("ClientData", ns));
                if (clientData == null)
                    throw new InvalidExcelDocumentException($"ClientData element is not found for control with name '{Control.Name}'");
                var checkedElement = clientData.Element(XName.Get("Sel", ns));
                checkedElement?.Remove();
                clientData.Add(new XElement(XName.Get("Sel", ns), (index + 1).ToString()));
                xdoc.Save(GlobalVmlDrawingPart.GetStream());
            }
        }

        private IEnumerable<IExcelCell> GetDropDownCells()
        {
            if (ControlPropertiesPart.FormControlProperties?.FmlaRange?.Value == null)
                throw new ArgumentException("This form control has no FmlaRange (maybe you are using it as dropdown, while it isn't dropdown)");
            var absoluteRange = ControlPropertiesPart.FormControlProperties.FmlaRange.Value;
            var (worksheetName, relativeRange) = SplitAbsoluteRange(absoluteRange);
            var range = ParseRelativeRange(relativeRange);
            var worksheet = worksheetName == null ? excelWorksheet : excelWorksheet.ExcelDocument.FindWorksheet(worksheetName);
            if (worksheet == null)
                throw new InvalidExcelDocumentException($"Worksheet with name {worksheetName} not found, but used in dropDown");
            return worksheet.GetSortedCellsInRange(range.from, range.to);
        }

        private (string worksheetName, string relativeRange) SplitAbsoluteRange([NotNull] string absoluteRange)
        {
            var parts = absoluteRange.Split('!').ToList();
            if (parts.Count == 1)
                return (null, parts[0]);
            if (parts.Count == 2)
            {
                if (parts[0].StartsWith("'") && parts[0].EndsWith("'"))
                    return (parts[0].Substring(1, parts[0].Length - 2), parts[1]);
                return (parts[0], parts[1]);
            }
            throw new InvalidExcelDocumentException($"Invalid absolute range: '{absoluteRange}'");
        }

        private (ExcelCellIndex from, ExcelCellIndex to) ParseRelativeRange([NotNull] string relativeRange)
        {
            var parts = relativeRange.Split(':').Select(x => x.Replace("$", "")).ToList();
            if (parts.Count != 2)
                throw new InvalidExcelDocumentException($"Invalid relative range: '{relativeRange}'");
            return (new ExcelCellIndex(parts[0]), new ExcelCellIndex(parts[1]));

        }
    }
}
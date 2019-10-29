using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using JetBrains.Annotations;

using SkbKontur.Excel.TemplateEngine.FileGenerating.DataTypes;

using Vostok.Logging.Abstractions;

namespace SkbKontur.Excel.TemplateEngine.FileGenerating.Primitives.Implementations
{
    public class ExcelDropDownControlInfo : BaseExcelFormControlInfo, IExcelDropDownControlInfo
    {
        public ExcelDropDownControlInfo([NotNull] IExcelWorksheet excelWorksheet, [NotNull] Control control, [NotNull] ControlPropertiesPart controlPropertiesPart, [NotNull] VmlDrawingPart vmlDrawingPart, [NotNull] ILog logger)
            : base(excelWorksheet, control, controlPropertiesPart, vmlDrawingPart)
        {
            this.logger = logger;
        }

        [CanBeNull]
        public string SelectedValue
        {
            get
            {
                if (ControlPropertiesPart.FormControlProperties?.Selected == null || !ControlPropertiesPart.FormControlProperties.Selected.HasValue)
                    return null;
                var cells = GetDropDownCells().ToList();
                var index = (int)ControlPropertiesPart.FormControlProperties.Selected.Value - 1;
                if (index < 0 || index >= cells.Count)
                    return null;
                return cells.ElementAt(index).GetStringValue();
            }
            set
            {
                var index = GetDropDownCells().Select(x => x.GetStringValue()).ToList().IndexOf(value);
                if (index == -1)
                    logger.Error($"Tried to set unknown dropbox value: '{value}'. Setting empty value instead");

                if (ControlPropertiesPart.FormControlProperties == null)
                    ControlPropertiesPart.FormControlProperties = new FormControlProperties();
                ControlPropertiesPart.FormControlProperties.Selected = (uint)index + 1;
                const string ns = "urn:schemas-microsoft-com:office:excel";
                lock (GlobalVmlDrawingPart)
                {
                    XDocument xdoc;
                    using (var stream = GlobalVmlDrawingPart.GetStream())
                        xdoc = XDocument.Load(stream);
                    // ReSharper disable once ConstantConditionalAccessQualifier
                    var clientData = xdoc.Root?.Elements()?.Single(x => x.Attribute("id")?.Value == Control.Name)?.Element(XName.Get("ClientData", ns));
                    if (clientData == null)
                        throw new InvalidOperationException($"ClientData element is not found for control with name '{Control.Name}'");
                    var checkedElement = clientData.Element(XName.Get("Sel", ns));
                    checkedElement?.Remove();
                    clientData.Add(new XElement(XName.Get("Sel", ns), (index + 1).ToString()));
                    using (var stream = GlobalVmlDrawingPart.GetStream())
                        xdoc.Save(stream);
                }
            }
        }

        [NotNull, ItemNotNull]
        private IEnumerable<IExcelCell> GetDropDownCells()
        {
            var absoluteRange = ControlPropertiesPart.FormControlProperties?.FmlaRange?.Value;
            if (absoluteRange == null)
                throw new InvalidOperationException("This form control has no FmlaRange (maybe you are using it as dropdown, while it isn't dropdown)");
            var (worksheetName, relativeRange) = SplitAbsoluteRange(absoluteRange);
            var (from, to) = ParseRelativeRange(relativeRange);
            var worksheet = worksheetName == null ? excelWorksheet : excelWorksheet.ExcelDocument.FindWorksheet(worksheetName);
            if (worksheet == null)
                throw new InvalidOperationException($"Worksheet with name {worksheetName} not found, but used in dropDown");
            return worksheet.GetSortedCellsInRange(from, to);
        }

        private static (string worksheetName, string relativeRange) SplitAbsoluteRange([NotNull] string absoluteRange)
        {
            var parts = absoluteRange.Split('!').ToList();
            if (parts.Count == 1)
                return (null, parts[0]);
            if (parts.Count == 2)
            {
                var worksheetName = parts[0];
                if (worksheetName.StartsWith("'") && worksheetName.EndsWith("'"))
                    return (worksheetName.Substring(1, worksheetName.Length - 2), parts[1]);
                return (worksheetName, parts[1]);
            }
            throw new InvalidOperationException($"Invalid absolute range: '{absoluteRange}'");
        }

        private static (ExcelCellIndex from, ExcelCellIndex to) ParseRelativeRange([NotNull] string relativeRange)
        {
            var parts = relativeRange.Split(':').Select(x => x.Replace("$", "")).ToList();
            if (parts.Count != 2)
                throw new InvalidOperationException($"Invalid relative range: '{relativeRange}'");
            return (new ExcelCellIndex(parts[0]), new ExcelCellIndex(parts[1]));
        }

        private readonly ILog logger;
    }
}
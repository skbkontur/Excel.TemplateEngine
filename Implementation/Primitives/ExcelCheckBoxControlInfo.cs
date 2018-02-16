using System.Linq;
using System.Xml.Linq;

using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelFileGenerator.Exceptions;
using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Primitives
{
    public class ExcelCheckBoxControlInfo : BaseExcelFormControlInfo, IExcelCheckBoxControlInfo
    {
        public ExcelCheckBoxControlInfo([NotNull] IExcelWorksheet excelWorksheet, [NotNull] Control control, [NotNull] ControlPropertiesPart controlPropertiesPart, [NotNull] VmlDrawingPart vmlDrawingPart, [NotNull] DrawingsPart drawingsPart)
            : base(excelWorksheet, control, controlPropertiesPart, vmlDrawingPart, drawingsPart)
        {
        }

        public bool IsChecked
        {
            get => ControlPropertiesPart.FormControlProperties?.Checked != null && ControlPropertiesPart.FormControlProperties.Checked.HasValue && ControlPropertiesPart.FormControlProperties.Checked.Value == CheckedValues.Checked;
            set
            {
                if(ControlPropertiesPart.FormControlProperties == null)
                    ControlPropertiesPart.FormControlProperties = new FormControlProperties();
                if(value)
                    ControlPropertiesPart.FormControlProperties.Checked = CheckedValues.Checked;
                else
                    ControlPropertiesPart.FormControlProperties.Checked = null;
                lock(GlobalVmlDrawingPart)
                {
                    var ns = "urn:schemas-microsoft-com:office:excel";
                    var xdoc = XDocument.Load(GlobalVmlDrawingPart.GetStream());
                    var clientData = xdoc.Root?.Elements()?.Single(x => x.Attribute("id")?.Value == Control.Name)?.Element(XName.Get("ClientData", ns));
                    if(clientData == null)
                        throw new InvalidExcelDocumentException($"ClientData element is not found for control with name '{Control.Name}'");
                    var checkedElement = clientData.Element(XName.Get("Checked", ns));
                    checkedElement?.Remove();
                    if(value)
                        clientData.Add(new XElement(XName.Get("Checked", ns), "1"));
                    xdoc.Save(GlobalVmlDrawingPart.GetStream());
                }
            }
        }
    }
}
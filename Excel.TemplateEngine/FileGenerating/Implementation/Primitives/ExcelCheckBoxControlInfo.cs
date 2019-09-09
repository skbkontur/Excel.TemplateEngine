using System.IO;
using System.Linq;
using System.Xml.Linq;

using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using Excel.TemplateEngine.FileGenerating.Interfaces;

using JetBrains.Annotations;

namespace Excel.TemplateEngine.FileGenerating.Implementation.Primitives
{
    public class ExcelCheckBoxControlInfo : BaseExcelFormControlInfo, IExcelCheckBoxControlInfo
    {
        public ExcelCheckBoxControlInfo([NotNull] IExcelWorksheet excelWorksheet, [NotNull] Control control, [NotNull] ControlPropertiesPart controlPropertiesPart, [NotNull] VmlDrawingPart vmlDrawingPart)
            : base(excelWorksheet, control, controlPropertiesPart, vmlDrawingPart)
        {
        }

        public bool IsChecked
        {
            get => ControlPropertiesPart.FormControlProperties?.Checked != null && ControlPropertiesPart.FormControlProperties.Checked.HasValue && ControlPropertiesPart.FormControlProperties.Checked.Value == CheckedValues.Checked;
            set
            {
                if (ControlPropertiesPart.FormControlProperties == null)
                    ControlPropertiesPart.FormControlProperties = new FormControlProperties();
                if (value)
                    ControlPropertiesPart.FormControlProperties.Checked = CheckedValues.Checked;
                else
                    ControlPropertiesPart.FormControlProperties.Checked = null;
                lock (GlobalVmlDrawingPart)
                {
                    const string ns = "urn:schemas-microsoft-com:office:excel";
                    var xdoc = XDocument.Load((Stream)GlobalVmlDrawingPart.GetStream());
                    // ReSharper disable once ConstantConditionalAccessQualifier
                    var clientData = xdoc.Root?.Elements()?.SingleOrDefault(x => x.Attribute("id")?.Value == Control.Name)?.Element(XName.Get("ClientData", ns));
                    if (clientData == null)
                        throw new InvalidProgramStateException($"ClientData element is not found for control with name '{Control.Name}'");
                    var checkedElement = clientData.Element(XName.Get("Checked", ns));
                    checkedElement?.Remove();
                    if (value)
                        clientData.Add(new XElement(XName.Get("Checked", ns), "1"));
                    xdoc.Save((Stream)GlobalVmlDrawingPart.GetStream());
                }
            }
        }
    }
}
using System.Linq;

using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;

using JetBrains.Annotations;

namespace SkbKontur.Excel.TemplateEngine.FileGenerating.Helpers;

public static class CustomFilePropertiesPartHelper
{
    [CanBeNull]
    public static CustomDocumentProperty GetProperty(this CustomFilePropertiesPart customFilePropertiesPart, string key)
    {
        return customFilePropertiesPart.Properties?.OfType<CustomDocumentProperty>()
                                       .FirstOrDefault(i => i?.Name?.Value == key);
    }
}
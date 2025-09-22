using DocumentFormat.OpenXml.Spreadsheet;

namespace SkbKontur.Excel.TemplateEngine.FileGenerating.Helpers;

internal static class DefinedNameHelper
{
    public static bool IsBelongsToSheet(this DefinedName definedName, Sheet sheet)
    {
        var definedNameInnerText = definedName.InnerText;
        var definedNameSheetName = definedName.InnerText.Substring(0, definedNameInnerText.LastIndexOf('!'));

        return definedNameSheetName == sheet.Name;
    }

    public static DefinedName UpdateLinkSheetName(this DefinedName definedName, string sheetName)
    {
        var exclamationIndex = definedName.InnerText.LastIndexOf('!');
        var cellAddress = definedName.InnerText.Substring(exclamationIndex + 1);
        var newDefinedNameInnerText = $"{sheetName}!{cellAddress}";

        return new DefinedName(newDefinedNameInnerText)
            {
                Name = definedName.Name,
            };
    }
}
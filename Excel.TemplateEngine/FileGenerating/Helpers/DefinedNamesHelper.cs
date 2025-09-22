#nullable enable

using DocumentFormat.OpenXml.Spreadsheet;

namespace SkbKontur.Excel.TemplateEngine.FileGenerating.Helpers;

internal static class DefinedNamesHelper
{
    public static DefinedNames UpdateAfterRenamingSheet(this DefinedNames definedNames,
                                                        Sheet sheetBeforeRenaming,
                                                        string newSheetName)
    {
        var updatedDefinedNames = new DefinedNames();
        foreach (var definedName in definedNames.Elements<DefinedName>())
        {
            if (definedName.IsBelongsToSheet(sheetBeforeRenaming))
            {
                var updatedDefinedName = definedName.UpdateLinkSheetName(newSheetName);
                updatedDefinedNames.AppendChild(updatedDefinedName.CloneNode(true));
            }
            else
                updatedDefinedNames.AppendChild(definedName.CloneNode(true));
        }

        return updatedDefinedNames;
    }
}
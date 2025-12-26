#nullable enable

using DocumentFormat.OpenXml.Spreadsheet;

namespace SkbKontur.Excel.TemplateEngine.FileGenerating.Helpers;

internal static class DefinedNamesHelper
{
    public static DefinedNames UpdateAfterRenamingSheet(this DefinedNames definedNames,
                                                        Sheet sheetBeforeRenaming,
                                                        string newName)
    {
        var updatedDefinedNames = new DefinedNames();
        foreach (var definedName in definedNames.Elements<DefinedName>())
        {
            if (definedName.BelongsToSheet(sheetBeforeRenaming))
            {
                var updatedDefinedName = definedName.ReplaceSheetNameInFormula(newName);
                updatedDefinedNames.AppendChild(updatedDefinedName.CloneNode(deep : true));
            }
            else
                updatedDefinedNames.AppendChild(definedName.CloneNode(deep : true));
        }

        return updatedDefinedNames;
    }
}
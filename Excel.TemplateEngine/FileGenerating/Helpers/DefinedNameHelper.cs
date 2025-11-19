using System;
using System.Text.RegularExpressions;

using DocumentFormat.OpenXml.Spreadsheet;

namespace SkbKontur.Excel.TemplateEngine.FileGenerating.Helpers;

internal static class DefinedNameHelper
{
    public static bool BelongsToSheet(this DefinedName definedName, Sheet sheet)
    {
        if (string.IsNullOrEmpty(definedName.InnerText))
            return false;

        var separatorIndex = definedName.InnerText.LastIndexOf(referencePartSeparator, StringComparison.Ordinal);
        var sheetNamePart = definedName.InnerText.Substring(0, separatorIndex);

        var escapedSheetName = sheet.Name?.Value?.Replace("'", "''");
        var isBelong = sheetNamePart.Contains($"'{escapedSheetName}'") || sheetNamePart.Contains($"{escapedSheetName}");

        return isBelong;
    }

    public static DefinedName ReplaceSheetNameInFormula(this DefinedName definedName, string newSheetName)
    {
        if (string.IsNullOrEmpty(newSheetName))
            throw new InvalidOperationException("Sheet name cannot be null or empty");

        var newSheetNamePart = FormatWorksheetNameForReference(newSheetName);

        var separatorIndex = definedName.InnerText.LastIndexOf(referencePartSeparator, StringComparison.Ordinal);
        var cellAddress = definedName.InnerText.Substring(separatorIndex + 1);

        var newInnerText = $"{newSheetNamePart}{referencePartSeparator}{cellAddress}";

        return new DefinedName(newInnerText) {Name = definedName.Name};
    }

    private static string FormatWorksheetNameForReference(string name)
    {
        if (NeedsSingleQuotes(name))
        {
            var escapedName = name.Replace("'", "''");
            return $"'{escapedName}'";
        }

        return name;
    }

    private static bool NeedsSingleQuotes(string name)
    {
        if (name.Contains(" "))
            return true;

        if (char.IsDigit(name[0]))
            return true;

        // Специальный символ (кроме нижнего подчеркивания)
        const string specialCharacterPattern = @"[^\p{L}\p{N}\s_]";
        if (Regex.IsMatch(name, specialCharacterPattern))
            return true;

        // A1-нотация: сперва английские буквы, затем цифры
        const string a1NotationPattern = @"^[A-Za-z]+[\d]+$";
        if (Regex.IsMatch(name, a1NotationPattern))
            return true;

        return false;
    }

    // Символ-разделитель в ссылках, отделяющий имя листа от диапазона ячеек (например, 'Рабочий лист'!B2:C10)
    private const string referencePartSeparator = "!";
}
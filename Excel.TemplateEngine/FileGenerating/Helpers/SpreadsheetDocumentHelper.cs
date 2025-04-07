using System;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using Vostok.Logging.Abstractions;

namespace SkbKontur.Excel.TemplateEngine.FileGenerating.Helpers;

internal static class SpreadsheetDocumentHelper
{
    public static Stylesheet GetOrCreateSpreadsheetStyles(this SpreadsheetDocument document, ILog logger)
    {
        var stylesPart = document.WorkbookPart!.WorkbookStylesPart ?? document.WorkbookPart.AddNewPart<WorkbookStylesPart>();

        Stylesheet stylesheet;
        try
        {
            stylesheet = stylesPart.Stylesheet ?? (stylesPart.Stylesheet = new Stylesheet());
        }
        catch (Exception ex)
        {
            logger.Warn(ex, "Can't parse styles of document");
            stylesheet = stylesPart.Stylesheet = new Stylesheet();
        }
        return stylesheet;
    }

    public static SharedStringTable GetOrCreateSpreadsheetSharedStrings(this SpreadsheetDocument document)
    {
        var sharedStringsPart = document.WorkbookPart!.SharedStringTablePart ?? document.WorkbookPart.AddNewPart<SharedStringTablePart>();
        return sharedStringsPart.SharedStringTable ?? (sharedStringsPart.SharedStringTable = new SharedStringTable());
    }
}
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Excel.TemplateEngine.Helpers
{
    internal static class SpreadsheetDocumentHelper
    {
        public static Stylesheet GetOrCreateSpreadsheetStyles(this SpreadsheetDocument document)
        {
            var stylesPart = document.WorkbookPart.WorkbookStylesPart ?? document.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            return stylesPart.Stylesheet ?? (stylesPart.Stylesheet = new Stylesheet());
        }

        public static SharedStringTable GetOrCreateSpreadsheetSharedStrings(this SpreadsheetDocument document)
        {
            var sharedStringsPart = document.WorkbookPart.SharedStringTablePart ?? document.WorkbookPart.AddNewPart<SharedStringTablePart>();
            return sharedStringsPart.SharedStringTable ?? (sharedStringsPart.SharedStringTable = new SharedStringTable());
        }
    }
}
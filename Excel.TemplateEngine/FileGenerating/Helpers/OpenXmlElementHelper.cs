using System;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SkbKontur.Excel.TemplateEngine.FileGenerating.Helpers
{
    public static class OpenXmlElementHelper
    {
        // todo (a.dobrynin, 05.08.2020): replace this method with targetWorksheet.AddChild(child) when DocumentFormat.OpenXml 2.12.0 is released
        // https://github.com/OfficeDev/Open-XML-SDK/pull/774
        /// <summary>
        ///     Adds node in the correct location according to the schema
        ///     http://msdn.microsoft.com/en-us/library/office/cc880096(v=office.15).aspx
        ///     https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.worksheet?view=openxml-2.8.1#definition
        ///     https://github.com/OfficeDev/Open-XML-SDK/blob/058ec42001ca97850fd82cc16e3b234c155a6e7e/src/DocumentFormat.OpenXml/GeneratedCode/schemas_microsoft_com_office_excel_2006_main.g.cs#L90
        /// </summary>
        public static T AddChildToCorrectLocation<T>(this OpenXmlCompositeElement parent, T child) where T : OpenXmlElement
        {
            if (!(parent is Worksheet))
                throw new InvalidOperationException("Schema-oriented insertion is supported only for Worksheets");

            var precedingElementTypes = openXmlWorksheetNodesOrder.TakeWhile(x => x != typeof(T));
            var closestPrecedingSibling = precedingElementTypes.Reverse()
                                                               .Select(precedingElementType => parent.Elements().LastOrDefault(x => x.GetType() == precedingElementType))
                                                               .FirstOrDefault(existingElement => existingElement != null);

            return closestPrecedingSibling != null
                       ? parent.InsertAfter(child, closestPrecedingSibling)
                       : parent.InsertAt(child, 0);
        }

        private static readonly Type[] openXmlWorksheetNodesOrder =
            {
                typeof(SheetProperties),
                typeof(SheetDimension),
                typeof(SheetViews),
                typeof(SheetFormatProperties),
                typeof(Columns),
                typeof(SheetData),
                typeof(SheetCalculationProperties),
                typeof(SheetProtection),
                typeof(ProtectedRanges),
                typeof(Scenarios),
                typeof(AutoFilter),
                typeof(SortState),
                typeof(DataConsolidate),
                typeof(CustomSheetViews),
                typeof(MergeCells),
                typeof(PhoneticProperties),
                typeof(ConditionalFormatting),
                typeof(DataValidations),
                typeof(Hyperlinks),
                typeof(PrintOptions),
                typeof(PageMargins),
                typeof(PageSetup),
                typeof(HeaderFooter),
                typeof(RowBreaks),
                typeof(ColumnBreaks),
                typeof(CustomProperties),
                typeof(CellWatches),
                typeof(IgnoredErrors),
                typeof(Drawing),
                typeof(LegacyDrawing),
                typeof(LegacyDrawingHeaderFooter),
                typeof(DrawingHeaderFooter),
                typeof(Picture),
                typeof(OleObjects),
                typeof(Controls),
                typeof(WebPublishItems),
                typeof(TableParts),
                typeof(WorksheetExtensionList),
            };
    }
}
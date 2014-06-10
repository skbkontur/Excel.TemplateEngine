using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Caches;
using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Primitives
{
    internal class ExcelSpreadsheet : IExcelSpreadsheet
    {
        public ExcelSpreadsheet(WorksheetPart worksheetPart, IExcelDocumentStyle documentStyle, IExcelSharedStrings excelSharedStrings, IExcelDocument document)
        {
            worksheet = worksheetPart.Worksheet;
            this.documentStyle = documentStyle;
            this.excelSharedStrings = excelSharedStrings;
            this.document = document;
        }

        public void MergeCells(int fromRow, int fromCol, int toRow, int toCol)
        {
            var mergeCells = worksheet.GetFirstChild<MergeCells>() ?? CreateMergeCellsWorksheetPart();
            mergeCells.AppendChild(new MergeCell
                {
                    Reference = string.Format("{0}{1}:{2}{3}", IndexHelpers.ToColumnName(fromCol), fromRow, IndexHelpers.ToColumnName(toCol), toRow)
                });
        }

        public void CreateHyperlink(int row, int col, int toSpreadsheet, int toRow, int toCol)
        {
            var hyperlinks = worksheet.GetFirstChild<Hyperlinks>() ?? CreateHyperlinksWorksheetPart();
            var spreadSheetName = document.GetSpreadsheetName(toSpreadsheet);
            hyperlinks.AppendChild(new Hyperlink
                {
                    Reference = IndexHelpers.ToCellName(row, col),
                    Location = string.Format("{0}!{1}", spreadSheetName, IndexHelpers.ToCellName(toRow, toCol))
                });
        }

        public IExcelRow CreateRow(int rowIndex)
        {
            var row = new Row
                {
                    RowIndex = new UInt32Value((uint)rowIndex)
                };
            worksheet.GetFirstChild<SheetData>().AppendChild(row);
            return new ExcelRow(row, documentStyle, excelSharedStrings);
        }

        private MergeCells CreateMergeCellsWorksheetPart()
        {
            // Имеет принципиальное значение, куда именно вставлять элемент MergeCells
            // см. http://msdn.microsoft.com/en-us/library/office/cc880096(v=office.15).aspx

            var mergeCells = new MergeCells();
            if(worksheet.Elements<CustomSheetView>().Any())
                worksheet.InsertAfter(mergeCells, worksheet.Elements<CustomSheetView>().First());
            else if(worksheet.Elements<DataConsolidate>().Any())
                worksheet.InsertAfter(mergeCells, worksheet.Elements<DataConsolidate>().First());
            else if(worksheet.Elements<SortState>().Any())
                worksheet.InsertAfter(mergeCells, worksheet.Elements<SortState>().First());
            else if(worksheet.Elements<AutoFilter>().Any())
                worksheet.InsertAfter(mergeCells, worksheet.Elements<AutoFilter>().First());
            else if(worksheet.Elements<Scenarios>().Any())
                worksheet.InsertAfter(mergeCells, worksheet.Elements<Scenarios>().First());
            else if(worksheet.Elements<ProtectedRanges>().Any())
                worksheet.InsertAfter(mergeCells, worksheet.Elements<ProtectedRanges>().First());
            else if(worksheet.Elements<SheetProtection>().Any())
                worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetProtection>().First());
            else if(worksheet.Elements<SheetCalculationProperties>().Any())
                worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetCalculationProperties>().First());
            else
                worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
            return mergeCells;
        }

        private Hyperlinks CreateHyperlinksWorksheetPart()
        {
            var hyperlinks = new Hyperlinks();
            if(worksheet.Elements<DataValidations>().Any())
                worksheet.InsertAfter(hyperlinks, worksheet.Elements<DataValidations>().First());
            else if(worksheet.Elements<ConditionalFormatting>().Any())
                worksheet.InsertAfter(hyperlinks, worksheet.Elements<ConditionalFormatting>().First());
            else if(worksheet.Elements<PhoneticProperties>().Any())
                worksheet.InsertAfter(hyperlinks, worksheet.Elements<PhoneticProperties>().First());
            else if(worksheet.Elements<MergeCells>().Any())
                worksheet.InsertAfter(hyperlinks, worksheet.Elements<MergeCells>().First());
            else if(worksheet.Elements<CustomSheetView>().Any())
                worksheet.InsertAfter(hyperlinks, worksheet.Elements<CustomSheetView>().First());
            else if(worksheet.Elements<DataConsolidate>().Any())
                worksheet.InsertAfter(hyperlinks, worksheet.Elements<DataConsolidate>().First());
            else if(worksheet.Elements<SortState>().Any())
                worksheet.InsertAfter(hyperlinks, worksheet.Elements<SortState>().First());
            else if(worksheet.Elements<AutoFilter>().Any())
                worksheet.InsertAfter(hyperlinks, worksheet.Elements<AutoFilter>().First());
            else if(worksheet.Elements<Scenarios>().Any())
                worksheet.InsertAfter(hyperlinks, worksheet.Elements<Scenarios>().First());
            else if(worksheet.Elements<ProtectedRanges>().Any())
                worksheet.InsertAfter(hyperlinks, worksheet.Elements<ProtectedRanges>().First());
            else if(worksheet.Elements<SheetProtection>().Any())
                worksheet.InsertAfter(hyperlinks, worksheet.Elements<SheetProtection>().First());
            else if(worksheet.Elements<SheetCalculationProperties>().Any())
                worksheet.InsertAfter(hyperlinks, worksheet.Elements<SheetCalculationProperties>().First());
            else
                worksheet.InsertAfter(hyperlinks, worksheet.Elements<SheetData>().First());
            return hyperlinks;
        }

        private readonly IExcelDocumentStyle documentStyle;
        private readonly IExcelSharedStrings excelSharedStrings;
        private readonly Worksheet worksheet;
        private readonly IExcelDocument document;
    }
}
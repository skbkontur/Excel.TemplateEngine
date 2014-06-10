using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SKBKontur.Catalogue.ExcelFileGenerator
{
    public interface IExcelSpreadsheet
    {
        void MergeCells(int fromRow, int fromCol, int toRow, int toCol);
        IExcelRow CreateRow(int rowIndex);
    }

    internal class ExcelSpreadsheet : IExcelSpreadsheet
    {
        public ExcelSpreadsheet(WorksheetPart worksheetPart, IExcelDocumentStyle documentStyle, IExcelSharedStrings excelSharedStrings)
        {
            this.worksheetPart = worksheetPart;
            this.documentStyle = documentStyle;
            this.excelSharedStrings = excelSharedStrings;
        }

        public void MergeCells(int fromRow, int fromCol, int toRow, int toCol)
        {
            var worksheet = worksheetPart.Worksheet;
            var mergeCells = worksheet.GetFirstChild<MergeCells>() ?? CreateMergeCellsWorksheetPart(worksheet);
            mergeCells.AppendChild(new MergeCell
                {
                    Reference = string.Format("{0}{1}:{2}{3}", IndexHelpers.ToColumnName(fromCol), fromRow, IndexHelpers.ToColumnName(toCol), toRow)
                });
        }

        public IExcelRow CreateRow(int rowIndex)
        {
            var row = new Row
                {
                    RowIndex = new UInt32Value((uint)rowIndex)
                };
            worksheetPart.Worksheet.GetFirstChild<SheetData>().AppendChild(row);
            return new ExcelRow(row, documentStyle, excelSharedStrings);
        }

        private static MergeCells CreateMergeCellsWorksheetPart(Worksheet worksheet)
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

        private readonly WorksheetPart worksheetPart;
        private readonly IExcelDocumentStyle documentStyle;
        private readonly IExcelSharedStrings excelSharedStrings;
    }
}
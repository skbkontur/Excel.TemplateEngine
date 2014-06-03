using System;
using System.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SKBKontur.Catalogue.ExcelFileGenerator
{
    public interface IExcelSpreadsheet
    {
        void MergeCells(int fromRow, int fromCol, int toRow, int toCol);
    }

    internal class ExcelSpreadsheet : IExcelSpreadsheet
    {
        public ExcelSpreadsheet(WorksheetPart worksheetPart)
        {
            this.worksheetPart = worksheetPart;
        }

        public void MergeCells(int fromRow, int fromCol, int toRow, int toCol)
        {
            var worksheet = worksheetPart.Worksheet;
            var mergeCells = worksheet.GetFirstChild<MergeCells>() ?? CreateMergeCellsWorksheetPart(worksheet);
            mergeCells.AppendChild(new MergeCell
                {
                    Reference = string.Format("{0}{1}:{2}{3}", ToRowIndex(fromCol), fromRow, ToRowIndex(toCol), toRow)
                });
            worksheet.Save();
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

        private string ToRowIndex(int fromCol)
        {
            if(fromCol > 0 && fromCol <= 26)
                return "" + (char)(fromCol + 'A' - 1);
            throw new Exception();
        }

        private readonly WorksheetPart worksheetPart;
    }
}
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using SKBKontur.Catalogue.ExcelFileGenerator.Helpers;
using SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Caches;
using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Primitives
{
    public class ExcelDocument : IExcelDocument
    {
        public ExcelDocument(byte[] template)
        {
            worksheetsCache = new Dictionary<string, WorksheetPart>();

            documentMemoryStream = new MemoryStream();
            documentMemoryStream.Write(template, 0, template.Length);
            spreadsheetDocument = SpreadsheetDocument.Open(documentMemoryStream, true);

            documentStyle = new ExcelDocumentStyle(spreadsheetDocument.GetOrCreateSpreadsheetStyles());
            excelSharedStrings = new ExcelSharedStrings(spreadsheetDocument.GetOrCreateSpreadsheetSharedStrings());
            spreadsheetDisposed = false;
        }

        private void ThrowIfSpreadsheetDisposed()
        {
            if(spreadsheetDisposed)
                throw new ObjectDisposedException(spreadsheetDocument.GetType().Name);
        }

        public void Dispose()
        {
            if(!spreadsheetDisposed)
                spreadsheetDocument.Dispose();
            documentMemoryStream.Dispose();
            spreadsheetDisposed = true;
        }

        public byte[] CloseAndGetDocumentBytes() //Закрывает документ OpenXml и делает все методы недоступными
        {
            ThrowIfSpreadsheetDisposed();
            spreadsheetDocument.Dispose();
            spreadsheetDisposed = true;
            return documentMemoryStream.ToArray();
        }

        public int GetWorksheetCount()
        {
            ThrowIfSpreadsheetDisposed();
            return spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Count();
        }

        public IExcelWorksheet GetWorksheet(int index)
        {
            ThrowIfSpreadsheetDisposed();
            var sheetId = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().ElementAt(index).Id.Value;
            WorksheetPart worksheetPart;
            if(!worksheetsCache.TryGetValue(sheetId, out worksheetPart))
            {
                worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(sheetId);
                worksheetsCache.Add(sheetId, worksheetPart);
            }
            return new ExcelWorksheet(worksheetPart, documentStyle, excelSharedStrings);
        }

        public void RenameWorksheet(int index, string name)
        {
            ThrowIfSpreadsheetDisposed();
            spreadsheetDocument.WorkbookPart.Workbook.Sheets.Elements<Sheet>().ElementAt(index).Name = name;
        }

        public IExcelWorksheet AddWorksheet(string worksheetName)
        {
            ThrowIfSpreadsheetDisposed();
            var worksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            var sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>();

            var sheetId = 1u;
            if(sheets.Elements<Sheet>().Any())
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;

            var sheet = new Sheet
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = sheetId,
                    Name = worksheetName
                };

            sheets.AppendChild(sheet);

            return new ExcelWorksheet(spreadsheetDocument.WorkbookPart.WorksheetParts.Last(), documentStyle, excelSharedStrings);
        }

        public override string ToString()
        {
            var result = "";
            for(var worksheetId = 0; worksheetId < GetWorksheetCount(); ++worksheetId)
            {
                result = string.Format("Worksheet #{0}:\n\n", worksheetId);
                var worksheet = GetWorksheet(worksheetId);
                result += string.Join("", worksheet
                                              .SearchCellsByText("")
                                              .Select(cell => string.Format("{0}:{1}\n{2}\n", cell.GetCellIndex().CellReference, cell.GetStringValue(), cell.GetStyle())));
                result += "\nMerged cells info:\n";
                result += string.Join("", worksheet.MergedCells
                                                   .Select(mergedCells => string.Format("{0}:{1}\n", mergedCells.Item1.CellReference, mergedCells.Item2.CellReference))
                                                   .OrderBy(s => s));
                result += "\nColumns info:\n";
                result += string.Join("", worksheet.Columns.Select(column => string.Format("Column index = {0} Column width = {1:0.0}\n", column.Index.ToString(), column.Width)));
                result += "\n\n";
            }
            return result;
        }

        private readonly IDictionary<string, WorksheetPart> worksheetsCache;
        private readonly MemoryStream documentMemoryStream;
        private readonly SpreadsheetDocument spreadsheetDocument;
        private readonly IExcelDocumentStyle documentStyle;
        private readonly IExcelSharedStrings excelSharedStrings;
        private bool spreadsheetDisposed;
    }
}
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Caches;
using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;
using SKBKontur.Catalogue.Objects;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Primitives
{
    public class ExcelDocument : IExcelDocument, IExcelDocumentMeta
    {
        public ExcelDocument(byte[] template)
        {
            worksheetsCache = new Dictionary<string, WorksheetPart>();

            documentMemoryStream = new MemoryStream();
            documentMemoryStream.Write(template, 0, template.Length);
            spreadsheetDocument = SpreadsheetDocument.Open(documentMemoryStream, true);

            documentStyle = new ExcelDocumentStyle(spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet);
            excelSharedStrings = new ExcelSharedStrings(spreadsheetDocument.WorkbookPart.With(wp => wp.SharedStringTablePart).With(sstp => sstp.SharedStringTable));
            spreadsheetDisposed = false;
            excelDocumentDisposed = false;
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
            excelDocumentDisposed = true;
        }

        public byte[] GetDocumentBytes()
        {
            if(excelDocumentDisposed)
                throw new ObjectDisposedException(GetType().Name);

            if(!spreadsheetDisposed)
            {
                Flush();
                spreadsheetDocument.Dispose();
                spreadsheetDisposed = true;
            }

            return documentMemoryStream.ToArray();
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
            return new ExcelWorksheet(worksheetPart, documentStyle, excelSharedStrings, this);
        }

        public void DeleteWorksheet(int index)
        {
            ThrowIfSpreadsheetDisposed();
            var workbookPart = spreadsheetDocument.WorkbookPart;
            var sheet = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().ElementAt(index);
            var sheetToDelete = sheet.Name.Value;
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            worksheetsCache.Remove(sheet.Id);
            sheet.Remove();
            workbookPart.DeletePart(worksheetPart);
            var definedNames = workbookPart.Workbook.Descendants<DefinedNames>().FirstOrDefault();
            if(definedNames != null)
            {
                var defNamesToDelete = definedNames.Cast<DefinedName>().Where(item => item.Text.Contains(sheetToDelete + "!")).ToList();
                foreach(var item in defNamesToDelete)
                    item.Remove();
            }
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

            spreadsheetDocument.WorkbookPart.Workbook.Save();

            return new ExcelWorksheet(spreadsheetDocument.WorkbookPart.WorksheetParts.Last(), documentStyle, excelSharedStrings, this);
        }

        public void SetPivotTableSource(int tableIndex, ExcelCellIndex upperLeft, ExcelCellIndex lowerRight)
        {
            ThrowIfSpreadsheetDisposed();
            var worksheetSource = spreadsheetDocument.WorkbookPart.PivotTableCacheDefinitionParts.ElementAt(tableIndex).PivotCacheDefinition.CacheSource.GetFirstChild<WorksheetSource>();
            worksheetSource.Reference = string.Format("{0}:{1}", upperLeft.CellReference, lowerRight.CellReference);
        }

        public string GetWorksheetName(int index)
        {
            ThrowIfSpreadsheetDisposed();
            return ((Sheet)spreadsheetDocument.WorkbookPart.Workbook.Sheets.ChildElements[index]).Name;
        }

        private void Flush()
        {
            ThrowIfSpreadsheetDisposed();
            // Saving document parts in OpenXml is not thread-safe. Detailed info at http://support2.microsoft.com/kb/951731
            lock(lockObject)
            {
                spreadsheetDocument.WorkbookPart.Workbook.RemoveAllChildren<DefinedNames>();
                foreach(var worksheetPart in worksheetsCache.Values)
                    worksheetPart.Worksheet.Save();
                documentStyle.Save();

                if(excelSharedStrings != null)
                    excelSharedStrings.Save();
                foreach(var pivotTableCacheDefinitionPart in spreadsheetDocument.WorkbookPart.PivotTableCacheDefinitionParts)
                    pivotTableCacheDefinitionPart.PivotCacheDefinition.Save();
                spreadsheetDocument.WorkbookPart.Workbook.Save();
            }
        }

        private readonly IDictionary<string, WorksheetPart> worksheetsCache;
        private readonly MemoryStream documentMemoryStream;
        private readonly SpreadsheetDocument spreadsheetDocument;
        private readonly IExcelDocumentStyle documentStyle;
        private readonly IExcelSharedStrings excelSharedStrings;
        private bool spreadsheetDisposed;
        private bool excelDocumentDisposed;

        private static readonly object lockObject = new object();
    }
}
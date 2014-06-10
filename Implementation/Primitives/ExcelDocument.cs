using System.Collections.Generic;
using System.IO;
using System.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Caches;
using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Primitives
{
    internal class ExcelDocument : IExcelDocument
    {
        public ExcelDocument(byte[] template)
        {
            worksheetsCache = new Dictionary<string, WorksheetPart>();

            documentMemoryStream = new MemoryStream();
            documentMemoryStream.Write(template, 0, template.Length);
            spreadsheetDocument = SpreadsheetDocument.Open(documentMemoryStream, true);

            documentStyle = new ExcelDocumentStyle(spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet);
            excelSharedStrings = new ExcelSharedStrings(spreadsheetDocument.WorkbookPart.SharedStringTablePart.SharedStringTable);
        }

        public void Dispose()
        {
            spreadsheetDocument.Dispose();
            documentMemoryStream.Dispose();
        }

        public byte[] GetDocumentBytes()
        {
            Flush();
            return documentMemoryStream.ToArray();
        }

        public IExcelSpreadsheet GetSpreadsheet(int index)
        {
            var sheetId = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().ElementAt(index).Id.Value;
            WorksheetPart worksheetPart;
            if(!worksheetsCache.TryGetValue(sheetId, out worksheetPart))
            {
                worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(sheetId);
                worksheetsCache.Add(sheetId, worksheetPart);
            }
            return new ExcelSpreadsheet(worksheetPart, documentStyle, excelSharedStrings, this);
        }

        public string GetSpreadsheetName(int index)
        {
            return ((Sheet)spreadsheetDocument.WorkbookPart.Workbook.Sheets.ChildElements[index]).Name;
        }

        private void Flush()
        {
            foreach(var worksheetPart in worksheetsCache.Values)
                worksheetPart.Worksheet.Save();
            documentStyle.Save();
            excelSharedStrings.Save();
        }

        private readonly IDictionary<string, WorksheetPart> worksheetsCache;
        private readonly MemoryStream documentMemoryStream;
        private readonly SpreadsheetDocument spreadsheetDocument;
        private readonly IExcelDocumentStyle documentStyle;
        private readonly IExcelSharedStrings excelSharedStrings;
    }
}
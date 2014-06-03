using System;
using System.IO;
using System.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SKBKontur.Catalogue.ExcelFileGenerator
{
    public interface IExcelDocument : IDisposable
    {
        byte[] GetDocumentBytes();
        IExcelSpreadsheet GetSpreadsheet(int index);
    }

    internal class ExcelDocument : IExcelDocument
    {
        public ExcelDocument(byte[] template)
        {
            documentMemoryStream = new MemoryStream();
            documentMemoryStream.Write(template, 0, template.Length);
            spreadsheetDocument = SpreadsheetDocument.Open(documentMemoryStream, true);
        }

        public void Dispose()
        {
            spreadsheetDocument.Dispose();
            documentMemoryStream.Dispose();
        }

        public byte[] GetDocumentBytes()
        {
            return documentMemoryStream.ToArray();
        }

        public IExcelSpreadsheet GetSpreadsheet(int index)
        {
            var sheet = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().ElementAt(index);
            var worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(sheet.Id.Value);
            return new ExcelSpreadsheet(worksheetPart);
        }

        private readonly MemoryStream documentMemoryStream;
        private readonly SpreadsheetDocument spreadsheetDocument;
    }
}
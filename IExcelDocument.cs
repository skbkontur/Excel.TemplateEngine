using System;
using System.IO;

using DocumentFormat.OpenXml.Packaging;

namespace SKBKontur.Catalogue.ExcelFileGenerator
{
    public interface IExcelDocument
    {
        byte[] GetDocumentBytes();
    }

    internal class ExcelDocument : IExcelDocument, IDisposable
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

        private readonly MemoryStream documentMemoryStream;
        private readonly SpreadsheetDocument spreadsheetDocument;
    }
}
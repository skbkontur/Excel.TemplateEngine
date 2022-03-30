using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using JetBrains.Annotations;

using SkbKontur.Excel.TemplateEngine.FileGenerating.Helpers;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.LazyParse
{
    /// <summary>
    ///     Reads current or following row. Can't read previous ones.
    /// </summary>
    public class LazyTableReader : IDisposable
    {
        public LazyTableReader([NotNull] byte[] excelData)
        {
            var stream = new MemoryStream();
            stream.Write(excelData, 0, excelData.Length);
            InitializeReader(stream);
        }

        /// <param name="stream">Stream will be disposed with the reader.</param>
        public LazyTableReader([NotNull] Stream stream)
        {
            InitializeReader(stream);
        }

        private void InitializeReader([NotNull] Stream stream)
        {
            this.stream = stream;
            spreadsheetDocument = SpreadsheetDocument.Open(stream, false);

            var sharedStringTable = spreadsheetDocument.GetOrCreateSpreadsheetSharedStrings();
            var sharedStringsArray = sharedStringTable.ChildElements
                                                      .Select(x => x.InnerText)
                                                      .ToArray();
            sharedStrings = Array.AsReadOnly(sharedStringsArray);

            var firstSheetId = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().ElementAt(0).Id.Value;
            var worksheet = (WorksheetPart)spreadsheetDocument.WorkbookPart?.GetPartById(firstSheetId!);
            if (worksheet == null)
                throw new ArgumentException("Incoming Excel document has no worksheets.");

            var sheetData = worksheet.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null)
                throw new ArgumentException("Incoming Excel document has no SheetData.");

            reader = OpenXmlReader.Create(sheetData);
        }

        [NotNull]
        private LazyRowReader LoadCurrentRowReader()
        {
            var row = (Row)reader.LoadCurrentElement();
            return new LazyRowReader(row!, sharedStrings);
        }

        [CanBeNull]
        public LazyRowReader TryReadRow(int rowIndex)
        {
            if (currentRowReader != null)
            {
                if (rowIndex < currentRowReader.RowIndex)
                    return null;

                if (rowIndex == currentRowReader.RowIndex)
                    return currentRowReader;
            }

            while (reader.Read())
            {
                if (reader.ElementType != typeof(Row))
                    continue;

                currentRowReader = LoadCurrentRowReader();

                if (rowIndex > currentRowReader.RowIndex)
                    continue;

                if (rowIndex < currentRowReader.RowIndex)
                    return null;

                if (rowIndex == currentRowReader.RowIndex)
                    return currentRowReader;
            }

            return null;
        }

        public void Dispose()
        {
            stream?.Dispose();
            spreadsheetDocument.Dispose();
            reader.Dispose();
            currentRowReader?.Dispose();
        }

        [NotNull]
        private OpenXmlReader reader;

        [NotNull]
        private IReadOnlyList<string> sharedStrings;

        [CanBeNull]
        private LazyRowReader currentRowReader;

        [CanBeNull]
        private Stream stream;

        [NotNull]
        private SpreadsheetDocument spreadsheetDocument;
    }
}
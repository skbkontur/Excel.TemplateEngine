using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using JetBrains.Annotations;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.ParseCollection.Parsers.Implementations
{
    [UsedImplicitly]
    class LazyTableReader : IEnumerable<Row>
    {
        public LazyTableReader([NotNull] byte[] excelData)
        {
            var stream = new MemoryStream();
            stream.Write(excelData, 0, excelData.Length);
            var spreadsheetDocument = SpreadsheetDocument.Open(stream, true);
            var worksheet = spreadsheetDocument.WorkbookPart?.WorksheetParts.FirstOrDefault();
            if (worksheet == null)
                throw new ArgumentException("Incoming Excel document has no worksheets.");
            var sheetData = worksheet.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null)
                throw new ArgumentException("Incoming Excel document has no SheetData.");
            reader = OpenXmlReader.Create(sheetData);
        }

        [NotNull]
        public Row GetCurrentRow()
        {
            if (reader.ElementType != typeof(Row))
                throw new Exception($"Current element is not {nameof(Row)}");

            return ((Row)reader.LoadCurrentElement())!;
        }

        [NotNull]
        public Row GetNextRow([CanBeNull] int? rowIndex = null)
        {
            while (reader.Read())
            {
                if (reader.ElementType != typeof(Row))
                    continue;

                var row = (Row)reader.LoadCurrentElement();
                if (rowIndex.HasValue && rowIndex.Value < row!.RowIndex)
                    throw new ArgumentException($"{nameof(rowIndex)} is less than current row index. Can't read previous rows.");

                if (rowIndex.HasValue && rowIndex.Value > row!.RowIndex)
                    continue;

                return row!;
            }
            throw new IndexOutOfRangeException(nameof(rowIndex));
        }

        public IEnumerator<Row> GetEnumerator()
        {
            while (reader.Read())
            {
                if (reader.ElementType != typeof(Row))
                    continue;

                yield return (Row)reader.LoadCurrentElement();
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        private readonly OpenXmlReader reader;
    }
}
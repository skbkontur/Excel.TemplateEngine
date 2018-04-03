using System;
using System.Collections.Concurrent;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using JetBrains.Annotations;

using MoreLinq;

using SKBKontur.Catalogue.ExcelFileGenerator.Helpers;
using SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Caches;
using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;
using SKBKontur.Catalogue.Objects;
using SKBKontur.Catalogue.ServiceLib.Logging;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Primitives
{
    public class ExcelDocument : IExcelDocument
    {
        public ExcelDocument([NotNull] byte[] template)
        {
            worksheetsCache = new ConcurrentDictionary<string, WorksheetPart>();

            documentMemoryStream = new MemoryStream();
            documentMemoryStream.Write(template, 0, template.Length);
            spreadsheetDocument = SpreadsheetDocument.Open(documentMemoryStream, true);

            documentStyle = new ExcelDocumentStyle(spreadsheetDocument.GetOrCreateSpreadsheetStyles(), spreadsheetDocument.WorkbookPart.ThemePart.Theme);
            excelSharedStrings = new ExcelSharedStrings(spreadsheetDocument.GetOrCreateSpreadsheetSharedStrings());
            spreadsheetDisposed = false;

            SetDefaultCreatorAndEditor();
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

        [NotNull]
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

        [CanBeNull]
        public IExcelWorksheet FindWorksheet([NotNull] string name)
        {
            ThrowIfSpreadsheetDisposed();
            var sheet = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().FirstOrDefault(x => x?.Name?.Value == name);
            if(sheet == null)
                return null;
            return GetWorksheetById(sheet.Id.Value);
        }

        [CanBeNull]
        public IExcelWorksheet TryGetWorksheet(int index)
        {
            ThrowIfSpreadsheetDisposed();
            try
            {
                var sheetId = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().ElementAt(index).Id.Value;
                return GetWorksheetById(sheetId);
            }
            catch(Exception ex)
            {
                Log.For(this).Error($"An error occurred while getting of an excel worksheet: {ex}");
                return null;
            }
        }

        [NotNull]
        public IExcelWorksheet GetWorksheet(int index)
        {
            ThrowIfSpreadsheetDisposed();
            var result = TryGetWorksheet(index);
            return result ?? throw new InvalidProgramStateException("An error occurred while getting of an excel worksheet");
        }

        [NotNull]
        public string GetWorksheetName(int index)
        {
            ThrowIfSpreadsheetDisposed();
            return spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().ElementAt(index).Name;
        }

        public void CopyVbaInfoFrom([NotNull] IExcelDocument excelDocument)
        {
            ThrowIfSpreadsheetDisposed();
            var part = ((ExcelDocument)excelDocument).spreadsheetDocument.WorkbookPart.VbaProjectPart;
            if(part == null)
                return;
            spreadsheetDocument.WorkbookPart.AddPart(part);
        }

        [CanBeNull]
        public string GetDescription()
        {
            ThrowIfSpreadsheetDisposed();
            return spreadsheetDocument.PackageProperties.Description;
        }

        public void AddDescription([NotNull] string text)
        {
            ThrowIfSpreadsheetDisposed();
            spreadsheetDocument.PackageProperties.Description = text;
        }

        private void SetDefaultCreatorAndEditor()
        {
            ThrowIfSpreadsheetDisposed();
            spreadsheetDocument.PackageProperties.Creator = "Контур.EDI";
            spreadsheetDocument.PackageProperties.Created = XmlConvert.ToDateTime("2014-01-01T00:00:00Z", XmlDateTimeSerializationMode.RoundtripKind);
            spreadsheetDocument.PackageProperties.Modified = XmlConvert.ToDateTime("2014-01-01T00:00:00Z", XmlDateTimeSerializationMode.RoundtripKind);
            spreadsheetDocument.PackageProperties.LastModifiedBy = "Контур.EDI";
        }

        [NotNull]
        private IExcelWorksheet GetWorksheetById([NotNull] string sheetId)
        {
            ThrowIfSpreadsheetDisposed();
            var worksheetPart = worksheetsCache.GetOrAdd(sheetId, x => (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(x));
            return new ExcelWorksheet(this, worksheetPart, documentStyle, excelSharedStrings);
        }

        public void RenameWorksheet(int index, [NotNull] string name)
        {
            ThrowIfSpreadsheetDisposed();
            AssertWorksheetNameValid(name);
            spreadsheetDocument.WorkbookPart.Workbook.Sheets.Elements<Sheet>().ElementAt(index).Name = name;
        }

        [NotNull]
        public IExcelWorksheet AddWorksheet([NotNull] string worksheetName)
        {
            ThrowIfSpreadsheetDisposed();
            AssertWorksheetNameValid(worksheetName);

            if(FindWorksheet(worksheetName) != null)
                throw new InvalidProgramStateException($"Sheet with name {worksheetName} already exists");

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
            return new ExcelWorksheet(this, spreadsheetDocument.WorkbookPart.WorksheetParts.Last(), documentStyle, excelSharedStrings);
        }

        private static void AssertWorksheetNameValid([NotNull] string worksheetName)
        {
            if(worksheetName.Length > 31)
                throw new InvalidProgramStateException($"Worksheet name ('{worksheetName}') is too long (allowed <=31 symbols, current - {worksheetName.Length})");
        }

        public override string ToString()
        {
            var stringBuilder = new StringBuilder();
            for(var worksheetId = 0; worksheetId < GetWorksheetCount(); ++worksheetId)
            {
                var worksheet = GetWorksheet(worksheetId);

                stringBuilder.AppendFormat("Worksheet #{0}:\n\n", worksheetId);
                worksheet.SearchCellsByText("")
                         .ForEach(cell => stringBuilder.AppendFormat("{0}:{1}\n{2}\n", cell.GetCellIndex().CellReference, cell.GetStringValue(), cell.GetStyle()));

                stringBuilder.Append("\nMerged cells info:\n");
                worksheet.MergedCells
                         .Select(mergedCells => string.Format("{0}:{1}\n", mergedCells.Item1.CellReference, mergedCells.Item2.CellReference))
                         .OrderBy(s => s)
                         .ForEach(s => stringBuilder.Append(s));

                stringBuilder.Append("\nColumns info:\n");
                worksheet.Columns
                         .ForEach(column => stringBuilder.AppendFormat("Column index = {0} Column width = {1:0.0}\n", column.Index.ToString(), column.Width));
                stringBuilder.Append("\n\n");
            }
            return stringBuilder.ToString();
        }

        private readonly ConcurrentDictionary<string, WorksheetPart> worksheetsCache;
        private readonly MemoryStream documentMemoryStream;
        private readonly SpreadsheetDocument spreadsheetDocument;
        private readonly IExcelDocumentStyle documentStyle;
        private readonly IExcelSharedStrings excelSharedStrings;
        private bool spreadsheetDisposed;
    }
}
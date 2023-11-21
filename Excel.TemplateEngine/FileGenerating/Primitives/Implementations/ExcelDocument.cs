using System;
using System.Collections.Concurrent;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.VariantTypes;

using JetBrains.Annotations;

using MoreLinq;

using SkbKontur.Excel.TemplateEngine.FileGenerating.Caches;
using SkbKontur.Excel.TemplateEngine.FileGenerating.Caches.Implementations;
using SkbKontur.Excel.TemplateEngine.FileGenerating.Helpers;

using Vostok.Logging.Abstractions;

namespace SkbKontur.Excel.TemplateEngine.FileGenerating.Primitives.Implementations
{
    internal class ExcelDocument : IExcelDocument
    {
        public ExcelDocument([NotNull] byte[] template, [NotNull] ILog logger)
        {
            this.logger = logger;
            worksheetsCache = new ConcurrentDictionary<string, WorksheetPart>();

            documentMemoryStream = new MemoryStream();
            documentMemoryStream.Write(template, 0, template.Length);
            spreadsheetDocument = SpreadsheetDocument.Open(documentMemoryStream, true);

            var theme = GetEmptyTheme();
            documentStyle = new ExcelDocumentStyle(spreadsheetDocument.GetOrCreateSpreadsheetStyles(), spreadsheetDocument.WorkbookPart?.ThemePart?.Theme ?? theme, this.logger);
            excelSharedStrings = new ExcelSharedStrings(spreadsheetDocument.GetOrCreateSpreadsheetSharedStrings());
            spreadsheetDisposed = false;

            SetDefaultCreatorAndEditor();
        }

        private static Theme GetEmptyTheme()
        {
            var theme = new Theme
                {
                    ThemeElements = new ThemeElements
                        {
                            ColorScheme = new ColorScheme()
                        }
                };
            return theme;
        }

        private void ThrowIfSpreadsheetDisposed()
        {
            if (spreadsheetDisposed)
                throw new ObjectDisposedException(spreadsheetDocument.GetType().Name);
        }

        public void Dispose()
        {
            if (!spreadsheetDisposed)
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
            if (sheet == null)
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
            catch (Exception ex)
            {
                logger.Error($"An error occurred while getting of an excel worksheet: {ex}");
                return null;
            }
        }

        [NotNull]
        public IExcelWorksheet GetWorksheet(int index)
        {
            ThrowIfSpreadsheetDisposed();
            var result = TryGetWorksheet(index);
            return result ?? throw new InvalidOperationException("An error occurred while getting of an excel worksheet");
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
            if (part == null)
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
            return new ExcelWorksheet(this, worksheetPart, documentStyle, excelSharedStrings, logger);
        }

        public void RenameWorksheet(int index, [NotNull] string name)
        {
            ThrowIfSpreadsheetDisposed();
            AssertWorksheetNameValid(name);
            spreadsheetDocument.WorkbookPart.Workbook.Sheets.Elements<Sheet>().ElementAt(index).Name = name;
        }

        public bool TryGetCustomProperty(string key, out string value)
        {
            ThrowIfSpreadsheetDisposed();
            var property = spreadsheetDocument.CustomFilePropertiesPart?.GetProperty(key);
            value = property?.InnerText;
            return value != null;
        }

        public void SetCustomProperty(string key, string value)
        {
            // https://learn.microsoft.com/en-us/office/open-xml/how-to-set-a-custom-property-in-a-word-processing-document
            const string customPropertyFormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
            var customProps = spreadsheetDocument.CustomFilePropertiesPart
                              ?? spreadsheetDocument.AddCustomFilePropertiesPart();
            // ReSharper disable once ConstantNullCoalescingCondition
            customProps.Properties ??= new Properties();

            var existentProperty = customProps.GetProperty(key);
            existentProperty?.Remove();
            customProps.Properties.AppendChild(new CustomDocumentProperty
                {
                    VTLPWSTR = new VTLPWSTR(value),
                    FormatId = customPropertyFormatId,
                    Name = key
                });
            var pid = 2;
            foreach (var item in customProps.Properties.OfType<CustomDocumentProperty>())
            {
                item.PropertyId = pid++;
            }
            customProps.Properties.Save();
        }

        [NotNull]
        public IExcelWorksheet AddWorksheet([NotNull] string worksheetName)
        {
            ThrowIfSpreadsheetDisposed();
            AssertWorksheetNameValid(worksheetName);

            if (FindWorksheet(worksheetName) != null)
                throw new InvalidOperationException($"Sheet with name {worksheetName} already exists");

            var worksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            var sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>();

            var sheetId = 1u;
            if (sheets.Elements<Sheet>().Any())
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;

            var sheet = new Sheet
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = sheetId,
                    Name = worksheetName
                };

            sheets.AppendChild(sheet);
            return new ExcelWorksheet(this, spreadsheetDocument.WorkbookPart.WorksheetParts.Last(), documentStyle, excelSharedStrings, logger);
        }

        private static void AssertWorksheetNameValid([NotNull] string worksheetName)
        {
            if (worksheetName.Length > 31)
                throw new InvalidOperationException($"Worksheet name ('{worksheetName}') is too long (allowed <=31 symbols, current - {worksheetName.Length})");
        }

        public override string ToString()
        {
            var stringBuilder = new StringBuilder();
            for (var worksheetId = 0; worksheetId < GetWorksheetCount(); ++worksheetId)
            {
                var worksheet = GetWorksheet(worksheetId);

                stringBuilder.AppendFormat("Worksheet #{0}:\n\n", worksheetId);
                worksheet.SearchCellsByText("")
                         .ForEach(cell => stringBuilder.AppendFormat("{0}:{1}\n{2}\n", cell.GetCellIndex().CellReference, cell.GetStringValue(), cell.GetStyle()));

                stringBuilder.Append("\nMerged cells info:\n");
                worksheet.MergedCells
                         .Select(mergedCells => $"{mergedCells.Item1.CellReference}:{mergedCells.Item2.CellReference}\n")
                         .OrderBy(s => s)
                         .ForEach(s => stringBuilder.Append((string)s));

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
        private readonly ILog logger;
    }
}
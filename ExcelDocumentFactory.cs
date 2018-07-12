using System;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Primitives;
using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;
using SKBKontur.Catalogue.Objects;
using SKBKontur.Catalogue.ServiceLib.Logging;

namespace SKBKontur.Catalogue.ExcelFileGenerator
{
    public static class ExcelDocumentFactory
    {
        [CanBeNull]
        public static IExcelDocument TryCreateFromTemplate([NotNull] byte[] template)
        {
            try
            {
                return new ExcelDocument(template);
            }
            catch (Exception ex)
            {
                Log.For<ExcelDocument>().Error($"An error occurred while creating of {nameof(ExcelDocument)}: {ex}");
                return null;
            }
        }

        [NotNull]
        public static IExcelDocument CreateFromTemplate([NotNull] byte[] template)
        {
            var result = TryCreateFromTemplate(template);
            return result ?? throw new InvalidProgramStateException($"An error occurred while creating of {nameof(ExcelDocument)}");
        }

        [NotNull]
        public static IExcelDocument CreateEmpty(bool useXlsm)
        {
            if (useXlsm)
                return CreateFromTemplate(GetFileBytes("empty.xlsm"));
            return CreateFromTemplate(GetFileBytes("empty.xlsx"));
        }

        private static byte[] GetFileBytes(string filename)
        {
            return typeof(ExcelDocumentFactory).Assembly.ReadAllBytesFromResource(filename);
        }
    }
}
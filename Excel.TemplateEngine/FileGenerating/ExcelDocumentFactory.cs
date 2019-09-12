using System;
using System.IO;

using Excel.TemplateEngine.Exceptions;
using Excel.TemplateEngine.FileGenerating.Primitives;
using Excel.TemplateEngine.FileGenerating.Primitives.Implementations;

using JetBrains.Annotations;

using Vostok.Logging.Abstractions;

namespace Excel.TemplateEngine.FileGenerating
{
    public static class ExcelDocumentFactory
    {
        [CanBeNull]
        public static IExcelDocument TryCreateFromTemplate([NotNull] byte[] template, [NotNull] ILog logger)
        {
            var excelFileGeneratorLogger = logger.ForContext("ExcelFileGenerator");
            try
            {
                return new ExcelDocument(template, excelFileGeneratorLogger);
            }
            catch (Exception ex)
            {
                excelFileGeneratorLogger.Error($"An error occurred while creating of {nameof(ExcelDocument)}: {ex}");
                return null;
            }
        }

        [NotNull]
        public static IExcelDocument CreateFromTemplate([NotNull] byte[] template, [NotNull] ILog logger)
            => TryCreateFromTemplate(template, logger)
               ?? throw new ExcelTemplateEngineException($"An error occurred while creating of {nameof(ExcelDocument)}");

        [NotNull]
        public static IExcelDocument CreateEmpty(bool useXlsm, [NotNull] ILog logger)
        {
            var resourceBytes = GetResourceBytes(useXlsm ? "empty.xlsm" : "empty.xlsx");
            return CreateFromTemplate(resourceBytes, logger);
        }

        private static byte[] GetResourceBytes(string resourceName)
        {
            var assembly = typeof(ExcelDocumentFactory).Assembly;
            using (var rs = assembly.GetManifestResourceStream(resourceName))
            using (var ms = new MemoryStream())
            {
                if (rs == null)
                    throw new ExcelTemplateEngineException($"Stream is null for resource: {resourceName}");
                rs.CopyTo(ms);
                return ms.ToArray();
            }
        }
    }
}
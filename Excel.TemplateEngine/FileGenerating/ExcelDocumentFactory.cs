#nullable enable

using System;
using System.IO;
using System.Linq;

using SkbKontur.Excel.TemplateEngine.FileGenerating.Primitives;
using SkbKontur.Excel.TemplateEngine.FileGenerating.Primitives.Implementations;

using Vostok.Logging.Abstractions;

namespace SkbKontur.Excel.TemplateEngine.FileGenerating;

public static class ExcelDocumentFactory
{
    public static IExcelDocument? TryCreateFromTemplate(byte[] template, ILog logger)
    {
        var excelFileGeneratorLogger = logger.ForContext("ExcelFileGenerator");
        try
        {
            return new ExcelDocument(template, excelFileGeneratorLogger);
        }
        catch (Exception ex)
        {
            excelFileGeneratorLogger.Warn(ex, $"An error occurred while creating of {nameof(ExcelDocument)}");
            return null;
        }
    }

    public static IExcelDocument CreateFromTemplate(byte[] template, ILog logger)
        => TryCreateFromTemplate(template, logger)
           ?? throw new InvalidOperationException($"An error occurred while creating of {nameof(ExcelDocument)}");

    public static IExcelDocument CreateEmpty(bool useXlsm, ILog logger)
    {
        var resourceBytes = GetResourceBytes(useXlsm ? "empty.xlsm" : "empty.xlsx");
        return CreateFromTemplate(resourceBytes, logger);
    }

    private static byte[] GetResourceBytes(string resourceName)
    {
        var assembly = typeof(ExcelDocumentFactory).Assembly;
        var resourceFullName = assembly.GetManifestResourceNames().Single(name => name.EndsWith(resourceName));
        using (var rs = assembly.GetManifestResourceStream(resourceFullName))
        using (var ms = new MemoryStream())
        {
            if (rs == null)
                throw new InvalidOperationException($"Stream is null for resource: {resourceName}");
            rs.CopyTo(ms);
            return ms.ToArray();
        }
    }
}
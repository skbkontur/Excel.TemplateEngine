using System;
using System.IO;
using System.Linq;

using SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Primitives;
using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;

namespace SKBKontur.Catalogue.ExcelFileGenerator
{
    public static class ExcelDocumentFactory
    {
        public static IExcelDocument CreateFromTemplate(byte[] template)
        {
            return new ExcelDocument(template);
        }

        public static IExcelDocument CreateEmpty()
        {
            return CreateFromTemplate(GetFileBytes(emptyFileName));
        }

        private static byte[] GetFileBytes(string filename)
        {
            var assembly = typeof(ExcelDocumentFactory).Assembly;
            var manifestResourceName = assembly.GetManifestResourceNames().First(x => x.EndsWith(filename));
            using(var stream = assembly.GetManifestResourceStream(manifestResourceName))
            using(var ms = new MemoryStream())
            {
                if(stream == null)
                    throw new Exception(string.Format("Resource '{0}' not found", filename));
                stream.CopyTo(ms);
                return ms.ToArray();
            }
        }

        private const string emptyFileName = "empty.xlsx";
    }
}
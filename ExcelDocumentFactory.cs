using SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Primitives;
using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;
using SKBKontur.Catalogue.Objects;

namespace SKBKontur.Catalogue.ExcelFileGenerator
{
    public static class ExcelDocumentFactory
    {
        public static IExcelDocument CreateFromTemplate(byte[] template)
        {
            return new ExcelDocument(template);
        }

        public static IExcelDocument CreateEmpty(bool useXlsm)
        {
            if(useXlsm)
                return CreateFromTemplate(GetFileBytes("empty.xlsm"));
            return CreateFromTemplate(GetFileBytes("empty.xlsx"));
        }

        private static byte[] GetFileBytes(string filename)
        {
            return typeof(ExcelDocumentFactory).Assembly.ReadAllBytesFromResource(filename);
        }
    }
}
namespace SKBKontur.Catalogue.ExcelFileGenerator
{
    public static class ExcelDocumentFactory
    {
        public static IExcelDocument CreateFromTemplate(byte[] template)
        {
            return new ExcelDocument(template);
        }
    }
}
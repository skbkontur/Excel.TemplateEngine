namespace SKBKontur.Catalogue.ExcelFileGenerator
{
    public class FormattedStringValue
    {
        public FormattedStringValueBlock[] Blocks { get; set; }
    }

    public class FormattedStringValueBlock
    {
        public string Value { get; set; }
        public ExcelColor Color { get; set; }
    }
}
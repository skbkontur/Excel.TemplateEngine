namespace SKBKontur.Catalogue.ExcelObjectPrinter.Exceptions
{
    public class FormControlParsingException : BaseExcelSerializationException
    {
        public FormControlParsingException(string name) : base($"Failed to parse value from control {name}")
        {
            Name = name;
        }

        public string Name { get; set; }
    }
}
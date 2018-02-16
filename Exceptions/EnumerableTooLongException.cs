namespace SKBKontur.Catalogue.ExcelObjectPrinter.Exceptions
{
    public class EnumerableTooLongException : BaseExcelSerializationException
    {
        public EnumerableTooLongException(int limit)
            : base($"IEnumerable was longer than {limit}")
        {
            Limit = limit;
        }

        public int Limit { get; set; }
    }
}
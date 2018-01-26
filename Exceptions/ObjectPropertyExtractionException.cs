using System;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.Exceptions
{
    public class ObjectPropertyExtractionException : BaseExcelSerializationException
    {
        public ObjectPropertyExtractionException()
        {
        }

        public ObjectPropertyExtractionException(string message)
            : base(message)
        {
        }

        public ObjectPropertyExtractionException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}
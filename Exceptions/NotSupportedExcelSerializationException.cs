using System;
using System.Runtime.Serialization;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.Exceptions
{
    public class NotSupportedExcelSerializationException : BaseExcelSerializationException
    {
        public NotSupportedExcelSerializationException()
        {
        }

        public NotSupportedExcelSerializationException(string message) : base(message)
        {
        }

        public NotSupportedExcelSerializationException(string message, Exception inner) : base(message, inner)
        {
        }
    }
}
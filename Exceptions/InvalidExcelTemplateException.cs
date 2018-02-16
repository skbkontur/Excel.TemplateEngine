using System;
using System.Runtime.Serialization;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.Exceptions
{
    public class InvalidExcelTemplateException : BaseExcelSerializationException
    {
        public InvalidExcelTemplateException()
        {
        }

        public InvalidExcelTemplateException(string message)
            : base(message)
        {
        }

        public InvalidExcelTemplateException(string message, Exception inner)
            : base(message, inner)
        {
        }

        protected InvalidExcelTemplateException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }
    }
}
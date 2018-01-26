using System;
using System.Runtime.Serialization;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Exceptions
{
    public class BaseExcelException : Exception
    {
        public BaseExcelException()
        {
        }

        public BaseExcelException(string message) : base(message)
        {
        }

        public BaseExcelException(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected BaseExcelException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}
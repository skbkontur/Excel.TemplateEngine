using System;
using System.Runtime.Serialization;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Exceptions
{
    public class NotSupportedExcelDocumentException : BaseExcelException
    {
        public NotSupportedExcelDocumentException()
        {
        }

        public NotSupportedExcelDocumentException(string message)
            : base(message)
        {
        }

        public NotSupportedExcelDocumentException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        protected NotSupportedExcelDocumentException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }
    }
}
using System;
using System.Runtime.Serialization;

using JetBrains.Annotations;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Exceptions
{
    public class InvalidExcelDocumentException : BaseExcelException
    {
        public InvalidExcelDocumentException()
        {
        }

        public InvalidExcelDocumentException(string message)
            : base(message)
        {
        }

        public InvalidExcelDocumentException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        protected InvalidExcelDocumentException([NotNull] SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }
    }
}

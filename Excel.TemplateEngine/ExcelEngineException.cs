using System;

using JetBrains.Annotations;

namespace Excel.TemplateEngine
{
    public class ExcelEngineException : Exception
    {
        public ExcelEngineException([NotNull] string message)
            : base(message)
        {
        }

        public ExcelEngineException([NotNull] string message, [NotNull] Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
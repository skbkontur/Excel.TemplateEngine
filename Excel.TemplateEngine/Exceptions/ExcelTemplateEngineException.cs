using System;

using JetBrains.Annotations;

namespace Excel.TemplateEngine.Exceptions
{
    public class ExcelTemplateEngineException : Exception
    {
        public ExcelTemplateEngineException([NotNull] string message)
            : base(message)
        {
        }

        public ExcelTemplateEngineException([NotNull] string message, [NotNull] Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
using System;

using JetBrains.Annotations;

namespace SkbKontur.Excel.TemplateEngine.Exceptions
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
#nullable enable

using DocumentFormat.OpenXml.Spreadsheet;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.LazyParse
{
    public interface IFormulaEvaluator
    {
        string? TryEvaluate(
            Cell cell);
    }
}

using System.Collections.Generic;

namespace SkbKontur.Excel.TemplateEngine.FileGenerating.Primitives
{
    public interface IExcelRow
    {
        IExcelCell CreateCell(int index);
        void SetHeight(double value);
        IEnumerable<IExcelCell> Cells { get; }
    }
}
using System.Collections.Generic;

namespace Excel.TemplateEngine.FileGenerating.Interfaces
{
    public interface IExcelRow
    {
        IExcelCell CreateCell(int index);
        void SetHeight(double value);
        IEnumerable<IExcelCell> Cells { get; }
    }
}
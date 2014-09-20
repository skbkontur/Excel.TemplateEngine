using System.Collections.Generic;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Interfaces
{
    public interface IExcelRow
    {
        IExcelCell CreateCell(int index);
        void SetHeight(double value);
        IEnumerable<IExcelCell> Cells { get; }
    }
}
using System.Collections.Generic;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives.Implementations
{
    public class ExcelTablePart : ITablePart
    {
        public ExcelTablePart(IReadOnlyList<IReadOnlyList<ICell>> cells)
        {
            Cells = cells;
        }

        public IReadOnlyList<IReadOnlyList<ICell>> Cells { get; }
    }
}